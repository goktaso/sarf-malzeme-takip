Option Explicit

Private address As String
Private token As String
Private localPipe As String
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal path As LongPtr, ByVal desiredAccess As Long, ByVal shareMode As Long, ByVal security As LongPtr, ByVal disposition As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal handle As LongPtr, ByRef buffer As Any, ByVal count As Long, ByRef received As Long, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal handle As LongPtr, ByRef buffer As Any, ByVal count As Long, ByRef written As Long, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long

Public Sub ConnectLocal(ByVal folder As String)
    If Len(folder) = 0 Then Err.Raise vbObjectError + 702, , "Yerel Windows oturumu açılamadı."
    localPipe = folder
End Sub

Public Sub Connect(ByVal endpoint As String, ByVal account As String, ByVal password As String)
    Dim pattern As Object
    address = Trim$(endpoint)
    If Right$(address, 1) = "/" Then address = Left$(address, Len(address) - 1)
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.IgnoreCase = True
    pattern.Pattern = "^(http://(localhost|127[.]0[.]0[.]1):[0-9]{1,5}|https://[a-z0-9.-]+(:[0-9]{1,5})?)$"
    If Not pattern.Test(address) Then
        Err.Raise vbObjectError + 701, , "Yalnizca yerel HTTP veya HTTPS adresi kullanin."
    End If
    If LCase$(Left$(address, 17)) = "http://localhost:" Then
        address = "http://127.0.0.1:" & Mid$(address, 18)
    End If
    token = AttributeValue(Xml("POST", "/login", "account=" & SarfEncode(account) & SarfParam("password", password), "application/x-www-form-urlencoded").DocumentElement, "session")
    If Len(token) = 0 Then Err.Raise vbObjectError + 702, , "Windows oturumu acilamadi."
End Sub

Public Property Get Connected() As Boolean
    Connected = (Len(token) > 0 Or Len(localPipe) > 0)
End Property

Private Function CallService(ByVal method As String, ByVal resource As String, _
                             ByVal payload As Variant, ByVal mime As String) As Object
    Dim request As Object, root As Object, description As String, phase As String, failure As Long
    On Error GoTo TransportFailed
    phase = "HTTP istemcisi"
    If Len(localPipe) > 0 Then
        phase = "yerel istek"
        Set request = CallLocalBridge(method, resource, payload, mime)
    Else
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        phase = "bağlantı ayarları"
        request.SetTimeouts 5000, 5000, 20000, 30000
        request.Open method, address & resource, False
        request.Option(6) = False   ' Never forward authentication or session headers to a redirect.
        request.SetAutoLogonPolicy 2
        If Left$(address, 7) = "http://" Then request.SetProxy 1
        If Len(token) > 0 Then request.SetRequestHeader "X-Sarf-Session", token
        If Len(mime) > 0 Then request.SetRequestHeader "Content-Type", mime
        phase = "sunucuya gönderim"
        If IsEmpty(payload) Then request.Send Else request.Send payload
    End If
    On Error GoTo 0
    If request.Status < 200 Or request.Status >= 300 Then
        Set root = CreateObject("MSXML2.DOMDocument.6.0")
        root.async = False
        root.resolveExternals = False
        If root.LoadXML(request.ResponseText) Then
            If Not root.SelectSingleNode("/result/message") Is Nothing Then
                description = root.SelectSingleNode("/result/message").Text
            End If
        End If
        If Len(description) = 0 Then description = "Servis yaniti: HTTP " & request.Status
        Err.Raise vbObjectError + 703, , description
    End If
    Set CallService = request
    Exit Function
TransportFailed:
    failure = Err.Number
    description = Err.Description
    If failure = &H80070005 Then
        description = "Windows bağlantıyı engelledi (" & phase & ", 0x" & Hex$(failure) & ")."
    ElseIf failure = &H80072EE2 And resource = "/login" Then
        description = "Giriş doğrulaması zamanında yanıt vermedi. Şirket ağına bağlıyken tekrar deneyin."
    End If
    Err.Raise failure, , description
End Function

Private Function CallLocalBridge(ByVal method As String, ByVal resource As String, _
                                 ByVal payload As Variant, ByVal mime As String) As cSarfLocalReply
    Dim document As Object, node As Object, reply As cSarfLocalReply, envelope As New cSarfLocalReply
    Dim path As String, attempt As Long, failure As String, failureNumber As Long
    Dim handle As LongPtr, bytes() As Byte, count As Long, transferred As Long, offset As Long
    On Error GoTo Failed
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = document.appendChild(document.createElement("request"))
    node.setAttribute "method", method
    node.setAttribute "resource", resource
    node.setAttribute "mime", mime
    If Not IsEmpty(payload) Then
        node.DataType = "bin.base64"
        If VarType(payload) = vbString Then
            node.nodeTypedValue = Utf8Bytes(CStr(payload))
        Else
            node.nodeTypedValue = payload
        End If
    End If
    bytes = Utf8Bytes(document.XML)
    count = UBound(bytes) + 1
    path = Chr$(92) & Chr$(92) & "." & Chr$(92) & "pipe" & Chr$(92) & localPipe
    For attempt = 1 To 120
        handle = CreateFile(StrPtr(path), &HC0000000, 0, 0, 3, 0, 0)
        If handle <> -1 Then Exit For
        Sleep 40
    Next
    If handle = -1 Then Err.Raise vbObjectError + 706, , "Yerel bağlantı başlatılamadı. Kısayoldan yeniden açın."
    If WriteFile(handle, count, 4, transferred, 0) = 0 Or transferred <> 4 Then GoTo ConnectionLost
    If WriteFile(handle, bytes(0), count, transferred, 0) = 0 Or transferred <> count Then GoTo ConnectionLost
    If ReadFile(handle, count, 4, transferred, 0) = 0 Or transferred <> 4 Then GoTo ConnectionLost
    If count < 1 Or count > 33554432 Then GoTo ConnectionLost
    ReDim bytes(0 To count - 1)
    Do While offset < count
        If ReadFile(handle, bytes(offset), count - offset, transferred, 0) = 0 Or transferred = 0 Then GoTo ConnectionLost
        offset = offset + transferred
    Loop
    CloseHandle handle
    handle = 0
    envelope.ResponseBody = bytes
    document.async = False
    If Not document.LoadXML(envelope.ResponseText) Then Err.Raise vbObjectError + 704, , "Yerel servis yanıtı okunamadı."
    Set node = document.DocumentElement
    failure = AttributeValue(node, "error")
    If Len(failure) > 0 Then Err.Raise vbObjectError + 706, , failure
    Set reply = New cSarfLocalReply
    reply.Status = CLng(node.getAttribute("status"))
    node.DataType = "bin.base64"
    reply.ResponseBody = node.nodeTypedValue
    Set CallLocalBridge = reply
    Exit Function
ConnectionLost:
    Err.Raise vbObjectError + 706, , "Yerel bağlantı kesildi. Uygulamayı yeniden açın."
Failed:
    failureNumber = Err.Number
    failure = Err.Description
    If handle <> 0 And handle <> -1 Then CloseHandle handle
    Err.Raise failureNumber, , failure
End Function

Private Function Utf8Bytes(ByVal value As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    Utf8Bytes = stream.Read
    stream.Close
End Function

Public Function Xml(ByVal method As String, ByVal resource As String, _
                    Optional ByVal payload As Variant, Optional ByVal mime As String = "") As Object
    Dim response As Object, document As Object
    If IsMissing(payload) Then payload = Empty
    Set response = CallService(method, resource, payload, mime)
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    If Not document.LoadXML(response.ResponseText) Then Err.Raise vbObjectError + 704, , "Servis yaniti okunamadi."
    Set Xml = document
End Function

Public Sub Download(ByVal resource As String, ByVal filename As String, _
                    Optional ByVal payload As Variant, Optional ByVal mime As String = "")
    Dim response As Object, stream As Object
    If IsMissing(payload) Then payload = Empty
    Set response = CallService(IIf(IsEmpty(payload), "GET", "POST"), resource, payload, mime)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write response.ResponseBody
    stream.SaveToFile filename, 2
    stream.Close
End Sub

Public Function Upload(ByVal resource As String, ByVal filename As String) As String
    Dim stream As Object, payload As Variant
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filename
    payload = stream.Read
    stream.Close
    Upload = Xml("POST", resource, payload, "application/octet-stream").SelectSingleNode("/result/message").Text
End Function

Public Sub Disconnect()
    If Connected Then
        On Error Resume Next
        Call Xml("POST", "/logout")
        On Error GoTo 0
    End If
    token = ""
    localPipe = ""
End Sub

Public Function AttributeValue(ByVal node As Object, ByVal key As String) As String
    If Not IsNull(node.getAttribute(key)) Then AttributeValue = CStr(node.getAttribute(key))
End Function
