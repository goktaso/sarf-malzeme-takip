Option Explicit

Public SarfClient As cSarfClient
Public SarfMenu As Object
Public SarfViewKey As String
Public SarfViewQuery As String
Public SarfBusy As Boolean

Public Sub SarfOturum()
    MsgBox "Uygulamayı Sarf Malzeme Takip kısayolundan açın.", vbInformation
    ThisWorkbook.Close SaveChanges:=False
End Sub

Public Sub SarfBaslaticiOturumu(ByVal bridgeFolder As String)
    On Error GoTo Failed
    If ThisWorkbook.Windows.Count > 0 Then ThisWorkbook.Windows(1).Visible = False
    Set SarfClient = New cSarfClient
    SarfClient.ConnectLocal bridgeFolder
    Set SarfMenu = SarfClient.Xml("GET", "/menu")
    frmSarfAna.Hazirla SarfMenu
    SarfShowDesktop frmSarfAna
Finished:
    On Error Resume Next
    Unload frmSarfAna
    SarfCikis
    Exit Sub
Failed:
    SarfHata
    Resume Finished
End Sub

Public Sub SarfBaglan(ByVal account As String, ByVal password As String, Optional ByVal endpoint As String = "")
    Set SarfClient = New cSarfClient
    If endpoint = "" Then endpoint = CStr(ThisWorkbook.Worksheets("Ana").Range("B7").Value2)
    SarfClient.Connect endpoint, account, password
    Set SarfMenu = SarfClient.Xml("GET", "/menu")
End Sub

Public Sub SarfCikis()
    If Not SarfClient Is Nothing Then SarfClient.Disconnect
    Set SarfClient = Nothing
    Set SarfMenu = Nothing
    SarfViewKey = ""
    With ThisWorkbook.Worksheets("Liste")
        If .ListObjects.Count > 0 Then .ListObjects(1).Unlist
        .Cells.ClearContents
    End With
    ThisWorkbook.Worksheets("Ana").Range("B9").ClearContents
End Sub

Public Sub SarfListe(ByVal key As String, ByVal query As String)
    Dim response As Object
    Set response = SarfClient.Xml("GET", "/list?name=" & SarfEncode(key) & query)
    SarfViewKey = key
    SarfViewQuery = query
    SarfRender response
End Sub

Public Sub SarfYenile()
    On Error GoTo Failed
    If SarfViewKey <> "" Then SarfListe SarfViewKey, SarfViewQuery
    Exit Sub
Failed:
    SarfHata
End Sub

Public Sub SarfYeniUrun()
    On Error GoTo Failed
    If SarfClient Is Nothing Then Exit Sub
    If SarfMenu Is Nothing Then Exit Sub
    If SarfMenu.SelectSingleNode("/menu/action[@key='product']") Is Nothing Then Exit Sub
    frmSarfIslem.Hazirla "product"
    SarfShowForm frmSarfIslem, 125
    Exit Sub
Failed:
    SarfHata
End Sub

Public Function SarfEncode(ByVal value As String) As String
    If Len(value) <> 0 Then SarfEncode = Application.WorksheetFunction.EncodeURL(value)
End Function

Public Function SarfParam(ByVal name As String, ByVal value As String) As String
    SarfParam = "&" & name & "=" & SarfEncode(value)
End Function

Public Function SarfOption(ByVal control As Object) As String
    If control.ListIndex >= 0 Then SarfOption = CStr(control.List(control.ListIndex, 1))
End Function

Public Sub SarfOptionAdd(ByVal control As Object, ByVal label As String, ByVal key As String)
    control.AddItem label
    control.List(control.ListCount - 1, 1) = key
End Sub

Public Sub SarfHata()
    MsgBox Err.Description, vbExclamation, "Sarf Malzeme Takip"
End Sub

Public Sub SarfDosya(ByVal mode As String, ByVal count As String)
    Dim dest As Variant, filename As String, countSuffix As String, uploaded As String
    filename = "Sarf-Aktarim-Sablonu.xlsx"
    If Len(count) Then
        countSuffix = "?count=" & SarfEncode(count)
        filename = "Sayim-" & Split(count, ":")(0) & "-Sablonu.xlsx"
    End If
    If mode = "template" Then
        dest = Application.GetSaveAsFilename(filename, "Excel (*.xlsx), *.xlsx")
        If VarType(dest) <> vbBoolean Then SarfClient.Download "/template" & countSuffix, CStr(dest)
    Else
        dest = Application.GetOpenFilename("Excel (*.xlsx), *.xlsx")
        If VarType(dest) <> vbBoolean Then
            uploaded = SarfClient.Upload("/upload" & countSuffix, CStr(dest))
            If Len(count) Then uploaded = uploaded & vbCrLf & "Stok için Sayımı stoğa işle düğmesini kullanın."
            MsgBox uploaded, vbInformation
        End If
    End If
End Sub
