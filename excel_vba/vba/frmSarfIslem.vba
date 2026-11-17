Option Explicit

Private actionKey As String
Private nonce As String
Private descriptions As Collection
Private editors As Collection
Private lookup As Object
Private fixedCount As Boolean
Public SavedCountId As String
Public CountEntryMode As String

Public Sub HazirlaDetay(ByVal details As String)
    Dim box As Object
    lblTitle.Caption = "Satır ayrıntısı"
    cmdLoad.Visible = False
    cmdManual.Visible = False
    cmdSave.Visible = False
    cmdCancel.Caption = "Kapat"
    Set box = fraFields.Controls.Add("Forms.TextBox.1", "rowDetails", True)
    With box
        .Move 8, 8, 505, 255
        .MultiLine = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .Font.Size = 11
        .Text = details
        .SelStart = 0
    End With
End Sub

Public Sub Hazirla(ByVal key As String, Optional ByVal record As String = "")
    Dim xml As Object, node As Object, opt As Object, control As Object, caption As Object
    Dim mode As String, current As String, top As Single, i As Long, j As Long
    If Not editors Is Nothing Then
        For i = 1 To editors.Count
            fraFields.Controls.Remove "field" & i
            fraFields.Controls.Remove "caption" & i
        Next
    End If
    Set descriptions = New Collection
    Set editors = New Collection
    Set lookup = Nothing
    actionKey = key
    fixedCount = (key = "count-line" And Len(record) > 0)
    cmdManual.Visible = (key = "count-new")
    cmdSave.Caption = "Kaydet"
    If key = "count-new" Then cmdSave.Caption = "Excel ile giriş"
    Set xml = SarfClient.Xml("GET", "/form?name=" & SarfEncode(key) & SarfParam("record", record))
    lblTitle.Caption = xml.DocumentElement.getAttribute("title")
    nonce = xml.DocumentElement.getAttribute("nonce")
    top = 8
    For Each node In xml.SelectNodes("/form/field")
        i = descriptions.Count + 1
        descriptions.Add node
        mode = node.getAttribute("kind")
        current = node.getAttribute("value")
        Set caption = fraFields.Controls.Add("Forms.Label.1", "caption" & i, True)
        caption.Caption = node.getAttribute("label")
        caption.Move 8, top + 3, 178, 20
        If mode = "select" Or mode = "multi" Then
            Set control = fraFields.Controls.Add(IIf(mode = "select", "Forms.ComboBox.1", "Forms.ListBox.1"), "field" & i, True)
            control.ColumnCount = 2
            control.ColumnWidths = "315 pt;0 pt"
            If mode = "select" Then control.Style = 2 Else control.MultiSelect = 1
            For Each opt In node.SelectNodes("option")
                SarfOptionAdd control, opt.Text, opt.getAttribute("key")
            Next
            For j = 0 To control.ListCount - 1
                If mode = "select" Then
                    If CStr(control.List(j, 1)) = current Then control.ListIndex = j
                Else
                    control.Selected(j) = (InStr(1, "," & current & ",", "," & control.List(j, 1) & ",", vbBinaryCompare) > 0)
                End If
            Next
        Else
            Set control = fraFields.Controls.Add("Forms.TextBox.1", "field" & i, True)
            control.Text = current
        End If
        control.Move 192, top, 318, IIf(mode = "multi", 125, 22)
        If fixedCount And node.getAttribute("key") = "count_key" Then control.Enabled = False
        editors.Add control
        If node.getAttribute("lookup") = "1" Then Set lookup = control
        top = top + IIf(mode = "multi", 134, 31)
    Next
    fraFields.ScrollHeight = top + 9
    fraFields.ScrollTop = 0
    cmdLoad.Visible = Not lookup Is Nothing
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo Failed
    If Not lookup Is Nothing Then
        If SarfOption(lookup) <> "" Then Hazirla actionKey, SarfOption(lookup)
    End If
    Exit Sub
Failed:
    SarfHata
End Sub

Private Sub cmdSave_Click()
    SaveAction "excel"
End Sub

Private Sub cmdManual_Click()
    SaveAction "manual"
End Sub

Private Sub SaveAction(ByVal entryMode As String)
    On Error GoTo Failed
    Dim payload As String, node As Object, control As Object, i As Long, j As Long, value As String, mode As String
    Dim response As Object
    payload = "action=" & SarfEncode(actionKey) & SarfParam("request_id", nonce)
    For i = 1 To descriptions.Count
        Set node = descriptions(i)
        Set control = editors(i)
        mode = node.getAttribute("kind")
        value = ""
        Select Case mode
            Case "select": value = SarfOption(control)
            Case "multi"
                For j = 0 To control.ListCount - 1
                    If control.Selected(j) Then value = value & IIf(value = "", "", ",") & control.List(j, 1)
                Next
            Case Else
                value = Trim$(control.Text)
                If mode = "required" And value = "" Then Err.Raise vbObjectError + 705, , node.getAttribute("label") & " girin."
        End Select
        payload = payload & SarfParam(node.getAttribute("key"), value)
    Next
    Set response = SarfClient.Xml("POST", "/apply", payload, "application/x-www-form-urlencoded")
    If actionKey = "count-new" Then
        SavedCountId = SarfClient.AttributeValue(response.DocumentElement, "count_id")
        CountEntryMode = entryMode
        Me.Hide
        Exit Sub
    End If
    MsgBox response.SelectSingleNode("/result/message").Text, vbInformation, "Sarf Malzeme Takip"
    Unload Me
    Exit Sub
Failed:
    SarfHata
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
