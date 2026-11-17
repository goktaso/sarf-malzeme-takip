Option Explicit

Private choices As Object
Private marked As Object
Private changing As Boolean
Public Applied As Boolean
Public Selection As Object

Public Sub Hazirla(ByVal title As String, ByVal values As Object, ByVal current As Object)
    Dim key As Variant
    lblTitle.Caption = title
    Set choices = values
    Set marked = CreateObject("Scripting.Dictionary")
    If current Is Nothing Then Set current = values
    For Each key In current.Keys
        marked(key) = True
    Next
    lstValues.ListStyle = 1
    lstValues.MultiSelect = 1
    Rebuild
End Sub

Private Sub Rebuild()
    Dim key As Variant
    changing = True
    lstValues.Clear
    For Each key In choices.Keys
        If InStr(1, choices(key), txtFind.Text, vbTextCompare) > 0 Or txtFind.Text = "" Then
            lstValues.AddItem choices(key)
            lstValues.List(lstValues.ListCount - 1, 1) = key
            lstValues.Selected(lstValues.ListCount - 1) = marked.Exists(key)
        End If
    Next
    chkAll.Value = choices.Count > 0 And AllChoicesMarked()
    changing = False
End Sub

Private Sub lstValues_Change()
    If changing Then Exit Sub
    Dim i As Long, key As String
    For i = 0 To lstValues.ListCount - 1
        key = lstValues.List(i, 1)
        If lstValues.Selected(i) Then
            marked(key) = True
        ElseIf marked.Exists(key) Then
            marked.Remove key
        End If
    Next
    changing = True
    chkAll.Value = choices.Count > 0 And AllChoicesMarked()
    changing = False
End Sub

Private Sub txtFind_Change()
    If Not choices Is Nothing Then Rebuild
End Sub

Private Sub chkAll_Click()
    If changing Or choices Is Nothing Then Exit Sub
    Dim key As Variant
    marked.RemoveAll
    If chkAll.Value Then
        For Each key In choices.Keys
            marked(key) = True
        Next
    End If
    Rebuild
End Sub

Private Sub cmdReset_Click()
    Applied = True
    Set Selection = Nothing
    Me.Hide
End Sub

Private Sub cmdApply_Click()
    Applied = True
    Set Selection = Nothing
    If marked.Count > 0 And Not AllChoicesMarked() Then Set Selection = marked
    Me.Hide
End Sub

Private Function AllChoicesMarked() As Boolean
    Dim key As Variant
    If marked.Count <> choices.Count Then Exit Function
    For Each key In choices.Keys
        If Not marked.Exists(key) Then Exit Function
    Next
    AllChoicesMarked = True
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub
