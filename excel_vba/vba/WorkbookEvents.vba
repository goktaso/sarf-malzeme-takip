Option Explicit

Private Sub Workbook_Open()
    SarfOturum
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not SarfClient Is Nothing Then SarfCikis
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not SarfClient Is Nothing Then
        If SarfClient.Connected Then SarfCikis
    End If
End Sub
