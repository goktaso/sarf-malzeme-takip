Option Explicit

Private Sub Worksheet_Calculate()
    SarfToplam
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If SarfBusy Then Exit Sub
    If Not Intersect(Target, Me.Rows(5)) Is Nothing Then SarfToplam
End Sub
