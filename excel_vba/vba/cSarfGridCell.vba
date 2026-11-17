Option Explicit

Public WithEvents Cell As MSForms.Label
Public WithEvents Header As MSForms.CommandButton
Public RowSlot As Long
Public ColumnSlot As Long
Public Owner As frmSarfAna

Private Sub Cell_Click()
    Owner.SelectGridRow RowSlot
End Sub

Private Sub Cell_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Owner.ShowGridRow RowSlot
End Sub

Private Sub Header_Click()
    Owner.FilterColumn ColumnSlot
End Sub
