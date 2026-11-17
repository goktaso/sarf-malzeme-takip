Option Explicit

Private Sub UserForm_Initialize()
    Dim network As Object, item As Variant
    For Each item In Array("lblTitle", "lblSubtitle", "lblAccount", "lblPassword", "lblMessage")
        Me.Controls(item).BackStyle = fmBackStyleTransparent
        Me.Controls(item).Font.Name = "Aptos"
        Me.Controls(item).Font.Size = 11
    Next
    lblBanner.BackColor = RGB(17, 47, 41)
    lblTitle.ForeColor = vbWhite
    lblSubtitle.ForeColor = RGB(192, 230, 213)
    lblMessage.ForeColor = RGB(154, 34, 34)
    lblTitle.Font.Size = 20
    txtAccount.Font.Size = 11
    txtPassword.Font.Size = 11
    txtAccount.BackColor = RGB(232, 241, 237)
    cmdCancel.BackColor = RGB(232, 241, 237)
    cmdCancel.ForeColor = RGB(17, 47, 41)
    cmdLogin.BackColor = RGB(15, 90, 64)
    lblTitle.Font.Bold = True
    Set network = CreateObject("WScript.Network")
    txtAccount.Text = network.UserDomain & Chr$(92) & network.UserName
    txtAccount.Locked = True
    txtAccount.TabStop = False
    txtPassword.PasswordChar = "*"
    cmdLogin.Default = True
    cmdCancel.Cancel = True
End Sub

Private Sub UserForm_Activate()
    If txtAccount.Locked Then
        txtPassword.SetFocus
    Else
        txtAccount.SetFocus
    End If
End Sub

Public Sub HesapDegistir()
    txtAccount.Locked = False
    txtAccount.TabStop = True
    txtAccount.BackColor = vbWhite
    txtAccount.Text = ""
    txtAccount.ControlTipText = "Windows hesabını ETKİALANI\kullanıcı biçiminde yazın."
    lblSubtitle.Caption = "Başka bir Windows hesabıyla giriş"
End Sub

Private Sub cmdLogin_Click()
    On Error GoTo Failed
    If Len(Trim$(txtAccount.Text)) = 0 Then
        lblMessage.Caption = "Windows hesabınızı girin."
        txtAccount.SetFocus
        Exit Sub
    End If
    If Len(txtPassword.Text) = 0 Then
        lblMessage.Caption = "Windows parolanızı girin."
        Exit Sub
    End If
    cmdLogin.Enabled = False
    SarfBaglan Trim$(txtAccount.Text), txtPassword.Text
    txtPassword.Text = ""
    Me.Hide
    Exit Sub
Failed:
    lblMessage.Caption = Err.Description
    txtPassword.Text = ""
    cmdLogin.Enabled = True
    txtPassword.SetFocus
End Sub

Private Sub cmdCancel_Click()
    txtPassword.Text = ""
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub
