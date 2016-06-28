Public Class MgmtPassword

    Private Sub MgmtPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox("Enter Password", MsgBoxStyle.Information, "System Message")
        LoginPage.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''Need facility to change password, temporary hardwired password below.
        'While TextBox1.Text <> "password"
        '    MsgBox("Password incorrect, try again", MsgBoxStyle.OkOnly, "Password Error")
        'End While
        MgmtLEMSMain.Show()
        Me.Close()
    End Sub
End Class