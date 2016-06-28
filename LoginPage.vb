Public Class LoginPage

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles MngButton.Click
        'Input one management password
        'Run a message box dialog for a password login
        'Close the login form, open a new form which will form the main basis for the Management side
        MgmtPassword.Show()

    End Sub

    Private Sub EngnrButton_Click(sender As Object, e As EventArgs) Handles EngnrButton.Click
        'Input engineer username
        'Run a message box dialog for a just a username login
        'Close the login form, open a new form which will form the main basis for the Engineer side
        EngrUsername.Show()
    End Sub
End Class
