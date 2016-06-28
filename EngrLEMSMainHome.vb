Public Module engrSideDecVars
    Public infoButBool As Boolean
    Public rostButBool As Boolean
    Public rqstButBool As Boolean
    Public crdsButBool As Boolean
    Public darcButBool As Boolean
    Public settButBool As Boolean
End Module

Public Class EngrLEMSMainHome
    Private Sub EngrLEMSMainHome_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LogoutButton_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LogoutButton.LinkClicked
        LoginPage.Show()
        Me.Close()
        'upon closing, the variable storing the engineer username needs to be cleared, and therefore the linking database string
    End Sub

    Private Sub MyInfoButton_Click(sender As Object, e As EventArgs) Handles MyInfoButton.Click
        infoButBool = True
        EngrLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub ViewRosterButton_Click(sender As Object, e As EventArgs) Handles ViewRosterButton.Click
        rostButBool = True
        EngrLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub RequestsButton_Click(sender As Object, e As EventArgs) Handles RequestsButton.Click
        rqstButBool = True
        EngrLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub CDSButton_Click(sender As Object, e As EventArgs) Handles CDSButton.Click
        crdsButBool = True
        EngrLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub DSArchiveButton_Click(sender As Object, e As EventArgs) Handles DSArchiveButton.Click
        darcButBool = True
        EngrLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub StngsButton_Click(sender As Object, e As EventArgs) Handles StngsButton.Click
        settButBool = True
        EngrLEMSProcess.Show()
        Me.Close()
    End Sub
End Class