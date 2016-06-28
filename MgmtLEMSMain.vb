'This routine is to manage the clicks from the homepage to the operations page
'It acts as a starting point
Public Module mgmmtSideDecVars
    Public mgmtMngEngrsBool As Boolean
    Public mgmtCreateRostBool As Boolean
    Public mgmtViewRostBool As Boolean
    Public mgmtMonthlyHrsBool As Boolean
    Public mgmtSignSheetBool As Boolean
    Public mgmtCurrentDSBool As Boolean
    Public mgmtDSArchiveBool As Boolean
    Public mgmtSettBool As Boolean
End Module
Public Class MgmtLEMSMain
    Private Sub MgmtLEMSMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub MgmtMngEngrsBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsBut.Click
        mgmtMngEngrsBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtCreRostBut_Click(sender As Object, e As EventArgs) Handles MgmtCreRostBut.Click
        mgmtCreateRostBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtVwRostBut_Click(sender As Object, e As EventArgs) Handles MgmtVwRostBut.Click
        mgmtViewRostBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtVwMnHrsBut_Click(sender As Object, e As EventArgs) Handles MgmtVwMnHrsBut.Click
        mgmtMonthlyHrsBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtPrtSignInShBut_Click(sender As Object, e As EventArgs) Handles MgmtPrtSignInShBut.Click
        mgmtSignSheetBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtCDSBut_Click(sender As Object, e As EventArgs) Handles MgmtCDSBut.Click
        mgmtCurrentDSBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtDsArcBut_Click(sender As Object, e As EventArgs) Handles MgmtDsArcBut.Click
        mgmtDSArchiveBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub MgmtSettBut_Click(sender As Object, e As EventArgs) Handles MgmtSettBut.Click
        mgmtSettBool = True
        MgmtLEMSProcess.Show()
        Me.Close()
    End Sub

    Private Sub LogoutButton_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LogoutButton.LinkClicked
        LoginPage.Show()
        Me.Close()
        'need to clear the password variable at some point
    End Sub
End Class