Public Class DailySheetOperationsEmailCredentials

    

    Private Sub DsOpsCredentialsDoneBut_Click(sender As Object, e As EventArgs) Handles DsOpsCredentialsDoneBut.Click
        'variables to hold the textbox information
        Dim employeeID As String
        Dim emailAddress As String
        Dim emailPassword As String

        'copy the data from the textboxes into the variables
        employeeID = DsOpsEmpNoTB.Text
        emailAddress = DsOpsEmailAddressTB.Text
        emailPassword = DsOpsEmailPasswordTB.Text

        'set the public variables to hold the credentials
        DailySheetOperations.senderEmpID = employeeID
        DailySheetOperations.senderAddress = emailAddress
        DailySheetOperations.senderPassword = emailPassword

        'close the form
        Me.Close()

    End Sub
End Class