Imports System.Data
Imports System.Data.OleDb
Imports System.Environment
Imports System.Windows.Forms
Imports System.Globalization
Imports System.IO

'Module RetrieveUsrNameData
'    Public EngrUsrNameChk As Integer
'    'pass the contents of the username to the search module
'    'if there is a match, it logs in
'    'will need some sql and database operations here for username re-entry and validation
'End Module
Public Class EngrUsername
    Public engrIDToPass As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MessageBox.Show("Enter your Employee Number to log in to LEMS", "LEMS Message: Login", MessageBoxButtons.OK, MessageBoxIcon.Information)
        LoginPage.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Access database with SQL
        'Fill Engineer view with relevant engineer information from database
        'The employee username is their employee number as this is used with the rest of the Delta systems
        'Pass the username to the next form
        'Temporary access:
        EngrLEMSMainHome.Show()
        Me.Close()

        'when ready, remove the above and replace with:
        'checkUsername
    End Sub

    Private Sub checkUsername()
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim dAdapReturnUnames As New OleDbDataAdapter
        Dim dsReturnUnames As New DataSet
        Dim sql As String
        Dim localReturnUnameArray() As String

        'queries the database to see if the input engineer username is correct
        'this needs to be sorted as it uses the SELECT with parameters
        sql = "SELECT [Employee_and_A&P].[Engineer_ID], [Employee_and_A&P].[Employee_Number] WHERE [Employee_and_A&P].[Employee_Number] = ?"
        Dim cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@Employee_Number", EngrUsernameLoginTB.Text)
        connec.Open()
        dAdapReturnUnames.Fill(dsReturnUnames, "surnamesList")
        connec.Close()

        'returns an array
        localReturnUnameArray = (From myRow In dsReturnUnames.Tables(0).AsEnumerable Select myRow.Field(Of String)("Employee_Number")).ToArray

        'if the array is empty, then the username does not exist
        If localReturnUnameArray(0) = Nothing Then
            MessageBox.Show("The Employee Number entered is incorrect or does not exist. Try again", "LEMS Warning: Login Failure", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
            EngrUsernameLoginTB.Clear()

        Else
            '"log in" to LEMS
            'pass the engineer ID associated with the employee number over to
            'the engineer side. Used later to query the DB for that
            'engineer's information to display
            engrIDToPass = dsReturnUnames.Tables(0).Rows(0)("Engineer_ID")
            MessageBox.Show("Successfully Logged into LEMS", "LEMS Message: Login Successful", MessageBoxButtons.OK)
            'pass the variable over to the next form
            EngrLEMSProcess.engrIDForUse = engrIDToPass
        End If
    End Sub

End Class