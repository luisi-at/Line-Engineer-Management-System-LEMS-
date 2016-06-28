Option Explicit On
Imports System.Data.OleDb
Imports System.Data
Imports System.Linq
Imports System.Environment
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

'used to store engineer IDs and their positions in lists
Public Structure tEngrIDAndListPostn
    Dim engrID As Integer
    Dim surname As String
    Dim workingOnDay As Boolean
End Structure

Public Class EngrLEMSProcess
    Public rosterMonth As String = MonthName(Month(Now), True)
    Public rosterYear As String = Year(Now)
    Public rosterFileName As String = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & rosterMonth & rosterYear & ".xls"
    Public engrIDForUse As Integer
    'colours for the shiftrosters
    Public shift1Colour As String
    Public shift2Colour As String
    Public shift3Colour As String
    Public restColour As String
    Public vacationColour As String
    Public sickColour As String
    Public trainingColour As String
    Public tdyColour As String
    Public ticColour As String
    Public engineerIDsPostns(1000) As tEngrIDAndListPostn
    Private Sub EngrLEMSProcess_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'start up routines
        'import the data
        'clean up the interface

        Dim startUpFileName As String
        startUpFileName = rosterFileName
        checkRosterFileExists(startUpFileName)
        loadEngrInfo(engrIDForUse)
        currentDailySheetFirstSave()

        'disable the remaining vacation days textbox
        EngrReqVacnDaysRemngTB.Enabled = False
        'disable the current shift on the specific day textbox
        EngrReqCurrentShiftTB.Enabled = False

        If infoButBool = True Then
            'show my info tab first
            Me.EngrTabCtrl.SelectedTab = MyInfoTab
        End If
        If rostButBool = True Then
            'show roster tab first
            Me.EngrTabCtrl.SelectedTab = ViewRosterTab
            'call the create/load roster routine
        End If
        If rqstButBool = True Then
            'show requests tab first
            Me.EngrTabCtrl.SelectedTab = RequestsTab
        End If
        If crdsButBool = True Then
            'show current daily sheet tab first
            Me.EngrTabCtrl.SelectedTab = CDSTab
        End If
        If darcButBool = True Then
            'show DS Archive tab first
            Me.EngrTabCtrl.SelectedTab = DSArchiveTav
        End If
        If settButBool = True Then
            'show settings tab first
            Me.EngrTabCtrl.SelectedTab = SettingsTab
        End If

    End Sub

    Private Sub loadOtherEngrs()
        'loads the other engineers so that the loaded engineer can choose a shift change
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim ds As New DataSet
        Dim dAdap As New OleDbDataAdapter
        Dim dAdapAllQuals As New OleDbDataAdapter
        Dim sql As String
        Dim engrIDs() As Integer
        Dim surnames() As String

        'sql string to query database
        sql = "SELECT [Engineers].[Engineer_ID], [Engineers].[Surname] FROM [Engineers]"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdap = New OleDbDataAdapter(sql, connec)
        'opends the database
        connec.Open()
        dAdap.Fill(ds, "Engineers")

        'copes the dataset columns to local arrays
        engrIDs = (From myRow In ds.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Surname")).ToArray
        surnames = (From myRow In ds.Tables(0).AsEnumerable Select myRow.Field(Of String)("Surname")).ToArray

        'inserts the data from the local arrays into the public structure
        For counter = 0 To engrIDs.Length - 1
            engineerIDsPostns(counter).engrID = engrIDs(counter)
            engineerIDsPostns(counter).surname = surnames(counter)
        Next

        'redefine the size of the structure to make searching more efficient
        ReDim Preserve engineerIDsPostns(engrIDs.Length)

        'add the engineers to the combo box
        For counter = 0 To engineerIDsPostns.Length - 1
            EngrReqOtrEngrsWorkingOnDayCB.Items.Add(engineerIDsPostns(counter).surname)
        Next

    End Sub

    Private Sub currentDailySheetFirstSave()
        'porblem here is that the SQL needs sorting out
        'other than that theres nothing wrong

        Try
            DailySheetOperations.mainControl()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub checkRosterFileExists(ByVal inRosterName As String)

        If File.Exists(inRosterName) Then
            'populate datagrid views from Excel file
            importFromExcelFile(inRosterName)
        Else
            MsgBox("The Shift Roster for " & Now.Month & " " & Now.Year & "does not currently exist. Contact your manager", MsgBoxStyle.OkOnly, "LEMS Message")
        End If
    End Sub

    Public Sub importFromExcelFile(ByVal inFileName As String)

        'the following code was adapted from:
        'http://vb.net-informations.com/excel-2007/vb.net_excel_2007_create_file.htm

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkBooks As Excel.Workbooks
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim workSheetRange As Excel.Range
        Dim inputArray(,) As Object
        'my variables
        Dim colTotal As Integer
        Dim rowTotal As Integer

        'copied code
        xlWorkBooks = xlApp.Workbooks
        xlWorkBook = xlWorkBooks.Open(inFileName)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        workSheetRange = xlWorkSheet.UsedRange

        inputArray = workSheetRange.Value

        colTotal = workSheetRange.Columns.Count 'System.DateTime.DaysInMonth(Year(Now), Month(Now))
        rowTotal = workSheetRange.Rows.Count

        fillDataGridViews(colTotal, rowTotal, inputArray)

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(workSheetRange)
        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkBooks)
        releaseObject(xlApp)

    End Sub

    Public Sub releaseObject(ByVal obj As Object)
        'the following code was taken and adapted from:
        'http://vb.net-informations.com/excel-2007/vb.net_excel_2007_create_file.htm

        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try

    End Sub

    Public Sub fillDataGridViews(ByVal noOfColumns As Integer, ByVal noOfRows As Integer, ByRef inArray(,) As Object)
        'the purpose of this subroutine is to provide reusable code to fill every relevent datagridview in the management form
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim ds As New DataSet
        Dim dAdap As New OleDbDataAdapter
        Dim dAdapAllQuals As New OleDbDataAdapter
        Dim sql As String
        Dim surnames() As String

        sql = "SELECT [Engineers].[Engineer_ID], [Engineers].[Surname], [Engineers].[Total_Qual_Score] FROM [Engineers] ORDER BY [Engineer_ID];"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdap = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdap.Fill(ds, "surnamesList")
        connec.Close()
        surnames = (From myRow In ds.Tables(0).AsEnumerable Select myRow.Field(Of String)("Surname")).ToArray

        For counter = 0 To surnames.Length - 1
            EngrRosterSpace.Rows.Add(surnames(counter))
        Next

        For rowCounter = 2 To noOfRows
            For columnCounter = 2 To noOfColumns
                EngrRosterSpace.Rows(rowCounter - 2).Cells(columnCounter - 1).Value = inArray(rowCounter, columnCounter)

            Next
        Next

        noOfColumns = System.DateTime.DaysInMonth(Year(Now), Month(Now))
        colorDataGrids(noOfRows, noOfColumns)

    End Sub

    Public Sub colorDataGrids(ByVal rowCount As Integer, ByVal colCount As Integer)
        'fills the areas in the datagridview where off days are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "R" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(restColour)
                End If
               
            Next
        Next
        'fills the areas in the datagridview where EE shifts are present
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "EE" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(shift1Colour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where EN shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "EN" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(shift2Colour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where LN shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "LN" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(shift3Colour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where TIC shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "TDY" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(tdyColour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where vacation shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "V" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(vacationColour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where planned sick days are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "SICK" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(sickColour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where training is present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "TRNG" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(trainingColour)
                End If
                
            Next
        Next
        'fills the areas in the datagridview where TDY is present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "TDY" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = Color.FromArgb(tdyColour)
                End If
               
            Next
        Next
    End Sub

    Private Sub loadEngrInfo(ByVal inEmpNo As String)
        'take the employee number and get the engineer id from the database
        'query the engineers table with that ID to return the details for that engineer

        'the code for the datareader was taken and adapted for use from:
        'http://www.dotnetcurry.com/showarticle.aspx?ID=143
        'after a Google search for: 'vb.net datareader into datatable'

        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim dAdapEngrInfo As New OleDbDataAdapter
        'Dim drEngrInfo As OleDbDataReader
        Dim dAdapEngrQualNames As New OleDbDataAdapter
        Dim drEngrVacation As OleDbDataReader
        Dim drEngrTDY As OleDbDataReader
        Dim drEngrRequestsApproved As OleDbDataReader
        Dim drEngrRequestsPending As OleDbDataReader
        Dim dtEngrInfo As New DataTable
        Dim dtEngrQualNames As New DataTable
        Dim dtEngrVacation As New DataTable
        Dim dtEngrTDY As New DataTable
        Dim dtEngrRequestsApproved As New DataTable
        Dim dtEngrRequestsPending As New DataTable
        Dim sql As String

        'local variables used to hold the data from the database:
        Dim localEngrForename As String
        Dim localEngrSurname As String
        Dim localEngrDoB As String
        Dim localEngrHireDate As String
        Dim localEngrEmpNo As Integer
        Dim localEngrAPNo As Integer
        Dim localEngrProfPic As String
        Dim localEngrQualsIDsHeldArray() As Integer
        Dim localEngrQualsHeldNamesArray(100) As String
        Dim localEngrEETotals As Integer
        Dim localEngrENTotals As Integer
        Dim localEngrLNTotals As Integer
        Dim localEngrTICTotals As Integer
        Dim localEngrVacDaysRemng As Integer

        'get all the relevant information from the engineer table
        'this needs sorting out as well
        sql = "SELECT * FROM [Engineers], [EN/QUALS] WHERE [Engineers].[Engineer_ID] = ?  AND [EN/QUALS].[EN_ID] = ?"
        Dim cmd = New OleDbCommand(sql, connec)

        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@Engineer_ID", engrIDForUse)
        cmd.Parameters.AddWithValue("@EN_ID", engrIDForUse)
        connec.Open()
        Dim drEngrInfo As OleDbDataReader
        drEngrInfo = cmd.ExecuteReader
        dtEngrInfo.Load(drEngrInfo)
        connec.Close()

        'fill the local arrays with the data from the database
        localEngrForename = dtEngrInfo.Rows(0)("Forename")
        localEngrSurname = dtEngrInfo.Rows(0)("Surname")
        localEngrProfPic = dtEngrInfo.Rows(0)("Engineer_Image")
        localEngrDoB = dtEngrInfo.Rows(0)("DoB")
        localEngrHireDate = dtEngrInfo.Rows(0)("Hire_Date")
        localEngrEETotals = dtEngrInfo.Rows(0)("EE_Shift_Occurrence")
        localEngrENTotals = dtEngrInfo.Rows(0)("EN_Shift_Occurrence")
        localEngrLNTotals = dtEngrInfo.Rows(0)("LN_Shift_Occurrence")
        localEngrTICTotals = dtEngrInfo.Rows(0)("TIC_Occurrence")
        localEngrVacDaysRemng = dtEngrInfo.Rows(0)("Vacation_Days_Remaining")
        localEngrEmpNo = dtEngrInfo.Rows(0)("Employee_Number")
        localEngrAPNo = dtEngrInfo.Rows(0)("A&P_License_Number")

        'assign the values of the labels and datagridview cells
        EngrMyInfoDispForenameLab.Text = localEngrForename
        EngrMyInfoDispSurnameLab.Text = localEngrSurname
        EngrMyInfoDispDoBLab.Text = localEngrDoB
        EngrMyInfoDispHireDateLab.Text = localEngrHireDate
        EngrMyInfoShiftTotalsDGV.Rows(0).Cells(0).Value = localEngrEETotals
        EngrMyInfoShiftTotalsDGV.Rows(0).Cells(1).Value = localEngrENTotals
        EngrMyInfoShiftTotalsDGV.Rows(0).Cells(2).Value = localEngrLNTotals
        EngrMyInfoShiftTotalsDGV.Rows(0).Cells(3).Value = localEngrTICTotals
        EngrMyInfoShiftTotalsDGV.Rows(0).Cells(4).Value = localEngrVacDaysRemng

        'assign the image location of the picture box
        EngrMyInfoProfPicPB.Image = Image.FromFile(localEngrProfPic)

        'get the IDs of the qualification that the engineer holds
        localEngrQualsIDsHeldArray = (From myRow In dtEngrInfo.AsEnumerable Select myRow.Field(Of Integer)("QUAL_ID")).ToArray

        'gets the name of each of the qualifications the engineer holds
        sql = "SELECT [Qualifications].[Qualification_Name] FROM [Qualifications] WHERE [Qualifications].[QUAL_ID] = ?"
        cmd = New OleDbCommand(sql, connec)
        connec.Open()
        For counter = 0 To localEngrQualsIDsHeldArray.Length - 1
            'add the parameter
            cmd.Parameters.AddWithValue("@QUAL_ID", localEngrQualsIDsHeldArray(counter).ToString)
            'create the data reader
            'created each time the loop goes round to prevent an error
            Dim drEngrQuals As OleDbDataReader = cmd.ExecuteReader()
            'fill the datatable
            dtEngrQualNames.Load(drEngrQuals)
            'copy the first row over to the array (as only one row returned)
            localEngrQualsHeldNamesArray(counter) = dtEngrQualNames.Rows(counter)("Qualification_Name")
            'clear the parameter to make ready for the next
            cmd.Parameters.Clear()
        Next

        connec.Close()

        'add the qualifications to the list box on the form
        For counter = 0 To localEngrQualsIDsHeldArray.Length - 1
            EngrHeldQualsLB.Items.Add(localEngrQualsHeldNamesArray(counter))
        Next

        'add the information from the vacation table in the database to the vacation datagridview
        sql = "SELECT [Vacation].[Vacation_ID], [Vacation].[V_Start_Date], [Vacation].[V_End_Date] FROM [Engineers], [Vacation], [EN/VA] WHERE [Engineers].[Engineer_ID] = ? AND [EN/VA].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/VA].[VA_ID] = [Vacation].[Vacation_ID]"
        cmd = New OleDbCommand(sql, connec)
        cmd.Parameters.AddWithValue("@Engineer_ID", engrIDForUse)
        connec.Open()
        drEngrVacation = cmd.ExecuteReader
        dtEngrVacation.Load(drEngrVacation)
        EngrMyInfoCurrentVacChoicesDGV.DataSource = dtEngrVacation

        'add the tdy assignments for the engineer from the database to the tdy datagridview
        sql = "SELECT [TDY].[TDY_ID], [TDY].[TDY_Start_Date], [TDY].[TDY_End_Date], [TDY].[TDY_Location] FROM [Engineers], [TDY], [EN/TDY] WHERE [Engineers].[Engineer_ID] = ? AND [EN/TDY].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/TDY].[TDY_ID] = [TDY].[TDY_ID]"
        cmd = New OleDbCommand(sql, connec)
        cmd.Parameters.AddWithValue("@Engineer_ID", engrIDForUse)
        connec.Open()
        drEngrTDY = cmd.ExecuteReader
        dtEngrTDY.Load(drEngrVacation)
        EngrMyInfoCurrentVacChoicesDGV.DataSource = dtEngrVacation

        'add the approved request information from the database to the approved request datagridview
        sql = "SELECT * FROM [Engineers], [Requests], [EN/REQS] WHERE [Engineers].[Engineer_ID] = ? AND [EN/REQS].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/REQS].[REQ_ID] = [REQS].[REQ_ID] AND [Requests].[Request_Approved] = TRUE"
        cmd = New OleDbCommand(sql, connec)
        cmd.Parameters.AddWithValue("@Engineer_ID", engrIDForUse)
        connec.Open()
        drEngrRequestsApproved = cmd.ExecuteReader
        dtEngrRequestsApproved.Load(drEngrRequestsApproved)
        EngrReqApprvdDGV.DataSource = dtEngrRequestsApproved

        'add the pending request information from the database to the pending request datagridview
        sql = "SELECT * FROM [Engineers], [Requests], [EN/REQS] WHERE [Engineers].[Engineer_ID] = ? AND [EN/REQS].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/REQS].[REQ_ID] = [REQS].[REQ_ID] AND [Requests].[Request_Approved] = FALSE"
        cmd = New OleDbCommand(sql, connec)
        cmd.Parameters.AddWithValue("@Engineer_ID", engrIDForUse)
        connec.Open()
        drEngrRequestsPending = cmd.ExecuteReader
        dtEngrRequestsPending.Load(drEngrRequestsPending)
        EngrReqPendingDGV.DataSource = dtEngrRequestsPending

    End Sub

    Private Sub loadShiftColours()
        'used to load the shift colours for the shift rosters

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift1.txt", OpenMode.Input)
        'assign the colour to the variable
        shift1Colour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift2.txt", OpenMode.Input)
        'assign the colour to the variable
        shift2Colour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift3.txt", OpenMode.Input)
        'assign the colour to the variable
        shift3Colour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Rest.txt", OpenMode.Input)
        'assign the colour to the variable
        restColour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Vacation.txt", OpenMode.Input)
        'assign the colour to the variable
        vacationColour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Sick.txt", OpenMode.Input)
        'assign the colour to the variable
        sickColour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Training.txt", OpenMode.Input)
        'assign the colour to the variable
        trainingColour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TDY.txt", OpenMode.Input)
        'assign the colour to the variable
        tdyColour = LineInput(1)
        'close the file
        FileClose(1)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TIC.txt", OpenMode.Input)
        'assign the colour to the variable
        ticColour = LineInput(1)
        'close the file
        FileClose(1)

    End Sub

    Private Sub HomeTab_Click(sender As Object, e As EventArgs) Handles HomeTab.Click
        EngrLEMSMainHome.Show()
        infoButBool = False
        rostButBool = False
        rqstButBool = False
        crdsButBool = False
        darcButBool = False
        settButBool = False
        Me.Close()
    End Sub

    Private Sub EngrTabCtrl_Selected(sender As Object, e As EventArgs) Handles EngrTabCtrl.Selected
        If EngrTabCtrl.SelectedTab.Name = "ViewRosterTab" Then
            'call the createRoster routine
        End If

    End Sub

    Private Sub LogoutButton_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LogoutButton.LinkClicked
        LoginPage.Show()
        Me.Close()
    End Sub

    Private Sub engrCdsSave()
        DailySheetOperations.saveDGV("engrLEMS")
    End Sub

    Private Sub EngrCdsSaveOnlyBut_Click(sender As Object, e As EventArgs) Handles EngrCdsSaveOnlyBut.Click
        engrCdsSave()
    End Sub

    Private Sub EngrCdsSaveSendBut_Click(sender As Object, e As EventArgs) Handles EngrCdsSaveSendBut.Click
        engrCdsSave()
        DailySheetOperations.createHardCopyDS()
    End Sub

    Private Sub EngrReqVacnChoiceSubBut_Click(sender As Object, e As EventArgs) Handles EngrReqVacnChoiceSubBut.Click
        'this sends a vacation request to the database
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim dAdapMostRecentVacationID As New OleDbDataAdapter
        Dim dsMostRecentVactionID As New DataSet
        Dim sql As String

        Dim tempStartDate As Date
        Dim tempEndDate As Date
        Dim startDateString As String
        Dim endDateString As String
        Dim vacChoice As String
        Dim mostRecentVacationID As Integer

        'take the values from the vacation date time pickers
        tempStartDate = EngrReqVacnChoiceStartDateDTP.Value
        tempEndDate = EngrReqVacnChoiceEndDateDTP.Value

        'send to the function that returns the string
        startDateString = returnStringDate(tempStartDate)
        endDateString = returnStringDate(tempEndDate)

        'writing to the database depending on the requirement
        If EngrReqVacnChoiceFirstRB.Checked = True Then
            'assign a unique value to the string so that the
            'preference can be determined by the program later
            'when processing the data to allocate the vacation
            vacChoice = "FIRST"
            'insert the engineer's vacation into the database
            sql = "INSERT INTO  [Vacation] ([Vacation].[V_Start_Date], [Vacation].[V_End_Date], [Vacation].[Vacation_Preference]) VALUES (?,?,?)"
            cmd = New OleDbCommand(sql, connec)
            cmd.CommandType = CommandType.Text
            connec.Open()
            cmd.Parameters.AddWithValue("@V_Start_Date", startDateString)
            cmd.Parameters.AddWithValue("@V_End_Date", endDateString)
            cmd.Parameters.AddWithValue("@Vacation Preference", vacChoice)
            cmd.ExecuteNonQuery()
            connec.Close()
        End If
        If EngrReqVacnChoiceSecondRB.Checked = True Then
            vacChoice = "SCND"
            'insert the engineer's vacation into the database
            sql = "INSERT INTO  [Vacation] ([Vacation].[V_Start_Date], [Vacation].[V_End_Date], [Vacation].[Vacation_Preference]) VALUES (?,?,?)"
            cmd = New OleDbCommand(sql, connec)
            cmd.CommandType = CommandType.Text
            connec.Open()
            cmd.Parameters.AddWithValue("@V_Start_Date", startDateString)
            cmd.Parameters.AddWithValue("@V_End_Date", endDateString)
            cmd.Parameters.AddWithValue("@Vacation Preference", vacChoice)
            cmd.ExecuteNonQuery()
            connec.Close()
        End If
        If EngrReqVacnChoiceThirdRB.Checked = True Then
            vacChoice = "THRD"
            'insert the engineer's vacation into the database
            sql = "INSERT INTO  [Vacation] ([Vacation].[V_Start_Date], [Vacation].[V_End_Date], [Vacation].[Vacation_Preference]) VALUES (?,?,?)"
            cmd = New OleDbCommand(sql, connec)
            cmd.CommandType = CommandType.Text
            connec.Open()
            cmd.Parameters.AddWithValue("@V_Start_Date", startDateString)
            cmd.Parameters.AddWithValue("@V_End_Date", endDateString)
            cmd.Parameters.AddWithValue("@Vacation Preference", vacChoice)
            cmd.ExecuteNonQuery()
            connec.Close()
        End If

        'return the last vacation ID so that the link table can be updated
        sql = "SELECT LAST ([Vacation].[Vacation_ID]) FROM [Vacation]"
        cmd = New OleDbCommand(sql, connec)
        dAdapMostRecentVacationID = New OleDbDataAdapter(sql, connec)
        dAdapMostRecentVacationID.Fill(dsMostRecentVactionID, "Vacation")
        connec.Close()
        'assign the most recent ID to a variable
        mostRecentVacationID = dsMostRecentVactionID.Tables(0).Rows(0)("Vacation_ID")

        'insert the engineer ID and the vacation ID into the link table
        sql = "INSERT INTO [EN/VA] ([EN/VA].[EN_ID], [EN/VA].[VACTN_ID]) VALUES (?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@EN_ID", engrIDForUse)
        cmd.Parameters.AddWithValue("@VACTN_ID", mostRecentVacationID)
        cmd.ExecuteNonQuery()
        connec.Close()

        'message box to show user successful database update for vacation
        MessageBox.Show("Vacation Successfully Requested", "LEMS Message: Successful Request", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub getEngrsWorkingOnDay(ByVal inDate As Date)
        'determines which engineers are working on which day
        Dim localString As String

        'loop through roster column, set to true if the engineer is working, i.e. satisfies the conditions below
        For rowCounter = 0 To EngrRosterSpace.Rows.Count - 1
            localString = EngrRosterSpace.Rows(rowCounter).Cells(inDate.Day).Value
            Select Case localString
                Case Is <> "V"
                    engineerIDsPostns(rowCounter).workingOnDay = True
                Case Is <> "MED"
                    engineerIDsPostns(rowCounter).workingOnDay = True
                Case Is <> "SICK"
                    engineerIDsPostns(rowCounter).workingOnDay = True
            End Select
            If Not localString.Contains("TRNG") Or Not localString.Contains("TDY") Then
                engineerIDsPostns(rowCounter).workingOnDay = True
            End If
        Next
    End Sub

    Private Sub EngrShiftChangeSubBut_Click(sender As Object, e As EventArgs) Handles EngrShiftChangeSubBut.Click
        'this sends a shift change request to the database
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim sql As String

        'some local date manipulation required
        Dim tempDate As Date
        Dim engrRequestType As String
        Dim engrDateRequiredString As String
        Dim engrShiftChangeFrom As String
        Dim engrShiftChangeTo As String
        Dim engrChangeWithEngr As String = Nothing
        Dim engrRequestDescrip As String

        'assign values to the variables to be used as parameters in the SQL
        tempDate = EngrReqShiftChangeDateDTP.Value
        getEngrsWorkingOnDay(tempDate)
        engrDateRequiredString = returnStringDate(tempDate)
        engrRequestType = "SC"
        engrShiftChangeFrom = EngrReqCurrentShiftTB.Text

        If IsNothing(EngrReqOtrEngrsWorkingOnDayCB.SelectedItem) Then
            MessageBox.Show("An engineer hasn't been selected to swap a shift with", "LEMS Warning: Engineer Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf engineerIDsPostns(EngrReqOtrEngrsWorkingOnDayCB.SelectedIndex).workingOnDay = False Then
            MessageBox.Show("The selected engineer is not eligable to work on this day", "LEMS Warning: Engineer Not Eligable to Swap", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Else
            engrChangeWithEngr = engineerIDsPostns(EngrReqOtrEngrsWorkingOnDayCB.SelectedIndex).engrID
            engrRequestDescrip = "Shift Change"

            'writing to the database depending on the requirement
            If EngrReqEEShiftRB.Checked = True Then
                engrShiftChangeTo = "EE"
                'insert the shift change of Early Early into the engineers request table
                sql = "INSERT INTO  [Requests] ([Requests].[Request_Type], [Request].[Request_Date], [Request].[Request_Shift_Change_From], [Request].[Request_Shift_Change_To], [Request].[Request_Shift_Change_With_Engineer], [Request].[Request_Description ) VALUES (?,?,?,?,?,?)"
                cmd = New OleDbCommand(sql, connec)
                cmd.CommandType = CommandType.Text
                connec.Open()
                cmd.Parameters.AddWithValue("@Request_Type", engrRequestType)
                cmd.Parameters.AddWithValue("@Request_Date", engrDateRequiredString)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_From", engrShiftChangeFrom)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_To", engrShiftChangeTo)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_With_Engineer", engrChangeWithEngr)
                cmd.Parameters.AddWithValue("@Request_Description", engrRequestDescrip)
                cmd.ExecuteNonQuery()
                connec.Close()
            End If
            If EngrReqENShiftRB.Checked = True Then
                engrShiftChangeTo = "EN"
                'insert the shift change of Early Normal into the engineers request table
                sql = "INSERT INTO  [Requests] ([Requests].[Request_Type], [Request].[Request_Date], [Request].[Request_Shift_Change_From], [Request].[Request_Shift_Change_To], [Request].[Request_Shift_Change_With_Engineer], [Request].[Request_Description ) VALUES (?,?,?,?,?,?)"
                cmd = New OleDbCommand(sql, connec)
                cmd.CommandType = CommandType.Text
                connec.Open()
                cmd.Parameters.AddWithValue("@Request_Type", engrRequestType)
                cmd.Parameters.AddWithValue("@Request_Date", engrDateRequiredString)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_From", engrShiftChangeFrom)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_To", engrShiftChangeTo)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_With_Engineer", engrChangeWithEngr)
                cmd.Parameters.AddWithValue("@Request_Description", engrRequestDescrip)
                cmd.ExecuteNonQuery()
                connec.Close()
            End If
            If EngrReqLNShiftRB.Checked = True Then
                engrShiftChangeTo = "LN"
                'insert the shift change of Late Normal into the engineers request table
                sql = "INSERT INTO  [Requests] ([Requests].[Request_Type], [Request].[Request_Date], [Request].[Request_Shift_Change_From], [Request].[Request_Shift_Change_To], [Request].[Request_Shift_Change_With_Engineer], [Request].[Request_Description ) VALUES (?,?,?,?,?,?)"
                cmd = New OleDbCommand(sql, connec)
                cmd.CommandType = CommandType.Text
                connec.Open()
                cmd.Parameters.AddWithValue("@Request_Type", engrRequestType)
                cmd.Parameters.AddWithValue("@Request_Date", engrDateRequiredString)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_From", engrShiftChangeFrom)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_To", engrShiftChangeTo)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_With_Engineer", engrChangeWithEngr)
                cmd.Parameters.AddWithValue("@Request_Description", engrRequestDescrip)
                cmd.ExecuteNonQuery()
                connec.Close()
            End If
            If EngrReqRestDayRB.Checked = True Then
                engrShiftChangeTo = "R"
                'insert the change for a rest day into the engineers request table
                sql = "INSERT INTO  [Requests] ([Requests].[Request_Type], [Request].[Request_Date], [Request].[Request_Shift_Change_From], [Request].[Request_Shift_Change_To], [Request].[Request_Shift_Change_With_Engineer], [Request].[Request_Description ) VALUES (?,?,?,?,?,?)"
                cmd = New OleDbCommand(sql, connec)
                cmd.CommandType = CommandType.Text
                connec.Open()
                cmd.Parameters.AddWithValue("@Request_Type", engrRequestType)
                cmd.Parameters.AddWithValue("@Request_Date", engrDateRequiredString)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_From", engrShiftChangeFrom)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_To", engrShiftChangeTo)
                cmd.Parameters.AddWithValue("@Request_Shift_Change_With_Engineer", engrChangeWithEngr)
                cmd.Parameters.AddWithValue("@Request_Description", engrRequestDescrip)
                cmd.ExecuteNonQuery()
                connec.Close()
            End If

            ''save the details to the link table
            saveToRequestsLinkTable()

            'message box to show user successful database update for vacation
            MessageBox.Show("Shift Change Successfully Requested", "LEMS Message: Successful Request", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

       

    End Sub

    Private Sub EngrReqRestToOTSubBut_Click(sender As Object, e As EventArgs) Handles EngrReqRestToOTSubBut.Click
        'this sends an overtime request to the database
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim sql As String

        'some local date manipulation required
        Dim tempDate As Date
        Dim engrRequestType As String
        Dim engrDateRequiredString As String
        Dim engrRequestDescrip As String

        'assign values to the variables to be used as parameters in the SQL
        tempDate = EngrReqRestDayOTDateDTP.Value
        engrDateRequiredString = returnStringDate(tempDate)
        engrRequestType = "OT"
        engrRequestDescrip = "Overtime"

        'writing to the database
        sql = "INSERT INTO  [Requests] ([Requests].[Request_Type], [Request].[Request_Date], [Request].[Request_Description ) VALUES (?,?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@Request_Type", engrRequestType)
        cmd.Parameters.AddWithValue("@Request_Date", engrDateRequiredString)
        cmd.Parameters.AddWithValue("@Request_Description", engrRequestDescrip)
        cmd.ExecuteNonQuery()
        connec.Close()

        'save to the link table
        saveToRequestsLinkTable()

        'message box to show user successful database update for vacation
        MessageBox.Show("Overtime Successfully Requested", "LEMS Message: Successful Request", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub EngrReqMedSubBut_Click(sender As Object, e As EventArgs) Handles EngrReqMedSubBut.Click
        'sends a medical appointment request to the database
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim sql As String

        'some local date manipulation required
        Dim tempDate As Date
        Dim engrRequestType As String
        Dim engrDateRequiredString As String
        Dim engrRequestDescrip As String

        'assign values to the variables to be used as parameters in the SQL
        tempDate = EngrReqRestDayOTDateDTP.Value
        engrDateRequiredString = returnStringDate(tempDate)
        engrRequestType = "MED"
        engrRequestDescrip = EngrReqMedDescrpTB.Text

        'writing to the database
        sql = "INSERT INTO  [Requests] ([Requests].[Request_Type], [Request].[Request_Date], [Request].[Request_Description ) VALUES (?,?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@Request_Type", engrRequestType)
        cmd.Parameters.AddWithValue("@Request_Date", engrDateRequiredString)
        cmd.Parameters.AddWithValue("@Request_Description", engrRequestDescrip)
        cmd.ExecuteNonQuery()
        connec.Close()

        'message box to show user successful database update for vacation
        MessageBox.Show("Medical Leave Successfully Requested", "LEMS Message: Successful Request", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Public Function returnStringDate(ByVal inDate As Date) As String
        'returns the string instance of the date so that it can be written to
        'the database
        Dim returnDateString As String

        returnDateString = inDate.Day.ToString & "/" & inDate.Month.ToString & "/" & inDate.Year.ToString

        Return returnDateString
    End Function

    Private Sub saveToRequestsLinkTable()
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim dAdapMostRecentVacationID As New OleDbDataAdapter
        Dim dsMostRecentVactionID As New DataSet
        Dim sql As String

        Dim mostRecentRequestID As Integer

        'return the last request ID so that the link table can be updated
        sql = "SELECT LAST ([Requests].[Request_ID]) FROM [Requests]"
        cmd = New OleDbCommand(sql, connec)
        dAdapMostRecentVacationID = New OleDbDataAdapter(sql, connec)
        dAdapMostRecentVacationID.Fill(dsMostRecentVactionID, "Requests")
        connec.Close()
        'assign the most recent ID to a variable
        mostRecentRequestID = dsMostRecentVactionID.Tables(0).Rows(0)("Request_ID")

        'insert the engineer ID and the request ID into the link table
        sql = "INSERT INTO [EN/REQ] ([EN/REQ].[EN_ID], [EN/REQ].[REQ_ID]) VALUES (?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@EN_ID", engrIDForUse)
        cmd.Parameters.AddWithValue("@REQ_ID", mostRecentRequestID)
        cmd.ExecuteNonQuery()
        connec.Close()
    End Sub

    Private Sub EngrResultList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles EngrResultList.SelectedIndexChanged
        'code taken and adapted for use from:
        'https://social.msdn.microsoft.com/Forums/vstudio/en-US/c165c82d-ac79-477e-abab-3efd330d149f/how-to-open-pdf-file-in-vbnet-applicatin

        EngrDSAAxAcroPDF1.src = EngrResultList.SelectedItem
    End Sub

    Private Sub EngrGeneralResultAdvSearchDSArchiveRB_CheckedChanged(sender As Object, e As EventArgs) Handles EngrGeneralResultAdvSearchDSArchiveRB.CheckedChanged
        If EngrGeneralResultAdvSearchDSArchiveRB.Checked Then
            EngrMetricsAdvSearchDSArchivePanel.Show()
            EngrDSAAxAcroPDF1.Show()
        End If
    End Sub

    Private Sub EngrBySheetAdvSearchDSArchiveRB_CheckedChanged(sender As Object, e As EventArgs) Handles EngrBySheetAdvSearchDSArchiveRB.CheckedChanged
        If EngrBySheetAdvSearchDSArchiveRB.Checked Then
            EngrDSAAxAcroPDF1.Show()
            EngrMetricsAdvSearchDSArchivePanel.Hide()
        End If
    End Sub
End Class