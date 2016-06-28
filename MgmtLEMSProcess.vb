'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'LINE ENGINEER MANAGEMENT SYSTEM
'MANAGEMENT SIDE CODE
'ALEXANDER LUISI
'CANDIDATE NUMBER:  8045
'CENTRE NUMBER:     61517
'ALL AUTO-GENERATED CODE IS REFERENCED IN THE PROGRAM LISTING AS IN-LINE COMMENTS
'ALL CODE SNIPPETS ARE REFERENCED IN THE PROGRAM LISTING AS IN-LINE COMMENTS
'WITH A LINK TO THE SOURCE
'///////////////////////////////////////////////////
Option Explicit On
Imports System
Imports System.Environment
Imports System.Windows.Forms
Imports System.Globalization
Imports System.IO
Imports System.Data.OleDb
Imports System.Data
Imports System.Linq
Imports System.Web.UI.WebControls
Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Drawing.Color

'This is the management side of the program. It provides all the tools for the managers at
'LHR/250 to roster the engineers' shifts, set training dates and general adminstrative functions
'There is also an engineer side, which effectively runs off of the information set up by the management form

'Need to put in some validation when the data in the Manage Engineers component is written out. Add later.
'structure to hold the shift minimums
Structure tShiftMinimums
    Dim shiftNumber As Integer
    Dim shiftName As String
End Structure

'structure for to hold qualification details
Public Structure tQualsDetails
    Dim qualID As Integer
    Dim qualName As String
    Dim qualRating As Decimal
End Structure

'structure to hold the engineer ID and the position in the roster
Public Structure tEngrIDAndRostPostn
    Dim engrID As Integer
    Dim engrRostPostn As Integer
End Structure

'structure to hold the engineer vacation details
Public Structure tEngrVacation
    Dim engrID As Integer
    Dim engrRostPosn As Integer
    Dim engrVactnID As Integer
    Dim engrVactnStartDate As Date
    Dim engrVactnEndDate As Date
    Dim engrVactnPrty As Integer
    Dim engrVactnPref As Integer
End Structure

'structure to hold the engineer TDY details
Public Structure tEngrTDY
    Dim engrID As Integer
    Dim engrRostPosn As Integer
    Dim engrTDYID As Integer
    Dim engrTDYStartDate As Date
    Dim engrTDYEndDate As Date
    Dim engrTDYLoc As String
End Structure

Public Class MgmtLEMSProcess
    'public declarations for the filename
    Public rosterMonth As String = MonthName(Month(Now), True)
    Public rosterYear As String = Year(Now)
    Public rosterFileName As String = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & rosterMonth & rosterYear & ".xls"
    'variables to hold the current roster month displayed
    Public advanceRosterMonth As Date = Now
    Public advanceRosterYear As Date = Now
    'variable to hold the total number of engineers
    Public totalCountEngrs As Integer
    'variable to determine if the daily sheet has been saved once already
    Public dsFirstSave As Boolean = False
    'dataset to hold the engineers that are working on the current day
    Public dsEngineersOnDay As New DataSet
    'array to hold all th qualifications for updating in the settings
    Public qualsForSettings(200) As tQualsDetails
    'public variable to hold the engr ID and their roster position
    Public engrIDRosterPostn(200) As tEngrIDAndRostPostn
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
    'minimum number of engineers
    Public minNoEngineers As Integer
    'total engineers on each workday
    Public totalsArray(32) As Integer

    Private Sub closingRoutine()
        'MgmtCdsSave()
        MessageBox.Show("Saving All Data", "LEMS Message: Saving Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub MgmtLEMSProcess_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'this routine was written with help from:
        'https://msdn.microsoft.com/en-us/library/system.windows.forms.form.formclosing(v=vs.110).aspx
        'need to check if the daily sheet email has been sent
        'i.e. a public boolean
        closingRoutine()
    End Sub

    Private Sub MgmtLEMSProcess_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Startup subroutines- pulling data in, checking files exist, setting dates
        Dim startUpFileName As String
        startUpFileName = rosterFileName
        'load the monthly hours
        createMonthlyHours()
        'get the surnames and IDs from the database
        surnameDataSelect()
        'load the saved shift colours
        loadShiftColours()
        'load the minimum number of engineers
        loadMinEngineers()
        'get the qualifications
        qualDataSelect()
        'set the dates for the current run of the program
        setDates()
        'check the roster file exists
        checkRosterFileExists(startUpFileName, CInt(Month(Now)))
        'set the date labels of the roster
        MgmtCreateRostDateLab.Text = MonthName(Month(advanceRosterMonth), False) & " " & advanceRosterYear.Year
        MgmtViewRostDateLab.Text = MgmtCreateRostDateLab.Text
        'fill the sign in sheets
        fillSignInSheetShifts()
        'get the engineers that are working
        selectEngrs()
        'count the number of engineers working
        countWorkDays()
        'load the vacation from the database
        loadVacation()
        'load the TDY from the database
        loadTDY()
        'fill the roster with the requests
        fillRosterWithApprovedRequests()
        'copy the shifts to the monthlyhours
        addShiftsToMonthlyHours()

        'subroutine/function to set up the daily sheet
        currentDailySheetFirstSave()
        'load the ds recipients

        'Startup UI commands, just cleans up the interface upon opening
        MgmtShiftSwapSignInPanel.Hide()
        MgmtMngEngrsAddEngrPanel.Hide()
        MgmtMngEngrsCompleteProfBut.Hide()
        MgmtMngEngrsDoneEditingBut.Hide()
        MgmtMngEngrsTDYPanel.Hide()
        CDST4AddEngrPanel.Hide()
        CDST3AddEngrPanel.Hide()
        disenableTextboxes()

        'makes the entry ID column not user editable
        CDST4DailySheetDGV.Columns(0).ReadOnly = True
        CDST3DailySheetDGV.Columns(0).ReadOnly = True

        'could use a number and then a case statement? More efficient?
        If mgmtMngEngrsBool = True Then
            'show manage engineers tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtMngEngrsTab
        End If
        If mgmtCreateRostBool = True Then
            'show create roster tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtCreateRostTab
        End If
        If mgmtViewRostBool = True Then
            'show view roster tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtVwRostTab
        End If
        If mgmtMonthlyHrsBool = True Then
            'show monthly hours tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtVwMonthHrsTab
        End If
        If mgmtSignSheetBool = True Then
            'show print sign-in sheet tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtPrtSignSheetTab
        End If
        If mgmtCurrentDSBool = True Then
            'show current daily sheet tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtCDSTab
        End If
        If mgmtDSArchiveBool = True Then
            'show daily sheet archive tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtDSArchiveTab
        End If
        If mgmtSettBool = True Then
            'show settings tab on opening
            Me.MgmtControlTab.SelectedTab = MgmtSettTab
        End If

    End Sub

    Private Sub setDates()
        Dim todayDate As Date

        todayDate = Now.Date

        MgmtSignSheetDateTodayLab.Text = todayDate.ToString("D", CultureInfo.CreateSpecificCulture("en-US"))
        MgmtDSDateLab.Text = todayDate.ToString("D", CultureInfo.CreateSpecificCulture("en-US"))
    End Sub

    Private Sub checkRosterFileExists(ByVal inRosterName As String, ByVal inMonth As Integer)

        If File.Exists(inRosterName) Then
            'populate datagrid views from Excel file
            importFromExcelFile(inRosterName)
        Else
            CreateRoster(inMonth)
            exportRosterToExcel(inRosterName)
            fillRosterWithApprovedRequests()
        End If
    End Sub

    Private Sub MgmtHomeTab_Click(sender As Object, e As EventArgs) Handles MgmtHomeTab.Click
        MgmtLEMSMain.Show()
        mgmtMngEngrsBool = False
        mgmtCreateRostBool = False
        mgmtViewRostBool = False
        mgmtMonthlyHrsBool = False
        mgmtSignSheetBool = False
        mgmtCurrentDSBool = False
        mgmtDSArchiveBool = False
        mgmtSettBool = False
        Me.Close()
    End Sub

    Private Sub surnameDataSelect()
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim ds As New DataSet
        Dim dAdap As New OleDbDataAdapter
        Dim dAdapAllQuals As New OleDbDataAdapter
        Dim sql As String
        Dim surnames() As String
        Dim localEngrIDs() As Integer

        sql = "SELECT [Engineers].[Engineer_ID], [Engineers].[Surname], [Engineers].[Total_Qual_Score] FROM [Engineers] ORDER BY [Engineer_ID];"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdap = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdap.Fill(ds, "surnamesList")
        dAdap.Fill(dsEngineersOnDay, "Engineers")

        surnames = (From myRow In ds.Tables(0).AsEnumerable Select myRow.Field(Of String)("Surname")).ToArray
        'engineerSurnamesOnDay = (From myRow In ds.Tables(1).AsEnumerable Select myRow.Field(Of String)("Surname")).ToArray
        'engineerIDsOnDay = (From myRow In ds.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Engineer_ID")).ToArray
        localEngrIDs = (From myRow In ds.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Engineer_ID")).ToArray
        connec.Close()
        totalCountEngrs = surnames.Length
        'Dim surnames(ds.Tables(0).Rows.Count - 1) As String

        'For counter = 0 To ds.Tables(0).Rows.Count - 1
        '    surnames(counter) = ds.Tables(0).Rows(counter)("Surname").ToString
        'Next

        'fills all of the data grid views in the program with the surnames from the shift roster algorithm
        'this is done on start up in subroutines branched from startup so that the data appears when
        'the form is opened
        fillListBoxSurnames(surnames)
        Dim tempTotal As Integer = surnames.Length
        'sets the default overtime cells in the sign in sheet view
        For counter = 0 To surnames.Length - 1
            MgmtEngrSignInSheetDGV.Rows(counter).Cells(3).Value = "NO"
        Next

        'fill the structure with the IDs
        ReDim engrIDRosterPostn(localEngrIDs.Length)
        For counter = 0 To localEngrIDs.Length - 1
            engrIDRosterPostn(counter).engrID = localEngrIDs(counter)
            engrIDRosterPostn(counter).engrRostPostn = counter
        Next

        'MsgBox(ds.Tables(0).Rows(0)("Surname").ToString)
        'MsgBox(surnames(0).ToString)


    End Sub
    Public dsQualsCompareWriteBack As New DataSet
    Private Sub qualDataSelect()
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim ds As New DataSet
        Dim dAdapAllQuals As New OleDbDataAdapter
        Dim sql As String
        Dim localQualsIDArray() As Integer
        Dim localQualsRatingArray() As Decimal

        'get the qualifications from the database
        sql = "SELECT * FROM [Qualifications]"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdapAllQuals = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapAllQuals.Fill(dsQualsCompareWriteBack, "Qualifications")
        connec.Close()
        allQualsForListArray = (From myRow In dsQualsCompareWriteBack.Tables(0).AsEnumerable Select myRow.Field(Of String)("Qualification_Name")).ToArray
        localQualsIDArray = (From myRow In dsQualsCompareWriteBack.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("QUAL_ID")).ToArray()
        localQualsRatingArray = (From myRow In dsQualsCompareWriteBack.Tables(0).AsEnumerable Select myRow.Field(Of Decimal)("Qualification_Rating")).ToArray()


        'fill the structure
        For counter = 0 To allQualsForListArray.Length - 1
            qualsForSettings(counter).qualID = localQualsIDArray(counter)
            qualsForSettings(counter).qualName = allQualsForListArray(counter)
            qualsForSettings(counter).qualRating = localQualsRatingArray(counter)
        Next

        'redim the structure so that the lisbox referencing funtions correctly
        ReDim Preserve qualsForSettings(allQualsForListArray.Length - 1)

        'fill the listboxes
        fillQualsListBoxes()

    End Sub

    Private Sub fillListBoxSurnames(ByRef surnames() As String)
        'print all the surnames to the relevent listboxes

        For counter = 0 To (surnames.Length) - 1
            EngrRosterSpace.Rows.Add(surnames(counter))
            MgmtViewEngrsRostSpace.Rows.Add(surnames(counter))
            'add the position of the engineer to the public structure
            engrIDRosterPostn(counter).engrRostPostn = counter
            MgmtEngrMonthlyHoursDGV.Rows.Add(surnames(counter))
            MgmtEngrSignInSheetDGV.Rows.Add(surnames(counter))
            MgmtMngEngrsListBox.Items.Add(surnames(counter))
            MgmtSwapEngrsListBox.Items.Add(surnames(counter))
            MgmtEngrListAdvSearchDSArchiveCB.Items.Add(surnames(counter))
            MgmtEngrTrngLB.Items.Add(surnames(counter))
        Next

    End Sub

    Private Sub fillQualsListBoxes()
        'fill the qualifications listboxes
        For counter = 0 To allQualsForListArray.Length - 1
            MgmtMngEngrsAllQualsList.Items.Add(allQualsForListArray(counter))
            MgmtSettAllQualsLB.Items.Add(allQualsForListArray(counter))
        Next
    End Sub

    'Private Sub TempCreate_Click(sender As Object, e As EventArgs) Handles TempCreate.Click
    '    CreateRoster()
    '    'what needs to happen is when the form loads, the program checks to see if the 
    '    'Excel file in the format ShiftRosterFor(MONTH)(YEAR) exists, comparing it to the current month and year
    '    'If the file does not exist, then a new roster is created
    '    'If the file does exist then the Excel file is backloaded to the datagridview
    '    'The file system allows for each datagridview to be populated from the same source
    'End Sub

    Private Sub CreateRoster(ByVal inMonth As Integer)
        Dim testArray(31) As String
        Dim checkCount As Integer
        Dim arrayStart As Integer
        Dim recordStart As Integer = 0
        Dim rowCount As Integer = totalCountEngrs
        Dim colCount As Integer

        'defines the number of columns to count to
        colCount = System.DateTime.DaysInMonth(Year(Now), Month(Now))

        'fills the array with the off days and working days
        'the four off days
        For counter = 19 To 22
            testArray(counter) = "R"
        Next
        'the work days, skipping three leaving a blank space for the group of three off days
        For counter = 18 To 1 Step -1
            testArray(counter) = "X"
            checkCount = checkCount + 1
            If checkCount = 4 Then
                checkCount = 0
                'the 'skip' value can be changed in the settings (need to update this value to a variable)
                counter = counter - 3 '<-- this value
            End If
        Next
        'fills the gaps with R to represent an off day
        For counter = 1 To 22
            If testArray(counter) = Nothing Then
                testArray(counter) = "R"
            End If
        Next
        'fills the Early Early (EE) shift
        For counter = 1 To 6
            If testArray(counter) = "X" Then
                testArray(counter) = "EE"
            End If
        Next
        'fills the Early Shift (EN, Early Normal)
        For counter = 6 To 12
            If testArray(counter) = "X" Then
                testArray(counter) = "EN"
            End If
        Next
        'fills the Late Shift (LN, Late Normal)
        For counter = 12 To 18
            If testArray(counter) = "X" Then
                testArray(counter) = "LN"
            End If
        Next

        Dim arrayCounter As Integer
        Dim referencePointer As Integer
        Dim getCharFromArray As String

        arrayStart = (inMonth)
        'this algorithm fills the datagridview, column by column, row by row
        For rowCounter = 0 To rowCount - 1
            If arrayStart > 22 Then
                arrayStart = 0
            End If
            'sets the start pointer of the array
            recordStart = 22 - arrayStart
            arrayCounter = recordStart
            'need to change the reference pointer to an offset based on the month (month number -1)
            'i.e. January offset = 0, February offset = 1 etc
            referencePointer = arrayStart
            'points to the array value to fill the datagridview
            For columnCounter = colCount To 1 Step -1
                If arrayCounter = 0 Then
                    arrayCounter = 22
                End If
                getCharFromArray = testArray(arrayCounter)
                arrayCounter = arrayCounter - 1
                EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = getCharFromArray
            Next
            'moves the start of the array across 4 to repeat the pattern for the next row
            arrayStart = arrayStart + 4
        Next

        countWorkDays()
        engrShiftTotals()
        colorDataGrids(rowCount, colCount)

    End Sub

    Public Sub exportRosterToExcel(ByVal inRosterFilename As String)
        'the following code was adapted from:
        'http://vb.net-informations.com/excel-2007/vb.net_excel_2007_create_file.htm
        'and
        'http://stackoverflow.com/questions/6983141/changing-cell-color-of-excel-sheet-via-vb-net

        'Excel file variables
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlworkBooks As Excel.Workbooks
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim misValue As Object = System.Reflection.Missing.Value
        'my variables
        Dim colCount As Integer

        'my code
        colCount = System.DateTime.DaysInMonth(Year(Now), Month(Now))

        'copied code, sets up the Excel file
        xlApp.DisplayAlerts = False
        xlworkBooks = xlApp.Workbooks
        xlWorkBook = xlworkBooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        
        'below is my code
        'this sets the rows headers
        xlWorkSheet.Cells(1, 1) = "Name"
        For counter = 1 To colCount
            xlWorkSheet.Cells(1, counter + 1) = counter
        Next
        'this writes the data to the file from the datagridview
        'also writes the colour of the cell to the file
        For rowCounter = 0 To EngrRosterSpace.RowCount - 2
            For columnCounter = colCount To 1 Step -1
                xlWorkSheet.Cells(rowCounter + 2, columnCounter + 1) = EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value
                xlWorkSheet.Range(xlWorkSheet.Cells(rowCounter + 2, columnCounter + 1), xlWorkSheet.Cells(rowCounter + 2, columnCounter + 1)).Interior.Color = System.Drawing.ColorTranslator.ToOle(EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor)
            Next
        Next
        For rowCounter = 0 To EngrRosterSpace.RowCount - 2
            xlWorkSheet.Cells(rowCounter + 2, 1) = EngrRosterSpace.Rows(rowCounter).Cells(0).Value
        Next

        'this is the copied code
        xlWorkBook.SaveAs(inRosterFilename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)

        xlWorkBook.Close(True, misValue, misValue)

        'this closes the application as it releases COM objects
        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlworkBooks)
        xlApp.Quit()
        releaseObject(xlApp)

        MsgBox("Roster was saved as " & inRosterFilename)

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

    Public Sub fillDataGridViews(ByVal noOfColumns As Integer, ByVal noOfRows As Integer, ByRef inArray(,) As Object)
        'the purpose of this subroutine is to provide reusable code to fill every relevent datagridview in the management form
        For rowCounter = 2 To noOfRows
            For columnCounter = 2 To noOfColumns
                EngrRosterSpace.Rows(rowCounter - 2).Cells(columnCounter - 1).Value = inArray(rowCounter, columnCounter)
                MgmtViewEngrsRostSpace.Rows(rowCounter - 2).Cells(columnCounter - 1).Value = inArray(rowCounter, columnCounter)
            Next
        Next

        noOfColumns = System.DateTime.DaysInMonth(Year(Now), Month(Now))
        colorDataGrids(noOfRows, noOfColumns)

    End Sub

    Public Sub colorDataGrids(ByVal rowCount As Integer, ByVal colCount As Integer)
        'this routine fills the roster with the colours as set by the user
        Dim localString As String

        'fills the areas in the datagridview where off days are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "R" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(restColour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "R" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(restColour)
                End If
            Next
        Next
        'fills the areas in the datagridview where EE shifts are present
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "EE" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(shift1Colour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "EE" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(shift1Colour)
                End If
            Next
        Next
        'fills the areas in the datagridview where EN shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "EN" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(shift2Colour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "EN" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(shift2Colour)
                End If
            Next
        Next
        'fills the areas in the datagridview where LN shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "LN" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(shift3Colour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "LN" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(shift3Colour)
                End If
            Next
        Next
        'fills the areas in the datagridview where TIC shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "TDY" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(tdyColour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "TDY" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(tdyColour)
                End If
            Next
        Next
        'fills the areas in the datagridview where vacation shifts are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "V" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(vacationColour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "V" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(vacationColour)
                End If
            Next
        Next
        'fills the areas in the datagridview where planned sick days are present 
        For rowCounter = 0 To rowCount - 1
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                If EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "Sick" Or EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value = "MED" Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(sickColour)
                End If
                If MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "Sick" Or MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Value = "MED" Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(sickColour)
                End If
            Next
        Next
        'fills the areas in the datagridview where training is present 
        For rowCounter = 0 To rowCount - 2
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                localString = EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value.ToString
                If localString.Contains("TRNG") Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(trainingColour)
                End If
                If localString.Contains("TRNG") Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(trainingColour)
                End If
            Next
        Next
        'fills the areas in the datagridview where TDY is present 
        For rowCounter = 0 To rowCount - 2
            For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
                localString = EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value.ToString
                If localString.Contains("TDY") Then
                    EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(tdyColour)
                End If
                If localString.Contains("TDY") Then
                    MgmtViewEngrsRostSpace.Rows(rowCounter).Cells(columnCounter).Style.BackColor = ColorTranslator.FromHtml(tdyColour)
                End If
            Next
        Next
        
    End Sub

    Public Sub fillSignInSheetShifts()
        Dim shiftsOnDayArray As New List(Of String)() 'this had to be used as a regular dynamic string array caused a nullexceptionerror at runtime
        Dim currentDay As Integer

        'gets the current day to reference the shift roster
        currentDay = Now.Day

        'puts all the current shifts into a list (see comment next to declaration as to why a list has been used
        For counter = 0 To MgmtViewEngrsRostSpace.Rows.Count - 1
            shiftsOnDayArray.Add(MgmtViewEngrsRostSpace.Rows(counter).Cells(currentDay).Value)
        Next

        'inserts the elements of the list into the datagridview
        For counter = 0 To MgmtEngrSignInSheetDGV.Rows.Count - 1
            MgmtEngrSignInSheetDGV.Rows(counter).Cells(1).Value = shiftsOnDayArray(counter)
        Next

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click

    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click

    End Sub

    Private Sub EditSelectedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditSelectedToolStripMenuItem.Click
        MgmtShiftSwapSignInPanel.Show()

    End Sub

    Private Sub LogoutButton_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LogoutButton.LinkClicked
        'resets all the values for the homepage buttons
        LoginPage.Show()
        mgmtMngEngrsBool = False
        mgmtCreateRostBool = False
        mgmtViewRostBool = False
        mgmtMonthlyHrsBool = False
        mgmtSignSheetBool = False
        mgmtCurrentDSBool = False
        mgmtDSArchiveBool = False
        mgmtSettBool = False
        'check to see if the daily sheet email has been sent
        'see closing routine at top of page
        closingRoutine()
        Me.Close()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked

    End Sub

    Private Sub MgmtItemDSArchiveSearch_Click(sender As Object, e As EventArgs) Handles MgmtItemDSArchiveSearchTB.Click
        MgmtItemDSArchiveSearchTB.Clear()
    End Sub

    Private Sub MgmtAircraftNoDSArchiveSearch_Click(sender As Object, e As EventArgs) Handles MgmtAircraftNoDSArchiveSearchTB.Click
        MgmtAircraftNoDSArchiveSearchTB.Clear()
    End Sub

    Private Sub MgmtGeneralResultAdvSearchDSArchiveRB_CheckedChanged(sender As Object, e As EventArgs) Handles MgmtGeneralResultAdvSearchDSArchiveRB.CheckedChanged
        If MgmtGeneralResultAdvSearchDSArchiveRB.Checked Then
            MgmtMetricsAdvSearchDSArchivePanel.Show()
            MgmtDSAAxAcroPDF1.Show()
        End If
    End Sub

    Private Sub MgmtBySheetAdvSearchDSArchiveRB_CheckedChanged(sender As Object, e As EventArgs) Handles MgmtBySheetAdvSearchDSArchiveRB.CheckedChanged
        If MgmtBySheetAdvSearchDSArchiveRB.Checked Then
            MgmtDSAAxAcroPDF1.Show()
            MgmtMetricsAdvSearchDSArchivePanel.Hide()
        End If
    End Sub

    Private Sub MgmtOverTimeCB_CheckedChanged(sender As Object, e As EventArgs) Handles MgmtOverTimeCB.CheckedChanged
        If MgmtOverTimeCB.Checked Then
            MgmtEngrSignInSheetDGV.CurrentRow.Cells(3).Value = "YES"
        Else
            MgmtEngrSignInSheetDGV.CurrentRow.Cells(3).Value = "NO"
        End If
    End Sub

    Private Sub selectEngrs()
        'Dim arrayLength As Integer

        'get all engineers, their ids and their quals from the db and put in structure
        'delete from structure when delete from sign in sheet
        'copy over structure to the public class

        For counter = 0 To MgmtEngrSignInSheetDGV.Rows.Count - 1
            If counter = MgmtEngrSignInSheetDGV.Rows.Count Then
                Exit For
            End If
            If MgmtEngrSignInSheetDGV.Rows(counter).Cells(1).Value = "R" Or MgmtEngrSignInSheetDGV.Rows(counter).Cells(1).Value = "V" Then
                MgmtEngrSignInSheetDGV.Rows.Remove(MgmtEngrSignInSheetDGV.Rows(counter))
                'removes the engineers that are not working
                'this is for use in the public class so that the correct engineer scores are calculated
                dsEngineersOnDay.Tables(0).Rows(counter).Delete()
                counter = counter - 1
            End If
        Next

        'For counter = 0 To MgmtEngrSignInSheetDGV.Rows.Count - 1
        '    engineersOnDay(counter) = MgmtEngrSignInSheetDGV.Rows(counter).Cells(0).Value.ToString
        'Next

        'For counter = 0 To engineersOnDay.Length - 1
        '    If Not engineersOnDay(counter) = Nothing Then
        '        arrayLength = arrayLength + 1
        '    End If
        'Next

        'ReDim Preserve engineersOnDay(arrayLength)

    End Sub

    Private Sub countWorkDays()
        'counts the number of personnel on each day and puts the figures into a datagridview
        Dim arrayUsed As Integer
        Dim valueString As String

        'clear the array before counting
        ReDim totalsArray(32)

        'goes along the roster and checks the values for each engineer on each day
        For columnCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
            For rowCounter = 0 To totalCountEngrs - 1 'this needs to reference the length of the surnames array
                valueString = EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value.ToString
                If valueString = "EE" Or valueString = "EN" Or valueString = "LN" Or valueString = "OT" Or valueString = "TIC" Then
                    totalsArray(columnCounter - 1) = totalsArray(columnCounter - 1) + 1
                End If
            Next
        Next

        'redims the array to save space
        For counter = 1 To totalsArray.Length - 1
            If totalsArray(counter) <> Nothing Then
                arrayUsed = arrayUsed + 1
            End If
        Next
        ReDim Preserve totalsArray(arrayUsed)

        'adds a row to the dgv
        MgmtDayTotals.Rows.Add(1)

        'fills the dgv
        For counter = 0 To System.DateTime.DaysInMonth(Year(Now), Month(Now)) - 1
            MgmtDayTotals.Rows(0).Cells(counter).Value = totalsArray(counter)
        Next

        'goes to the routine to check numbers against the minimum
        checkNumbersOnRoster()

    End Sub

    Private Sub MgmtSignSheetPrintBut_Click_1(sender As Object, e As EventArgs) Handles MgmtSignSheetPrintBut.Click
        'Ths subroutine was adapted from KPY Notes 'Writing to Word'
        '[need to insert a reference!]

        'The subroutine exports the sign-in sheet to word, formats it in a Word table, shows a print preview
        'and then prints.

        Dim wordApp As Word.Application = New Microsoft.Office.Interop.Word.Application()
        Dim wordDoc As Word.Document
        Dim wordTable As Word.Table
        Dim wordPara As Word.Paragraph

        wordDoc = New Microsoft.Office.Interop.Word.Document

        wordDoc = wordApp.Documents.Add
        wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
        wordPara = wordDoc.Paragraphs.Add

        wordPara.Range.Font.Name = "Times New Roman"
        wordPara.Range.Text = "SIGN IN SHEET FOR " & Now.Date.ToString("D", CultureInfo.CreateSpecificCulture("en-US"))
        wordPara.Range.Style = "Heading 1"
        wordPara.Range.InsertParagraphBefore()
        wordPara.Range.Style = "Normal"

        wordTable = wordDoc.Tables.Add(wordDoc.Bookmarks("\endofdoc").Range, MgmtEngrSignInSheetDGV.Rows.Count, MgmtEngrSignInSheetDGV.Columns.Count)

        For columnCounter = 1 To MgmtEngrSignInSheetDGV.Columns.Count
            wordTable.Cell(1, columnCounter).Range.Text = MgmtEngrSignInSheetDGV.Columns(columnCounter - 1).HeaderText
        Next

        For rowCount = 1 To MgmtEngrSignInSheetDGV.Rows.Count - 1
            For columnCounter = 1 To MgmtEngrSignInSheetDGV.Columns.Count - 1
                wordTable.Cell(rowCount + 1, columnCounter).Range.Text = MgmtEngrSignInSheetDGV.Rows(rowCount - 1).Cells(columnCounter - 1).Value
            Next
        Next

        wordDoc.Tables(1).Borders.Enable = True

        wordApp.PrintPreview = True
        'wordApp.PrintOut()

        wordApp.Quit()
    End Sub

    Private Sub MgmtEngrAddTrngBut_Click(sender As Object, e As EventArgs) Handles MgmtEngrAddTrngBut.Click
        'this routine adds the training to the roster for the dates selected by the user
        Dim timeOffTypes(50) As String
        Dim fromDate As Date
        Dim toDate As Date

        'copy the values from the datetimepickers to the local variables
        fromDate = MgmtEngrDateFromTrngDTP.Value
        toDate = MgmtEngrDateToTrngDTP.Value

        'validation to ensure that acceptable dates are entered
        If toDate < fromDate Or fromDate < Now.Date Or toDate < Now.Date Then
            MessageBox.Show("The end date cannot be before the start date, enter another", "LEMS Warning: Date Error", MessageBoxButtons.OK)
        Else
            'checks to see that an index from the list has been selected
            If MgmtEngrTrngLB.SelectedIndex = -1 Then
                MessageBox.Show("An engineer was not selected. For a change, an engineer must be selected", "LEMS Warning: Engineer Selection Error", MessageBoxButtons.OK)
            ElseIf MgmtEngrTrngLctnCB.SelectedIndex = -1 Then
                MessageBox.Show("A location was not selected. For a change, an type must be selected", "LEMS Warning: Location Selection Error", MessageBoxButtons.OK)
            ElseIf MgmtEngrTrngTypeCB.SelectedIndex = -1 Then
                MessageBox.Show("A type was not selected. For a change, an type must be selected", "LEMS Warning: Training Type Selection Error", MessageBoxButtons.OK)
            Else

                'pass over to the filling subroutine
                addTraining(fromDate, toDate, MgmtEngrTrngLB.SelectedIndex)

                'rejig the rosters if the numbers fall short of the minimum
                For counter = fromDate.Day To toDate.Day
                    If totalsArray(counter) < minNoEngineers Then
                        'call the subroutine to rejig the roster 
                        rearrangeRosterChanges(counter)
                    End If
                Next

                'recalculate the numbers and recolour the roster
                engrShiftTotals()
                colorDataGrids(EngrRosterSpace.Rows.Count, EngrRosterSpace.Columns.Count)
                countWorkDays()

                'clean up, reset dates and values
                MgmtEngrTrngLB.SelectedIndex = -1
                MgmtEngrTrngLctnCB.SelectedIndex = -1
                MgmtEngrTrngTypeCB.SelectedIndex = -1
                MgmtEngrDateFromTrngDTP.Value = Date.Now()
                MgmtEngrDateToTrngDTP.Value = Date.Now()

            End If
        End If
        

    End Sub

    Private Sub addTraining(ByVal inStartDate As Date, ByVal inEndDate As Date, ByVal engrPostn As Integer)
        'this subroutine adds the training to the roster

        'total days of vacation
        Dim totalDays As Integer
        Dim manipulateDays As Integer

        'variables to hold the differences between the months and years
        Dim monthDifference As Integer
        Dim monthAdvance As Integer
        Dim remainingMonths As Integer

        'create the array to analyse 
        Dim workingDaysArray(totalDays) As Boolean

        'control variables
        Dim arrayCount As Integer
        Dim localFileName As String
        Dim checkingCount As Integer

        totalDays = (inEndDate - inStartDate).Days
        manipulateDays = totalDays
        'gets the number of months between the two dates
        monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))

        'check the roster exists for the month regarding the vacation
        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
        checkRosterFileExists(localFileName, CInt(inStartDate.Month))

        'analyse the boolean array
        'different routines depending if the vacation is in the same month
        arrayCount = 0
        If inStartDate.Month = inEndDate.Month Then
            For counter = inStartDate.Day To inEndDate.Day
                EngrRosterSpace.Rows(engrPostn).Cells(counter).Value = "TRNG " & MgmtEngrTrngTypeCB.SelectedItem & " at " & MgmtEngrTrngLctnCB.SelectedItem
                arrayCount = arrayCount + 1
            Next
            exportRosterToExcel(localFileName)
        End If

        'separate if statements to simplify code layout
        'checks the workdays between non-contiguous months in the same year
        If inStartDate.Month <> inEndDate.Month And inStartDate.Year = inEndDate.Year Then
            monthDifference = inEndDate.Month - inStartDate.Month
            'find if the minimum shift numbers have been met for the first month
            For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                EngrRosterSpace.Rows(engrPostn).Cells(innerCounter).Value = "TRNG " & MgmtEngrTrngTypeCB.SelectedItem & " at " & MgmtEngrTrngLctnCB.SelectedItem
                arrayCount = arrayCount + 1
                manipulateDays = manipulateDays - 1
            Next
            localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
            exportRosterToExcel(localFileName)
            For counter = 1 To monthDifference
                'check the roster exists for the months regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month + counter, True) & inStartDate.Year & ".xls"
                checkRosterFileExists(localFileName, CInt(inStartDate.Month))
                'loop for the next months
                checkingCount = 0
                For innerCounter = 0 To manipulateDays
                    checkingCount = checkingCount + 1
                    If checkingCount > System.DateTime.DaysInMonth(inStartDate.Year, CInt(inStartDate.Month + counter)) Then
                        checkingCount = 1
                        Exit For
                    End If
                    EngrRosterSpace.Rows(engrPostn).Cells(checkingCount).Value = "TRNG " & MgmtEngrTrngTypeCB.SelectedItem & " at " & MgmtEngrTrngLctnCB.SelectedItem
                    arrayCount = arrayCount + 1
                    manipulateDays = manipulateDays - 1
                    If innerCounter > totalDays Then
                        'exit this for loop to move to the next month
                        Exit For
                    End If
                Next
                exportRosterToExcel(localFileName)
            Next
        End If

        'checks the workdays between non-contigous years
        If inStartDate.Year <> inEndDate.Year Then
            'get the differences between the vacation
            monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))
            'check the first month
            For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                EngrRosterSpace.Rows(engrPostn).Cells(innerCounter).Value = "TRNG " & MgmtEngrTrngTypeCB.SelectedItem & " at " & MgmtEngrTrngLctnCB.SelectedItem
                arrayCount = arrayCount + 1
                'manipulateDays = manipulateDays - 1
            Next
            localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
            exportRosterToExcel(localFileName)
            'continue the loop until the month advance is greater than 12
            remainingMonths = monthDifference
            For monthCounter = 1 To monthDifference
                monthAdvance = inStartDate.Month + monthCounter
                'resets the months when the count is greater than 12
                'this signifies that the new year analysis has been started
                If monthAdvance > 12 Then
                    Exit For
                End If
                'check the roster exists for the months regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & inStartDate.Year & ".xls"
                checkRosterFileExists(localFileName, monthAdvance)
                'loop for the next months
                For innerCounter = 0 To manipulateDays - 1
                    EngrRosterSpace.Rows(engrPostn).Cells(innerCounter).Value = "TRNG " & MgmtEngrTrngTypeCB.SelectedItem & " at " & MgmtEngrTrngLctnCB.SelectedItem
                    arrayCount = arrayCount + 1
                    'manipulateDays = manipulateDays - 1
                    If innerCounter > totalDays Then
                        'exit this for loop to move to the next month
                        Exit For
                    End If
                    remainingMonths = remainingMonths - 1
                Next

                exportRosterToExcel(localFileName)
            Next

            'go to the next year
            For monthCounter = 1 To remainingMonths
                monthAdvance = inStartDate.Month + monthCounter
                If monthAdvance > 12 Then
                    monthAdvance = monthCounter
                End If
                'check the roster exists for the months regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & (inEndDate.Year) & ".xls"
                checkRosterFileExists(localFileName, monthAdvance)
                'loop for the next months
                For innerCounter = 0 To manipulateDays - 1
                    EngrRosterSpace.Rows(engrPostn).Cells(innerCounter).Value = "TRNG " & MgmtEngrTrngTypeCB.SelectedItem & " at " & MgmtEngrTrngLctnCB.SelectedItem
                    arrayCount = arrayCount + 1
                    'manipulateDays = manipulateDays - 1
                    If innerCounter > totalDays Then
                        'exit this for loop to move to the next month
                        Exit For
                    End If
                Next
                exportRosterToExcel(localFileName)
            Next

        End If

        importFromExcelFile(rosterFileName)
        addShiftsToMonthlyHours()
    End Sub

    Private Sub MgmtEngrAddOtherTimeOffBut_Click(sender As Object, e As EventArgs) Handles MgmtEngrAddOtherTimeOffBut.Click
        'This subroutine is required so that changes can be made manually to the shift roster
        Dim selectionCounter As Integer
        Dim timeOffTypes(50) As String
        Dim fromDay As Integer
        Dim endDay As Integer
        Dim localFilename As String

        'This subroutine is required so that changes can be made manually to the shift roster


        selectionCounter = MgmtEngrTrngLB.SelectedIndex

        fromDay = CInt(MgmtEngrOtherTimeOffFromDTP.Value.Day)
        endDay = CInt(MgmtEngrOthertimeOffToDTP.Value.Day)

        'checks to see if the end date is before the start date and flags invalid if it is
        If endDay < fromDay Then
            MessageBox.Show("The end date is before the start date, enter another", "LEMS Warning: Date Error", MessageBoxButtons.OK)
        End If

        'this checks to see if an item has been selected from the lists
        If selectionCounter = -1 Then
            MessageBox.Show("An engineer was not selected. For a change, an engineer must be selected", "LEMS Warning: Engineer Selection Error", MessageBoxButtons.OK)
        ElseIf MgmtEngrTimeOffTypeCB.SelectedIndex = -1 Then
            MessageBox.Show("A type was not selected. For a change, an type must be selected", "LEMS Warning: Time-Off Type Selection Error", MessageBoxButtons.OK)
        ElseIf MgmtEngrTimeOffTypeCB.SelectedItem = "Spare Vacation Day" And totalsArray(fromDay - 1) + minNoEngineers Then
            MessageBox.Show("Spare vacation cannot be added to days where the numbers are at minimum", "LEMS Warning: Minimum Number of Engineers Error", MessageBoxButtons.OK)
            'this inserts the change into the roster
            For counter = fromDay To endDay
                If MgmtEngrTimeOffTypeCB.SelectedItem = "Spare Vacation Day" Then
                    EngrRosterSpace.Rows(selectionCounter).Cells(counter).Value = "V"
                Else
                    EngrRosterSpace.Rows(selectionCounter).Cells(counter).Value = MgmtEngrTimeOffTypeCB.SelectedItem
                End If
            Next
        End If

        'recount the number of engineers for each day
        countWorkDays()

        'call a subroutine to rejig the roster
        selectionCounter = 0
        selectionCounter = 0
        checkNumbersOnRoster()
        engrShiftTotals()
        colorDataGrids(EngrRosterSpace.Rows.Count, EngrRosterSpace.Columns.Count)

        're-generate the monthly hours
        addShiftsToMonthlyHours()

        'generate the new shift totals
        engrShiftTotals()

        'export to excel
        localFilename = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(MgmtEngrOtherTimeOffFromDTP.Value.Month, True) & MgmtEngrOtherTimeOffFromDTP.Value.Year & ".xls"
        exportRosterToExcel(localFilename)

    End Sub

    Sub checkNumbersOnRoster()
        'this  routine checks if the roster contains any days where the
        'number of engineers is at or below the minimum as set by the user

        For counter = 0 To totalsArray.Length - 1
            'need to set 6 to a value that is in the settings window
            If totalsArray(counter) < minNoEngineers Then
                MessageBox.Show("There are a number of personnel deficiencies. LEMS will rearrange the shift pattern.", "LEMS Message: Not Enough Engineers on Duty", MessageBoxButtons.OK, MessageBoxIcon.Information)
                rearrangeRosterChanges(counter + 1)
            End If
        Next

    End Sub

    Sub rearrangeRosterChanges(ByVal inReference As Integer)
        'needs a few parameters passed
        'this subroutine uses an almost first fit algorithm
        'but is amended slightly to take into account inconvenience
        'due to the position of off days
        Dim nextOffCheck As Boolean
        Dim nextOnCheck As Boolean
        'declaring a structure so that the corresponding shift depending on the deficiency
        'can be found
        Dim shiftLacks(3) As tShiftMinimums
        'these variables store the running totals 
        Dim eeCount As Integer
        Dim enCount As Integer
        Dim lnCount As Integer

        Dim tempMin As Integer
        Dim arrayPoint As Integer

        For rowCounter = 0 To EngrRosterSpace.Rows.Count - 2
            Select Case EngrRosterSpace.Rows(rowCounter).Cells(inReference).Value
                Case Is = "EE"
                    eeCount = eeCount + 1
                Case Is = "EN"
                    enCount = enCount + 1
                Case Is = "LN"
                    lnCount = lnCount + 1
            End Select
        Next


        'this algorithm finds the numbers of each shift on that columnn
        'then finds the least of those shifts
        'and replaces the off day instance with the one that is lacking

        shiftLacks(1).shiftNumber = eeCount
        shiftLacks(1).shiftName = "EE"
        shiftLacks(2).shiftNumber = enCount
        shiftLacks(2).shiftName = "EN"
        shiftLacks(3).shiftNumber = lnCount
        shiftLacks(3).shiftName = "LN"

        tempMin = Math.Min(shiftLacks(1).shiftNumber, shiftLacks(2).shiftNumber)
        tempMin = Math.Min(tempMin, shiftLacks(3).shiftNumber)

        For counter = 1 To 3
            If shiftLacks(counter).shiftNumber = tempMin Then
                arrayPoint = counter
            End If
        Next

        For rowCounter = 0 To EngrRosterSpace.Rows.Count - 1
            If inReference + 1 > System.DateTime.DaysInMonth(Year(Now), Month(Now)) Then
                Exit For
            End If
            If EngrRosterSpace.Rows(rowCounter).Cells(inReference).Value = "R" And EngrRosterSpace.Rows(rowCounter).Cells(inReference + 1).Value <> "R" And EngrRosterSpace.Rows(rowCounter).Cells(inReference + 1).Value <> "Sick" Then
                nextOnCheck = True
                If nextOnCheck = True Then
                    EngrRosterSpace.Rows(rowCounter).Cells(inReference).Value = shiftLacks(arrayPoint).shiftName
                    exportRosterToExcel(rosterFileName)
                    Exit For
                End If
            ElseIf EngrRosterSpace.Rows(rowCounter).Cells(inReference).Value <> "R" And EngrRosterSpace.Rows(rowCounter).Cells(inReference + 1).Value <> "Sick" And EngrRosterSpace.Rows(rowCounter).Cells(inReference + 1).Value = "R" Then
                nextOffCheck = True
                If nextOffCheck = True Then
                    EngrRosterSpace.Rows(rowCounter).Cells(inReference).Value = shiftLacks(arrayPoint).shiftName
                    exportRosterToExcel(rosterFileName)
                    Exit For
                End If
            End If
        Next

        'import the roster after the changes have been made
        importFromExcelFile(rosterFileName)

    End Sub

    Private Sub MgmtEngrShiftRostNextMonthBut_Click(sender As Object, e As EventArgs) Handles MgmtEngrShiftRostNextMonthBut.Click
        nextRosterButClick()
    End Sub

    Public Sub nextRosterButClick()
        'loads the next month's roster into the viewer
        Dim localFileName As String
        Dim localMonth As String
        Dim localMonthInt As Integer
        Dim localYear As Integer

        'alter the months
        advanceRosterMonth = advanceRosterMonth.AddMonths(1)
        localMonth = MonthName(Month(advanceRosterMonth), True)
        localMonthInt = CInt(Month(advanceRosterMonth))
        localYear = Year(Now)
        'alter the year should the new month be January
        If localMonth = "Jan" Then
            advanceRosterYear = advanceRosterYear.AddYears(1)
            localYear = Year(advanceRosterYear)
        End If
        'create a new filename
        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & localMonth & localYear & ".xls"
        'check the file exists
        checkRosterFileExists(localFileName, localMonthInt)
        MgmtCreateRostDateLab.Text = MonthName(Month(advanceRosterMonth), False) & " " & advanceRosterYear.Year
        MgmtViewRostDateLab.Text = MgmtCreateRostDateLab.Text
    End Sub

    Private Sub MgmtEngrShiftRostPrevMonthBut_Click(sender As Object, e As EventArgs) Handles MgmtEngrShiftRostPrevMonthBut.Click
        prevRosterButClick()
    End Sub

    Public Sub prevRosterButClick()
        'loads previous rosters
        Dim localFileName As String
        Dim localMonth As String
        Dim localMonthInt As Integer
        Dim localYear As Integer

        'alter the months
        advanceRosterMonth = advanceRosterMonth.AddMonths(-1)
        localMonth = MonthName(Month(advanceRosterMonth), True)
        localMonthInt = CInt(Month(advanceRosterMonth))
        'alter the year shuold the new month be December
        If localMonth = "Dec" Then
            advanceRosterYear = advanceRosterYear.AddYears(-1)
            localYear = Year(advanceRosterYear)
        Else
            localYear = advanceRosterYear.Year
        End If
        'create a new filename
        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & localMonth & localYear & ".xls"
        'check the file exists
        If localMonth = "Nov" And localYear = 2014 Then
            MessageBox.Show("Rosters before December 2014 cannot be created", "LEMS Error: Past roster limit reached", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            checkRosterFileExists(localFileName, localMonthInt)
            MgmtCreateRostDateLab.Text = MonthName(Month(advanceRosterMonth), False) & " " & advanceRosterYear.Year
            MgmtViewRostDateLab.Text = MgmtCreateRostDateLab.Text
        End If
    End Sub

    Public compareQualsArray() As String
    Public allQualsForListArray() As String
    Public qualsIDArray() As Integer
    Private Sub MgmtMngEngrsListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MgmtMngEngrsListBox.SelectedIndexChanged

        MgmtMngEngrsQualsListBox.Items.Clear()
        readInfoFromDB()

    End Sub

    Public engineerID As Integer

    Private Sub readInfoFromDB()
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dsEngineers As New DataSet
        Dim dsEmpNoAandP As New DataSet
        Dim dsEngrQuals As New DataSet
        Dim dsEngrTDY As New DataSet
        Dim dsEngrVactn As New DataSet
        Dim dsEngrReqs As New DataSet
        Dim dsEngrReqsPndng As New DataSet
        Dim dAdapEngrs As New OleDbDataAdapter
        Dim dAdapEmpNoAandPNo As New OleDbDataAdapter
        Dim dAdapQuals As New OleDbDataAdapter
        Dim dAdapTDY As New OleDbDataAdapter
        Dim dAdapVactn As New OleDbDataAdapter
        Dim dAdapAllQuals As New OleDbDataAdapter
        Dim dAdapRequests As New OleDbDataAdapter
        Dim dAdapRequestsPndng As New OleDbDataAdapter
        Dim sql As String
        Dim engineerID As Integer
        Dim selectedIndex As Integer

        'gets the selected item from the listox of engineers
        selectedIndex = MgmtMngEngrsListBox.SelectedIndex

        sql = "SELECT * FROM [Engineers] ORDER BY [Engineer_ID];"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdapEngrs = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapEngrs.Fill(dsEngineers, "Engineers")
        connec.Close()

        engineerID = CInt(dsEngineers.Tables(0).Rows(selectedIndex)("Engineer_ID"))

        'below are the sql statements and database operations to read in all the relevant data
        'the statement is defined as a string and then bound to a database command
        'a dataset is created and the data from the database is put through an adaptor
        'and into the dataset
        '--> Possible re-write with parameters

        sql = "SELECT [Employee_Number], [A&P_License_Number] FROM [Employee_and_A&P] WHERE [Engineer_ID] = " & engineerID.ToString & ";"
        cmd = New OleDbCommand(sql, connec)
        dAdapEmpNoAandPNo = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapEmpNoAandPNo.Fill(dsEmpNoAandP, "EmployeeNumbersAandP")
        connec.Close()

        'the operations below get all the relevant data using the relationships in the database
        sql = "SELECT [Qualifications].[QUAL_ID], [Qualifications].[Qualification_Name], [Qualifications].[Qualification_Rating] FROM [Engineers], [Qualifications], [EN/QUALS] WHERE [Engineers].[Engineer_ID] = " & engineerID.ToString & " AND [EN/QUALS].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/QUALS].[QUAL_ID] = [Qualifications].[QUAL_ID]"
        cmd = New OleDbCommand(sql, connec)
        dAdapQuals = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapQuals.Fill(dsEngrQuals, "Qualifications")
        connec.Close()
        compareQualsArray = (From myRow In dsEngrQuals.Tables(0).AsEnumerable Select myRow.Field(Of String)("Qualification_Name")).ToArray
        qualsIDArray = (From myRow In dsEngrQuals.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("QUAL_ID")).ToArray


        sql = "SELECT [TDY].[TDY_ID], [TDY].[TDY_Start_Date], [TDY].[TDY_End_Date] FROM [Engineers], [TDY], [EN/TDY] WHERE [Engineers].[Engineer_ID] = " & engineerID.ToString & " AND [EN/TDY].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/TDY].[TDY_ID] = [TDY].[TDY_ID]"
        cmd = New OleDbCommand(sql, connec)
        dAdapTDY = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapTDY.Fill(dsEngrTDY, "TDY")
        connec.Close()

        sql = "SELECT [Vacation].[Vacation_ID], [Vacation].[V_Start_Date], [Vacation].[V_End_Date] FROM [Engineers], [Vacation], [EN/VA] WHERE [Engineers].[Engineer_ID] = " & engineerID.ToString & " AND [EN/VA].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/VA].[VACTN_ID] = [Vacation].[Vacation_ID]"
        cmd = New OleDbCommand(sql, connec)
        dAdapVactn = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapVactn.Fill(dsEngrVactn, "Vacation")
        connec.Close()

        'gets the approved requests
        sql = "SELECT [Requests].[Request_Type], [Requests].[Request_Date], [Requests].[Request_Shift_Change_From], [Requests].[Request_Shift_Change_To], [Requests].[Request_Shift_Change_With_Engineer], [Requests].[Request_Description] FROM [Requests], [EN/REQS], [Engineers] WHERE [Engineers].[Engineer_ID] = " & engineerID.ToString & " AND [EN/REQS].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/REQS].[REQ_ID] = [Requests].[REQ_ID] AND [Requests].[Request_Approved] = TRUE"
        cmd = New OleDbCommand(sql, connec)
        dAdapRequests = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapRequests.Fill(dsEngrReqs, "Requests")
        connec.Close()
        'set the source of the datagridview
        MgmtMngEngrsApprvdRqstsDGV.DataSource = dsEngrReqs.Tables(0)

        'gets the pending requests
        sql = "SELECT * FROM [Requests], [EN/REQS], [Engineers] WHERE [Engineers].[Engineer_ID] = " & engineerID.ToString & " AND [EN/REQS].[EN_ID] = [Engineers].[Engineer_ID] AND [EN/REQS].[REQ_ID] = [Requests].[REQ_ID] AND [Requests].[Request_Approved] = FALSE"
        cmd = New OleDbCommand(sql, connec)
        dAdapRequestsPndng = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapRequestsPndng.Fill(dsEngrReqsPndng, "Requests")
        connec.Close()
        'set the source of the datagridview
        MgmtMngEngrsPendingRqstsDGV.DataSource = dsEngrReqsPndng.Tables(0)

        'this is used to get the qualification id when the data is written back. 
        'the code below was removed and put into a load routine because the the data is 'static'
        'sql = "SELECT [Qualifications].[QUAL_ID], [Qualifications].[Qualification_Name] FROM [Qualifications]"
        'cmd = New OleDbCommand(sql, connec)
        'dAdapAllQuals = New OleDbDataAdapter(sql, connec)
        'connec.Open()
        'dAdapAllQuals.Fill(dsQualsCompareWriteBack, "Qualifications")
        'connec.Close()
        'allQualsForListArray = (From myRow In dsQualsCompareWriteBack.Tables(0).AsEnumerable Select myRow.Field(Of String)("Qualification_Name")).ToArray

        fillMgmtEngrInfo(dsEngineers, dsEmpNoAandP, dsEngrQuals, dsEngrTDY, dsEngrVactn, selectedIndex)
    End Sub

    Private Sub fillMgmtEngrInfo(ByRef engrInfo As DataSet, ByRef engrEmpAPNos As DataSet, ByRef engrQuals As DataSet, ByRef engrTDY As DataSet, ByRef engrVactn As DataSet, ByVal selectedIndex As Integer)

        engineerID = CInt(engrInfo.Tables(0).Rows(selectedIndex)("Engineer_ID"))

        'this subroutine fills all the relevent sections in the view after the data has been imported
        'from the database
        'the program design was changed because instead of storing the image directly, the file path
        'is stored. This is more space efficient and allows the user to more easily select files

        'this small algorithm prevents the user editing the information in the textboxes
        'until they are ready to edit it
        For Each cntrl As Control In MgmtMngEngrsTab.Controls
            If TypeOf cntrl Is System.Windows.Forms.TextBox Or TypeOf cntrl Is System.Windows.Forms.MaskedTextBox Then
                cntrl.Enabled = False
            End If
        Next

        MgmtMngEngrsDoBMTB.Enabled = False
        MgmtMngEngrsHireDateMTB.Enabled = False
        MgmtEngrCanBeTICCB.Enabled = False

        MgmtEngrForeNameTB.Text = engrInfo.Tables(0).Rows(selectedIndex)("Forename").ToString()
        MgmtEngrSurnameTB.Text = engrInfo.Tables(0).Rows(selectedIndex)("Surname").ToString()
        MgmtMngEngrsDoBMTB.Text = engrInfo.Tables(0).Rows(selectedIndex)("DoB").ToString
        MgmtMngEngrsHireDateMTB.Text = engrInfo.Tables(0).Rows(selectedIndex)("Hire_Date").ToString
        MgmtMngEngrsProfileImgLoctnTB.Text = (engrInfo.Tables(0).Rows(selectedIndex)("Engineer_Image")).ToString
        If MgmtMngEngrsProfileImgLoctnTB.Text <> Nothing Then
            MgmtMngEngrsPicBox.Image = Bitmap.FromFile(MgmtMngEngrsProfileImgLoctnTB.Text)
        Else
            MgmtMngEngrsPicBox.Image = Nothing
        End If
        MgmtEngrENIDLab.Text = engrInfo.Tables(0).Rows(selectedIndex)("Engineer_ID").ToString()
        MgmtEngrCanBeTICCB.Checked = engrInfo.Tables(0).Rows(selectedIndex)("Can_Be_TIC")
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(0).Value = CInt(engrInfo.Tables(0).Rows(selectedIndex)("EE_Shift_Occurrence"))
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(1).Value = CInt(engrInfo.Tables(0).Rows(selectedIndex)("EN_Shift_Occurrence"))
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(2).Value = CInt(engrInfo.Tables(0).Rows(selectedIndex)("LN_Shift_Occurrence"))
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(3).Value = CInt(engrInfo.Tables(0).Rows(selectedIndex)("TIC_Occurrence"))
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(4).Value = CInt(engrInfo.Tables(0).Rows(selectedIndex)("Vacation_Days_Remaining"))
        MgmtEngrQualScoreLab.Text = engrInfo.Tables(0).Rows(selectedIndex)("Total_Qual_Score").ToString

        MgmtEngrEmpNoTB.Text = (engrEmpAPNos.Tables(0).Rows(0)("Employee_Number")).ToString
        MgmtEngrAPLiNo.Text = (engrEmpAPNos.Tables(0).Rows(0)("A&P_License_Number")).ToString

        'this fills the listbox that shows the qualifications that engineer holds
        For counter = 0 To engrQuals.Tables(0).Rows.Count - 1
            MgmtMngEngrsQualsListBox.Items.Add((engrQuals.Tables(0).Rows(counter)("Qualification_Name").ToString))
        Next

        MgmtMngEngrsAllQualsList.Items.Clear()
        'fills the available qualifications listbox with qualifications the engineer does not hold
        For counter = 0 To allQualsForListArray.Length - 1
            If Not compareQualsArray.Contains(allQualsForListArray(counter)) Then
                MgmtMngEngrsAllQualsList.Items.Add(allQualsForListArray(counter))
            End If
        Next

    End Sub

    Private Sub MgmtMngEngrEditEngrProfDrpDwnBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrEditEngrProfDrpDwnBut.Click
        're-enables the textboxes
        For Each cntrl As Control In MgmtMngEngrsTab.Controls
            If TypeOf cntrl Is System.Windows.Forms.TextBox Then
                cntrl.Enabled = True
            End If
        Next
        MgmtEngrCanBeTICCB.Enabled = True
        MgmtMngEngrsDoBMTB.Enabled = True
        MgmtMngEngrsHireDateMTB.Enabled = True
        MgmtMngEngrsAddEngrPanel.Show()
        MgmtMngEngrsDoneEditingBut.Show()

    End Sub

    Private Sub MgmtMngEngrsProfileImgLoctnTB_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsProfileImgLoctnTB.Click
        'allows the user to choose an image to give to insert into the 
        'selected engineer's profile

        'set the filters
        MgmtMngEngrsSelectEngrImageOFD.InitialDirectory = "C:\Users\zande_000\Documents\LEMSEngineerResources\LEMSEngineerPictures"
        MgmtMngEngrsSelectEngrImageOFD.Filter = "Image Files(*.JPG;*.PNG)|*.JPG;*.PNG|All files (*.*)|*.*"
        'show the dialog
        MgmtMngEngrsSelectEngrImageOFD.ShowDialog()
        'check the result of the dialog
        If MgmtMngEngrsSelectEngrImageOFD.ShowDialog = Windows.Forms.DialogResult.OK Then
            MgmtMngEngrsProfileImgLoctnTB.Text = MgmtMngEngrsSelectEngrImageOFD.FileName.ToString
        End If


    End Sub

    Private Sub SaveChangesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveChangesToolStripMenuItem.Click
        'calls the subroutine that writes the info back to the database
        MgmtEngrInfoWriteEditBack()
    End Sub

    Private Sub MgmtMngEngrAddEngrProfDrpDwnBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrAddEngrProfDrpDwnBut.Click
        For counter = 0 To allQualsForListArray.Length - 1
            MgmtMngEngrsAllQualsList.Items.Add(allQualsForListArray(counter))
        Next
        clearValuesAfterUpdate()
        'MgmtMngEngrsListBox.ClearSelected()
        MgmtEngrENIDLab.Text = "##"
        MgmtMngEngrsAddEngrPanel.Show()
        MgmtMngEngrsCompleteProfBut.Show()
        MgmtMngEngrsDoneEditingBut.Hide()
        'enable and clear the textboxes
        enableTextboxes()
        'validate them
        'create a new record- put the data in
        'update the engineers list box- in the done button routine
    End Sub

    Private Sub MgmtMngEngrsDoneEditingBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsDoneEditingBut.Click
        MgmtEngrInfoWriteEditBack()
        clearValuesAfterUpdate()
        readInfoFromDB()
    End Sub

    Private Sub clearValuesAfterUpdate()
        For Each cntrl As Control In MgmtMngEngrsTab.Controls
            If TypeOf cntrl Is System.Windows.Forms.TextBox Then
                cntrl.Text = Nothing
            End If
        Next
        MgmtEngrCanBeTICCB.CheckState = CheckState.Unchecked
        MgmtMngEngrsDoBMTB.Text = Nothing
        MgmtMngEngrsHireDateMTB.Text = Nothing
        MgmtMngEngrsQualsListBox.Items.Clear()
    End Sub

    Private Sub enableTextboxes()
        For Each cntrl As Control In MgmtMngEngrsTab.Controls
            If TypeOf cntrl Is System.Windows.Forms.TextBox Or TypeOf cntrl Is System.Windows.Forms.MaskedTextBox Then
                cntrl.Enabled = True
            End If
        Next

        MgmtEngrCanBeTICCB.Enabled = True

    End Sub

    Private Sub disenableTextboxes()
        For Each cntrl As Control In MgmtMngEngrsTab.Controls
            If TypeOf cntrl Is System.Windows.Forms.TextBox Or TypeOf cntrl Is System.Windows.Forms.MaskedTextBox Then
                cntrl.Enabled = False
            End If
        Next

        MgmtEngrCanBeTICCB.Enabled = False
    End Sub

    Private Sub MgmtEngrInfoWriteEditBack()
        'Use of parameters was adapted from http://www.mikesdotnetting.com/Article/26/parameter-queries-in-asp-net-with-ms-access

        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String
        Dim qualIDToWrite As Integer
        Dim qualScoreToWrite As Double

        'update the names and dates:
        'sql = "UPDATE [Engineers] SET [Engineers.Forename]= " & MgmtEngrForeNameTB.Text
        'sql = sql & ", [Engineers.Surname]= " & MgmtEngrSurnameTB.Text
        'sql = sql & ", [Engineers.Engineer_Image]= " & MgmtMngEngrsProfileImgLoctnTB.Text
        'sql = sql & ", [Engineers.DoB]= " & MgmtEngrDoBTB.Text
        'sql = sql & ", [Engineers.Hire_Date]= " & MgmtEngrHireDateTB.Text
        'sql = sql & ", [Engineers.Can_Be_TIC]= " & MgmtEngrCanBeTICCB.CheckState
        'sql = sql & ", WHERE [Engineers.Engineer_ID]= " & engineerID.ToString & ";"

        qualScoreToWrite = selectedEngineerQualScore()
        sql = "UPDATE [Engineers] SET [Engineers].[Forename]= ?, [Engineers].[Surname]= ?, [Engineers].[Engineer_Image]= ?, [Engineers].[DoB]= ?, [Engineers].[Hire_Date]= ?, [Engineers].[Can_Be_TIC]= ?, [Engineers].[Total_Qual_Score]= ? WHERE [Engineers].[Engineer_ID] = ?"

        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@Forename", MgmtEngrForeNameTB.Text.ToString)
        cmd.Parameters.AddWithValue("@Surname", MgmtEngrSurnameTB.Text.ToString)
        cmd.Parameters.AddWithValue("@Engineer_Image", MgmtMngEngrsProfileImgLoctnTB.Text.ToString)
        cmd.Parameters.AddWithValue("@DoB", MgmtMngEngrsDoBMTB.Text.ToString)
        cmd.Parameters.AddWithValue("@Hire_Date", MgmtMngEngrsHireDateMTB.Text.ToString)
        cmd.Parameters.AddWithValue("@Can_Be_TIC", MgmtEngrCanBeTICCB.CheckState)
        cmd.Parameters.AddWithValue("@Total_Qual_Score", qualScoreToWrite)
        cmd.Parameters.AddWithValue("@Engineer_ID", CInt(engineerID))
        connec.Open()
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        connec.Close()

        'need to match the engineer id with the qualification id
        'the data only goes into the linking table
        'compare the values in the list box to those in the array
        'if a value is in the listbox that is not in the array then write out that value to the database
        'need the id of the qualification

        sql = "INSERT INTO [EN/QUALS] VALUES (?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        For listCounter = 0 To MgmtMngEngrsQualsListBox.Items.Count - 1
            If Not compareQualsArray.Contains(MgmtMngEngrsQualsListBox.Items(listCounter)) Then
                'go to a function that searches through the dataset and returns the corresponding Qualification ID
                qualIDToWrite = returnQualID(MgmtMngEngrsQualsListBox.Items(listCounter))
                cmd.Parameters.AddWithValue("@EN_ID", engineerID)
                cmd.Parameters.AddWithValue("@QUAL_ID", qualIDToWrite)
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
            End If
        Next
        connec.Close()

        MgmtMngEngrsDoneEditingBut.Hide()
        readInfoFromDB()
        CreateRoster(CInt(Month(Now)))
    End Sub

    Private Function returnQualID(ByVal inQualName As String) As Integer
        Dim returnID As Integer

        For counter = 0 To dsQualsCompareWriteBack.Tables(0).Rows.Count - 1
            If inQualName = dsQualsCompareWriteBack.Tables(0).Rows(counter)("Qualification_Name") Then
                returnID = dsQualsCompareWriteBack.Tables(0).Rows(counter)("QUAL_ID")
            End If
        Next

        Return returnID

    End Function

    Private Sub MgmtMngEngrsAddEngrQualsDoneBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsAddEngrQualsDoneBut.Click
        'qualifications added/removed to/from list box
        'panel disappears
        'recalculate qual score and assign the value to the label NOT TO THE DATABASE
        'the value in the label will be recalculated as a check and then written to the db
        MgmtEngrQualScoreLab.Text = selectedEngineerQualScore().ToString
        MgmtMngEngrsAddEngrPanel.Hide()
    End Sub

    Private Sub MgmtMngEngrsCompleteProfBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsCompleteProfBut.Click
        'this is for adding an entirely new record.
        'also need to write out the qualification score
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String
        Dim newEnIDForQuals As Integer
        Dim qualIDToWrite As Integer

        sql = "INSERT INTO [Engineers] ([Forename], [Surname], [Engineer_Image], [DoB], [Hire_Date], [Can_Be_TIC]) VALUES (?,?,?,?,?,?)"

        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@Forename", MgmtEngrForeNameTB.Text)
        cmd.Parameters.AddWithValue("@Surname", MgmtEngrSurnameTB.Text)
        cmd.Parameters.AddWithValue("@Engineer_Image", MgmtMngEngrsProfileImgLoctnTB.Text)
        cmd.Parameters.AddWithValue("@DoB", MgmtMngEngrsDoBMTB.Text)
        cmd.Parameters.AddWithValue("@Hire_Date", MgmtMngEngrsHireDateMTB.Text)
        cmd.Parameters.AddWithValue("@Can_Be_TIC", MgmtEngrCanBeTICCB.CheckState)
        connec.Open()
        cmd.ExecuteNonQuery()
        connec.Close()

        'writes out the qualifications below
        newEnIDForQuals = returnNewEngineerID()
        sql = "INSERT INTO [EN/QUALS] ([EN_ID],[QUAL_ID]) VALUES (?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        For listCounter = 0 To MgmtMngEngrsQualsListBox.Items.Count - 1
            'go to a function that searches through the dataset and returns the corresponding Qualification ID
            'this function is called every time. Recursion could be a better method? Look into.
            qualIDToWrite = returnQualID(MgmtMngEngrsQualsListBox.Items(listCounter))
            cmd.Parameters.AddWithValue("@EN_ID", newEnIDForQuals)
            cmd.Parameters.AddWithValue("@QUAL_ID", qualIDToWrite)
            cmd.ExecuteNonQuery()
            cmd.Parameters.Clear()
        Next
        connec.Close()

        sql = "INSERT INTO [Employee_and_A&P] VALUES (?,?,?)"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@Engineer_ID", newEnIDForQuals)
        cmd.Parameters.AddWithValue("@Employee_Number", CInt(MgmtEngrEmpNoTB.Text))
        cmd.Parameters.AddWithValue("@A&P_License_Number", CInt(MgmtEngrAPLiNo.Text))
        connec.Open()
        cmd.ExecuteNonQuery()
        connec.Close()

        'hides the panel
        MgmtMngEngrsAddEngrPanel.Hide()
        'reorganises the listboxes and rosters in the program
        MgmtMngEngrsListBox.SelectedItem = MgmtMngEngrsListBox.Items()
        'reconfigure roster
        afterUpdates()

        'message to notify user of successful addition
        MessageBox.Show("Engineer successfully added to the database", "LEMS Message: Successful Database Addition", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Function returnNewEngineerID() As Integer
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dAdapMostRecntENID As OleDbDataAdapter
        Dim dsMostRecntENID As New DataSet
        Dim sql As String

        Dim tempENID As Integer

        'this gets the most recently added record from the database
        'as it is an autonumber, access assigns it
        'this is called after writing so that the qualifications can be written with an engineer ID reference
        sql = "SELECT MAX([Engineer_ID]) AS engID FROM [Engineers]"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdapMostRecntENID = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapMostRecntENID.Fill(dsMostRecntENID, "Engineers")
        connec.Close()

        tempENID = CInt(dsMostRecntENID.Tables(0).Rows(0)("engID"))
        Return tempENID
    End Function

    Private Sub MgmtMngEngrDelEngrProfDrpDwnBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrDelEngrProfDrpDwnBut.Click
        'deletes the selected engineer from the database and list
        'the item in the list must also be deleted, and the list index re-jigged
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        'message box to check if the user is sure they wish to delete the record:
        MsgBox("Are you sure you want to delete this record? This action cannot be undone.", MsgBoxStyle.YesNo, "LEMS Message: Deleting Record")
        If MsgBoxResult.Yes Then
            sql = "DELETE FROM [Engineers] WHERE [Engineers].[Engineer_ID] = ?"
            Dim cmd As New OleDbCommand(sql, connec)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("Engineer_ID", engineerID)
            connec.Open()
            cmd.ExecuteNonQuery()
            connec.Close()
            'by using referential integrity in the Access Database, all the related records should be deleted
        End If

        're-jig the list box
        MgmtMngEngrsListBox.Items.Remove(MgmtMngEngrsListBox.SelectedIndex)
        MgmtMngEngrsQualsListBox.Items.Clear()
        MgmtMngEngrsAllVactnDGV.Rows.Clear()
        MgmtMngEngrsTDYDatesDGV.Rows.Clear()
        For Each cntrl As Control In MgmtMngEngrsTab.Controls
            If TypeOf cntrl Is System.Windows.Forms.TextBox Then
                cntrl.Text = Nothing
            End If
        Next
        MgmtEngrCanBeTICCB.CheckState = CheckState.Unchecked
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(0).Value = 0
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(1).Value = 0
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(2).Value = 0
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(3).Value = 0
        MgmtMngEngrsOccurrencesDGV.Rows(0).Cells(4).Value = 0
        MgmtEngrQualScoreLab.Text = "##"
        MgmtEngrENIDLab.Text = "##"

        'reload the listbox and reconfigure roster

        afterUpdates()

        'inform the user that the record has been deleted
        MessageBox.Show("All instances of the selected engineer have been deleted", "LEMS Message: Record Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub afterUpdates()
        'ensure the datagridview is clear
        EngrRosterSpace.Rows.Clear()
        MgmtMngEngrsListBox.Items.Clear()
        surnameDataSelect()
        CreateRoster(CInt(Month(Now)))
        exportRosterToExcel(rosterFileName)
    End Sub

    Private Function selectedEngineerQualScore() As Integer
        'recalculate when edited
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim inDataSet As New DataSet
        Dim dAdapQualCompareData As New OleDbDataAdapter
        Dim sql As String
        Dim runQualTotal As Decimal

        sql = "SELECT * FROM [Qualifications]"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdapQualCompareData = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapQualCompareData.Fill(inDataSet, "Qualifications")
        connec.Close()

        'two for loops are used to search through and compare the qualifications in the listbox to those in the engineer's
        'loaded dataset. 
        For checkCount = 0 To MgmtMngEngrsQualsListBox.Items.Count - 1
            For dataSetCounter = 0 To inDataSet.Tables(0).Rows.Count - 1
                If MgmtMngEngrsQualsListBox.Items.Item(checkCount).ToString = inDataSet.Tables(0).Rows(dataSetCounter)("Qualification_Name").ToString Then
                    runQualTotal = runQualTotal + CDec(inDataSet.Tables(0).Rows(dataSetCounter)("Qualification_Rating"))
                    Exit For
                End If
            Next
        Next

        Return runQualTotal

    End Function

    'the two subroutines below add and remove qualifications to and from the engineer profile
    'there is a routine to add the qualifications the engineer DOES NOT HOLD to the qualifications available list box
    Private Sub MgmtMngEngrsAddQualToEngrBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsAddQualToEngrBut.Click

        MgmtMngEngrsQualsListBox.Items.Add(MgmtMngEngrsAllQualsList.SelectedItem.ToString)
        MgmtMngEngrsAllQualsList.Items.Remove(MgmtMngEngrsAllQualsList.SelectedItem)

    End Sub

    Private Sub MgmtMngEngrsRemQualFrmEngrBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsRemQualFrmEngrBut.Click

        MgmtMngEngrsAllQualsList.Items.Add(MgmtMngEngrsQualsListBox.SelectedItem.ToString)
        MgmtMngEngrsQualsListBox.Items.Remove(MgmtMngEngrsQualsListBox.SelectedItem)

    End Sub

    'Private Sub CDST4DailySheetDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles CDST4DailySheetDGV.CellClick
    '    If e.ColumnIndex = cdsT4Amts.Index Then
    '        'pull up panel of T4 engineers
    '    End If
    'End Sub

    'Private Sub CDST3DailySheetDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles CDST3DailySheetDGV.CellClick
    '    If e.ColumnIndex = cdsT3Amts.Index Then
    '        'pull up panel of T3 engineers
    '    End If
    'End Sub

    Private Sub currentDailySheetFirstSave()
        DailySheetOperations.mainControl()
    End Sub

    Private Sub MgmtCdsSaveOnlyBut_Click(sender As Object, e As EventArgs) Handles MgmtCdsSaveOnlyBut.Click
        MgmtCdsSave()
    End Sub

    Sub MgmtCdsSave()
        DailySheetOperations.saveDGV("mgmtLEMS")
    End Sub

    Private Sub MgmtCdsSaveSendBut_Click(sender As Object, e As EventArgs) Handles MgmtCdsSaveSendBut.Click
        'save the daily sheet
        MgmtCdsSave()
        'create the hard copy
        DailySheetOperations.createHardCopyDS()
        'email the hard copy
        DailySheetOperations.sendDS()
    End Sub

    Private Sub MgmtResultList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MgmtResultList.SelectedIndexChanged
        'code taken and adapted for use from:
        'https://social.msdn.microsoft.com/Forums/vstudio/en-US/c165c82d-ac79-477e-abab-3efd330d149f/how-to-open-pdf-file-in-vbnet-applicatin

        MgmtDSAAxAcroPDF1.src = MgmtResultList.SelectedItem

    End Sub

    Private Sub loadShiftColours()
        'used to load the shift colours for the shift rosters

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift1.txt", OpenMode.Input)
        'assign the colour to the variable
        shift1Colour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettShift1ColorPB.BackColor = ColorTranslator.FromHtml(shift1Colour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift2.txt", OpenMode.Input)
        'assign the colour to the variable
        shift2Colour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettShift2ColorPB.BackColor = ColorTranslator.FromHtml(shift2Colour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift3.txt", OpenMode.Input)
        'assign the colour to the variable
        shift3Colour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettShift3ColorPB.BackColor = ColorTranslator.FromHtml(shift3Colour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Rest.txt", OpenMode.Input)
        'assign the colour to the variable
        restColour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettRestDayColorPB.BackColor = ColorTranslator.FromHtml(restColour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Vacation.txt", OpenMode.Input)
        'assign the colour to the variable
        vacationColour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettVacColorPB.BackColor = ColorTranslator.FromHtml(vacationColour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Sick.txt", OpenMode.Input)
        'assign the colour to the variable
        sickColour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettShiftSickColorPB.BackColor = ColorTranslator.FromHtml(sickColour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Training.txt", OpenMode.Input)
        'assign the colour to the variable
        trainingColour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettTrngColorPB.BackColor = ColorTranslator.FromHtml(trainingColour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TDY.txt", OpenMode.Input)
        'assign the colour to the variable
        tdyColour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettShiftTDYColorPB.BackColor = ColorTranslator.FromHtml(tdyColour)

        'open the files
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TIC.txt", OpenMode.Input)
        'assign the colour to the variable
        ticColour = LineInput(1)
        'close the file
        FileClose(1)
        'assign the colour to the corresponding picturebox
        MgmtSettShiftTICColorPB.BackColor = ColorTranslator.FromHtml(ticColour)


    End Sub

    Private Sub MgmtSettShift1ColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettShift1ColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettShift1ColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift1.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If

    End Sub

    Private Sub MgmtSettShift2ColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettShift2ColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettShift2ColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift2.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettShift3ColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettShift3ColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettShift3ColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Shift3.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettRestDayColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettRestDayColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettRestDayColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Rest.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettVacColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettVacColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettVacColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Vacation.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettShiftSickColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettShiftSickColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettShiftSickColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Sick.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettTrngColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettTrngColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettTrngColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the correspondinf file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\Training.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettShiftTDYColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettShiftTDYColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettShiftTDYColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TDY.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettShiftTICColorPB_Click(sender As Object, e As EventArgs) Handles MgmtSettShiftTICColorPB.Click
        'set the color of the picture box to the selected colour
        If MgmtShiftColorDialog.ShowDialog() = DialogResult.OK Then
            MgmtSettShiftTICColorPB.BackColor = MgmtShiftColorDialog.Color
            'save the choice to the corresponding file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TIC.txt", OpenMode.Output)
            PrintLine(1, MgmtShiftColorDialog.Color.ToArgb)
            FileClose(1)
        End If
    End Sub

    Private Sub loadDSRecipients()
        'load the daily sheet recipients
        Dim localDSRecipients(50) As String
        Dim localCounter As Integer

        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\DSrecipients.txt", OpenMode.Input)
        Do
            localDSRecipients(localCounter) = LineInput(1)
            localCounter = localCounter + 1
        Loop Until EOF(1)
        FileClose(1)

        For counter = 0 To localCounter
            MgmtSettDsRecipientsLB.Items.Add(localDSRecipients(counter))
        Next

    End Sub

    Private Sub AddNewDailySheetRecipientToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddNewDailySheetRecipientToolStripMenuItem.Click
        'shows the add new recipient panel
        MgmtSettAddNewDsRecipientPanel.Show()
    End Sub

    Private Sub EditDailySheetRecipientToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditDailySheetRecipientToolStripMenuItem.Click
        'shows the edit recipient panel
        MgmtSettEditDsRecpntPanel.Show()
    End Sub

    Private Sub DeleteDailySheetRecipientToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteDailySheetRecipientToolStripMenuItem.Click
        'deletes a daily sheet recipient
        'check a qualification has been selected
        If MgmtSettDsRecipientsLB.SelectedIndex = -1 Then
            'message to user to notify them nothing has been selected
            MessageBox.Show("A daily sheet recipient has not been selected to delete. Select a recipient from the listbox", "LEMS Warning: Nothing Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'clear the item from the listbox
            MgmtSettDsRecipientsLB.Items.Remove(MgmtSettDsRecipientsLB.SelectedItem)
            'resave the file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\DSrecipients.txt", OpenMode.Output)
            For counter = 0 To MgmtSettDsRecipientsLB.Items.Count - 1
                PrintLine(1, MgmtSettDsRecipientsLB.Items(counter))
            Next
            FileClose(1)
        End If
    End Sub

    Private Sub MgmtSettAddNewDSRecpntBut_Click(sender As Object, e As EventArgs) Handles MgmtSettAddNewDSRecpntBut.Click
        'appends a new recipient to the text file
        Dim localNewRecipient As String


        'get the value of the textbox
        localNewRecipient = MgmtSettDsNewRecpntTB.Text

        If IsNothing(MgmtSettEditDsRecpntTB.Text) Then
            'messagebox to user to alert them nothing has been entered
            MessageBox.Show("A recipient has not been selected for editing. Select a recipient from the listbox", "LEMS Warning: Nothing Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf Not MgmtSettEditDsRecpntTB.Text.Contains("@") Then
            'messagebox to user to alert them the entered string is not a valid email address
            MessageBox.Show("The email address entered is not valid. A valid email address contains an '@' and a domain, for example, lems@lems.com. Enter a valid email address", "LEMS Warning: Invalid Email Address", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'save it to the file, by overwriting entire file
            'the file is only a few kb so this is the most efficient way of
            'doing it rather than finding the name then updating it
            localNewRecipient = MgmtSettEditDsRecpntBut.Text
            MgmtSettDsRecipientsLB.Items.Add(localNewRecipient)
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\DSrecipients.txt", OpenMode.Output)
            For counter = 0 To MgmtSettDsRecipientsLB.Items.Count - 1
                PrintLine(1, MgmtSettDsRecipientsLB.Items(counter))
            Next
            FileClose(1)
        End If

        'clear the textbox
        MgmtSettDsNewRecpntTB.Clear()

        'refill the listbox
        loadDSRecipients()

        'close the panel
        MgmtSettAddNewDsRecipientPanel.Hide()

    End Sub

    Private Sub MgmtSettEditDsRecpntBut_Click(sender As Object, e As EventArgs) Handles MgmtSettEditDsRecpntBut.Click
        'alters the recipient's email address
        Dim localEditRecipient As String

        'checks that the textbox contains a string and a @ in an effort to check it is a valid email address
        If IsNothing(MgmtSettEditDsRecpntTB.Text) Then
            'messagebox to user to alert them nothing has been entered
            MessageBox.Show("A recipient has not been selected for editing. Select a recipient from the listbox", "LEMS Warning: Nothing Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf Not MgmtSettEditDsRecpntTB.Text.Contains("@") Then
            'messagebox to user to alert them the entered string is not a valid email address
            MessageBox.Show("The email address entered is not valid. A valid email address contains an '@' and a domain, for example, lems@lems.com. Enter a valid email address", "LEMS Warning: Invalid Email Address", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'save it to the file, by overwriting entire file
            'the file is only a few kb so this is the most efficient way of
            'doing it rather than finding the name then updating it
            localEditRecipient = MgmtSettEditDsRecpntBut.Text
            MgmtSettDsRecipientsLB.SelectedItem = localEditRecipient
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\DSrecipients.txt", OpenMode.Output)
            For counter = 0 To MgmtSettDsRecipientsLB.Items.Count - 1
                PrintLine(1, MgmtSettDsRecipientsLB.Items(counter))
            Next
            FileClose(1)
        End If

        'reload the listbox
        loadDSRecipients()

        'close the panel
        MgmtSettEditDsRecpntPanel.Hide()

    End Sub

    Private Sub AddQualificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddQualificationToolStripMenuItem.Click
        'show the panel
        MgmtSettAddNewQualPanel.Show()

    End Sub

    Private Sub MgmtSettAddNewQualBut_Click(sender As Object, e As EventArgs) Handles MgmtSettAddNewQualBut.Click
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        Dim localQualName As String
        Dim localQualScore As Integer

        'presence check the textboxes
        If MgmtSettAddNewQualNameTB.Text = Nothing Or MgmtSettAddNewQualScoreTB.Text = Nothing Then
            MessageBox.Show("The qualification fields are empty. Enter the values", "LEMS Warning: Empty Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        'get the values
        localQualName = MgmtSettAddNewQualNameTB.Text
        'get and validate the score
        If Not IsNumeric(MgmtSettAddNewQualScoreTB.Text) Then
            MessageBox.Show("The qualification score is not a numerical value. Enter a number", "LEMS Warning: Non-Numeric Rating", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'assign the value if it is numeric
            localQualScore = MgmtSettAddNewQualScoreTB.Text
        End If

        'write the scores out the database
        sql = "INSERT INTO [Qualifications] ([Qualifications].[Qualification_Name], [Qualifications].[Qualification_Rating]) VALUES (?,?) "
        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@Qualification_Name", localQualName)
        cmd.Parameters.AddWithValue("@Qualification_Rating", localQualScore)
        cmd.ExecuteNonQuery()
        connec.Close()

        'clear the texboxes
        MgmtSettAddNewQualNameTB.Clear()
        MgmtSettAddNewQualScoreTB.Clear()

        'hide the panel
        MgmtSettAddNewQualPanel.Hide()

        'message to notify user of successful addition
        MessageBox.Show("Qualification successfully added to the database", "LEMS Message: Successful Database Addition", MessageBoxButtons.OK, MessageBoxIcon.Information)

        'reload the qualifications listbox
        qualDataSelect()

    End Sub

    Private Sub EditQualificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditQualificationToolStripMenuItem.Click
        MgmtSettEditSelectedQualPanel.Show()

    End Sub

    Private Sub MgmtSettEditSelectedQualBut_Click(sender As Object, e As EventArgs) Handles MgmtSettEditSelectedQualBut.Click
        'update the selected qualification
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        Dim localQualNameToUpdate As String
        Dim localQualRatingToUpdate As Decimal

        'check to see if the user has selected an item
        If MgmtSettAllQualsLB.SelectedIndex = -1 Then
            'message to user to notify them nothing has been selected
            MessageBox.Show("A Qualification has not been selected for editing. Select a qualification from the listbox", "LEMS Warning: Nothing Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else

            If MgmtSettEditQualNameTB.Text = Nothing Or MgmtSettEditQualRatingTB.Text = Nothing Then
                'message to user to alert them that the fields are empty
                MessageBox.Show("The qualification fields are empty. Enter the values", "LEMS Warning: Empty Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ElseIf Not MgmtSettEditQualRatingTB.Text Then
                'message to alert user that the value is not numeric for the qualification rating
                MessageBox.Show("The qualification score is not a numerical value. Enter a number", "LEMS Warning: Non-Numeric Rating", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                'set the textboxes equal to the selected item of the listbox
                MgmtSettEditQualNameTB.Text = qualsForSettings(MgmtSettAllQualsLB.SelectedIndex).qualName
                MgmtSettEditQualRatingTB.Text = qualsForSettings(MgmtSettAllQualsLB.SelectedIndex).qualRating

                localQualNameToUpdate = MgmtSettEditQualNameTB.Text
                localQualRatingToUpdate = MgmtSettEditQualRatingTB.Text

                'update the qualifications
                sql = "UPDATE [Qualifications] SET [Qualifications].[Qualification_Name]= ?, [Qualifications].[Qualification_Rating] WHERE [Qualifications].[QUAL_ID] = ?"

                Dim cmd As New OleDbCommand(sql, connec)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("@Qualification_Name", localQualNameToUpdate)
                cmd.Parameters.AddWithValue("@Qualification_Rating", localQualRatingToUpdate)
                cmd.Parameters.AddWithValue("@QUAL_ID", qualsForSettings(MgmtSettAllQualsLB.SelectedIndex).qualRating)
                connec.Open()
            End If
        End If
        connec.Close()

        'hide the panel
        MgmtSettEditSelectedQualPanel.Hide()

        'reload the qualifications listboxes
        qualDataSelect()

        'message to notify user of successful amendment
        MessageBox.Show("Qualification successfully edited in the database", "LEMS Message: Successful Database Amendment", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub DeleteQualificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteQualificationToolStripMenuItem.Click
        'delete the selected qualification
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        Dim localQualIDToDelete As Integer
        Dim localDialogResult As Integer

        'check a qualification has been selected
        If MgmtSettAllQualsLB.SelectedIndex = -1 Then
            'message to user to notify them nothing has been selected
            MessageBox.Show("A Qualification has not been selected to delete. Select a qualification from the listbox", "LEMS Warning: Nothing Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        'get the ID of the qualification to delete
        localQualIDToDelete = qualsForSettings(MgmtSettAllQualsLB.SelectedIndex).qualID

        localDialogResult = MessageBox.Show("Are you sure you want to delete this qualification?", "LEMS Message: Deleting Record", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If localDialogResult = DialogResult.Yes Then
            sql = "DELETE FROM [Qualifications] WHERE [Qualifications].[QUAL_ID] = ?"
            Dim cmd As New OleDbCommand(sql, connec)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@QUAL_ID", localQualIDToDelete)
            connec.Open()
            cmd.ExecuteNonQuery()
            connec.Close()
        End If

        'reload the qualifications listboxes
        qualDataSelect()

        'alert the user the qualification has been successfully deleted
        MessageBox.Show("Qualification successfully deleted in the database", "LEMS Message: Successful Database Deletion", MessageBoxButtons.OK, MessageBoxIcon.Information)


    End Sub

    Private Sub loadNoTerminals()
        'set the selected text to the item in the file
        'needs to be included in the onload routine
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TIC.txt", OpenMode.Input)
        MgmtSettNoTerminalsCB.SelectedText = LineInput(1)
        FileClose(1)
    End Sub

    Private Sub MgmtSettNoTerminalsCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MgmtSettNoTerminalsCB.SelectedIndexChanged
        'write the value of the combobox to a text file
        'save the choice to the corresponding file
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\TIC.txt", OpenMode.Output)
        PrintLine(1, MgmtSettNoTerminalsCB.SelectedText)
        FileClose(1)
    End Sub

    Private Sub fillRosterWithApprovedRequests()
        'get the approved requests from the database
        'match the requests to the engineers in the roster
        'fill the requests for each engineer, one at a time
        'after each 'fill' check that there are enough people
        'if not, skip and deny the request

        'database variables
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dsApprvdRqsts As New DataSet
        Dim dAdapApprvdRqsts As New OleDbDataAdapter
        Dim sql As String
        'request and roster-filling variables
        Dim localIDsArray() As Integer
        Dim localRostPostn As Integer
        Dim requestDate As Date
        'filename to save the roster
        Dim localFilename As String

        'get the data
        sql = "SELECT * FROM [Requests] INNER JOIN [EN/REQS] ON [Requests].[REQ_ID] = [EN/REQS].[REQ_ID] WHERE [Requests].[Request_Approved] = TRUE"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdapApprvdRqsts = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapApprvdRqsts.Fill(dsApprvdRqsts, "Requests")
        connec.Close()

        'get the engineer IDs related to the requests
        localIDsArray = (From myRow In dsApprvdRqsts.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("EN_ID")).ToArray

        'need to compare the current month with the month of the request
        'need to get the day and month of the request
        'use DateTime.Parse

        If dsApprvdRqsts.Tables(0).Rows.Count > 0 Then
            'loop through the table to add the requests to the dgv
            For counter = 0 To dsApprvdRqsts.Tables(0).Rows.Count - 1
                'check if the month of the request if equal to the current month
                requestDate = DateTime.Parse(dsApprvdRqsts.Tables(0).Rows(counter)("Request_Date"))
                If totalsArray(requestDate.Day - 1) > minNoEngineers Then
                    'function to return the roster position of the engineer with the request
                    localRostPostn = returnEngrRostPostn(localIDsArray(counter))
                    If requestDate.Month = Now.Month And requestDate.Year = Now.Year Then
                        'add the request
                        Select Case dsApprvdRqsts.Tables(0).Rows(counter)("Request_Type")
                            Case Is = "SC"
                                EngrRosterSpace.Rows(localRostPostn).Cells(requestDate.Day).Value = dsApprvdRqsts.Tables(0).Rows(counter)("Request_Shift_Change_To")
                                localRostPostn = returnEngrRostPostn(dsApprvdRqsts.Tables(0).Rows(counter)("Request_Shift_Change_With_Engineer"))
                                EngrRosterSpace.Rows(localRostPostn).Cells(requestDate.Day).Value = dsApprvdRqsts.Tables(0).Rows(counter)("Request_Shift_Change_From")
                            Case Else
                                EngrRosterSpace.Rows(localRostPostn).Cells(requestDate.Day).Value = dsApprvdRqsts.Tables(0).Rows(counter)("Request_Type")
                        End Select
                    End If
                ElseIf totalsArray(requestDate.Day - 1) <= minNoEngineers Then
                    'delete the request from the database if the number of engineers is less than 6
                    sql = "DELETE FROM [Requests] WHERE [Requests].[REQ_ID] = ?"
                    cmd = New OleDbCommand(sql, connec)
                    cmd.Parameters.AddWithValue("@REQ_ID", dsApprvdRqsts.Tables(0).Rows(counter)("Requests.REQ_ID"))
                    connec.Open()
                    cmd.ExecuteNonQuery()
                    connec.Close()
                End If
            Next
            localFilename = rosterFileName
            exportRosterToExcel(localFilename)
        End If

        'import the updated roster
        localFilename = rosterFileName
        importFromExcelFile(localFilename)

    End Sub

    Private Function returnEngrRostPostn(ByVal inEngrID As Integer) As Integer
        'linear search to find the engineer ID and the corresponding
        'position in the roster
        Dim posToRtrn As Integer

        'loop through the structure
        For counter = 0 To engrIDRosterPostn.Length - 1
            If engrIDRosterPostn(counter).engrID = inEngrID Then
                posToRtrn = engrIDRosterPostn(counter).engrRostPostn
            End If
        Next

        'return the id
        Return posToRtrn
    End Function

    Private Sub engrShiftTotals()
        'calculate the number of shifts each engineer is scheduled
        'to work during the month

        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String
        Dim eeCount As Integer
        Dim enCount As Integer
        Dim lnCount As Integer
        Dim ticCount As Integer

        'go along the shift roster and total the shifts
        'for each engineer
        'then write each variable to database in parameter query
        sql = "UPDATE [Engineers] SET [Engineers].[TIC_Occurrence]= ?, [Engineers].[EE_Shift_Occurrence]= ?, [Engineers].[EN_Shift_Occurrence]= ?, [Engineers].[LN_Shift_Occurrence] = ? WHERE [Engineers].[Engineer_ID] = ?"
        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        'go down the rows
        For outerCounter = 0 To EngrRosterSpace.Rows.Count - 1
            'reinitialise the variables
            eeCount = 0
            enCount = 0
            lnCount = 0
            ticCount = 0
            'go along the columns for each engineer
            For innerCounter = 1 To System.DateTime.DaysInMonth(Now.Year, Now.Month) - 1
                Select Case EngrRosterSpace.Rows(outerCounter).Cells(innerCounter).Value
                    Case Is = "EE"
                        eeCount = eeCount + 1
                    Case Is = "EN"
                        enCount = enCount + 1
                    Case Is = "LN"
                        lnCount = lnCount + 1
                    Case Is = "TIC"
                        ticCount = ticCount + 1
                End Select
            Next
            'write out the values to the database
            cmd.Parameters.AddWithValue("@TIC_Occurrence", ticCount)
            cmd.Parameters.AddWithValue("@EE_Occurrence", eeCount)
            cmd.Parameters.AddWithValue("@EN_Occurrence", enCount)
            cmd.Parameters.AddWithValue("@LN_Occurrence", lnCount)
            cmd.Parameters.AddWithValue("@Engineer_ID", engrIDRosterPostn(outerCounter).engrID)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            'clear the parameters for the next query
            cmd.Parameters.Clear()
        Next

        connec.Close()

    End Sub

    Private Sub loadVacation()
        'load the engineer vacation here that hasn't been filled
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim cmd As New OleDbCommand
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dsNonAddedVactn As New DataSet
        Dim dAdapNonAddedVactn As New OleDbDataAdapter
        Dim sql As String

        'sql string, orders it by preference- 1 being the highest preference and 3 being the lowest
        'so the dataset will contain smallest to largest
        sql = "SELECT * FROM [Vacation] INNER JOIN [EN/VA] ON [Vacation].[Vacation_ID] = [EN/VA].[VACTN_ID] WHERE [Vacation].[Vacation_Added] = FALSE"
        cmd = New OleDbCommand(sql, connec)
        dAdapNonAddedVactn = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapNonAddedVactn.Fill(dsNonAddedVactn, "Vacation")
        connec.Close()

        'call the sorting routine if the dataset is not empty
        If dsNonAddedVactn.Tables(0).Rows.Count > 0 Then
            sortVacation(dsNonAddedVactn)
        End If

    End Sub

    Private Sub sortVacation(ByRef inDataSet As DataSet)
        'sort the engineer vacation here
        'arrays to create from the dataset and then fill the roster
        Dim localEngrID() As Integer
        Dim localVacID() As Integer
        Dim localVacStartDateString() As String
        Dim localVacEndDateString() As String

        Dim localVacPriority() As Integer
        Dim localVacPref() As Integer
        Dim temp As tEngrVacation

        'fill the arrays from the dataset
        localEngrID = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("EN_ID")).ToArray
        localVacID = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Vacation_ID")).ToArray
        localVacStartDateString = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of String)("V_Start_Date")).ToArray
        localVacEndDateString = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of String)("V_End_Date")).ToArray
        localVacPriority = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Vacation_Priority")).ToArray
        localVacPref = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Vacation_Preference")).ToArray

        'declare the structure used to store the engineer and vacation details
        Dim localEngrVactn(localEngrID.Length - 1) As tEngrVacation

        'fill the structure
        For counter = 0 To localEngrID.Length - 1
            localEngrVactn(counter).engrID = localEngrID(counter)
            localEngrVactn(counter).engrRostPosn = returnEngrRostPostn(localEngrVactn(counter).engrID)
            localEngrVactn(counter).engrVactnID = localVacID(counter)
            'directly parse the string dates to System.DateTime dates so that the functions
            'associated with DateTime can be used (so the number of days can be caluclated etc)
            localEngrVactn(counter).engrVactnStartDate = DateTime.Parse(localVacStartDateString(counter))
            localEngrVactn(counter).engrVactnEndDate = DateTime.Parse(localVacEndDateString(counter))
            localEngrVactn(counter).engrVactnPrty = localVacPriority(counter)
            localEngrVactn(counter).engrVactnPref = localVacPref(counter)
        Next

        'sort the structure with a modified bubble sort
        'the highest order of vacation is where the engineer has the most
        'seniority and the lowest vacation preference (i.e 1 = first choice vacation,
        'so lower = more weighting)
        For outCount = 0 To localEngrVactn.Length - 1
            For inCount = 0 To localEngrVactn.Length - 1
                If inCount + 1 > localEngrVactn.Length - 1 Then
                    Exit For
                End If
                If localEngrVactn(inCount).engrVactnPrty = localEngrVactn(inCount + 1).engrVactnPrty And localEngrVactn(inCount).engrVactnPref = localEngrVactn(inCount + 1).engrVactnPref And
                    localEngrVactn(inCount).engrVactnStartDate = localEngrVactn(inCount + 1).engrVactnStartDate And localEngrVactn(inCount).engrVactnEndDate = localEngrVactn(inCount + 1).engrVactnEndDate Then
                    localEngrVactn(inCount).engrVactnStartDate = Nothing
                    localEngrVactn(inCount + 1).engrVactnStartDate = Nothing
                End If
                If localEngrVactn(inCount).engrVactnPrty < localEngrVactn(inCount + 1).engrVactnPrty And localEngrVactn(inCount).engrVactnPref > localEngrVactn(inCount + 1).engrVactnPref Then
                    temp = localEngrVactn(inCount)
                    localEngrVactn(inCount) = localEngrVactn(inCount + 1)
                    localEngrVactn(inCount + 1) = temp
                End If
            Next
        Next

        'fill the roster with the vacation
        fillRosterWithVacation(localEngrVactn)

    End Sub

    Private Sub fillRosterWithVacation(ByRef inVactnStruct() As tEngrVacation)
        'fill the roster with the vacation
        'with the vacation for specific engineers
        'also need to call roster colours and check there are enough engineers
        'then need to write back to the database to specify the number of vacation
        'days remaining

        'if the vacation overlaps to the next month, check if the roster exists
        'if not then, create the roster and fill the number of vacation days
        Dim checkWorkDays As Boolean

        Dim inStartDate As Date
        Dim inEndDate As Date

        'total days of vacation
        Dim totalDays As Integer
        Dim manipulateDays As Integer

        'variables to hold the differences between the months and years
        Dim monthDifference As Integer
        Dim monthAdvance As Integer
        Dim remainingMonths As Integer

        'create the array to analyse 
        Dim workingDaysArray(totalDays) As Boolean

        'control variables
        Dim arrayCount As Integer
        Dim localFileName As String
        Dim checkingCount As Integer

        For overCounter = 0 To inVactnStruct.Length - 1
            If inVactnStruct(overCounter).engrVactnStartDate <> Nothing Then
                checkWorkDays = checkWorkDaysForVacation(inVactnStruct(overCounter).engrVactnStartDate, inVactnStruct(overCounter).engrVactnEndDate)
            End If
            If checkWorkDays = True And inVactnStruct(overCounter).engrVactnStartDate <> Nothing Then
                inStartDate = inVactnStruct(overCounter).engrVactnStartDate
                inEndDate = inVactnStruct(overCounter).engrVactnEndDate

                totalDays = (inEndDate - inStartDate).Days
                manipulateDays = totalDays
                'gets the number of months between the two dates
                monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))

                'check the roster exists for the month regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
                checkRosterFileExists(localFileName, CInt(inStartDate.Month))

                'analyse the boolean array
                'different routines depending if the vacation is in the same month
                arrayCount = 0
                If inStartDate.Month = inEndDate.Month Then
                    For counter = inStartDate.Day To inEndDate.Day
                        EngrRosterSpace.Rows(inVactnStruct(overCounter).engrRostPosn).Cells(counter).Value = "V"
                        arrayCount = arrayCount + 1
                    Next
                    exportRosterToExcel(localFileName)
                End If

                'separate if statements to simplify code layout
                'checks the workdays between non-contiguous months in the same year
                If inStartDate.Month <> inEndDate.Month And inStartDate.Year = inEndDate.Year Then
                    monthDifference = inEndDate.Month - inStartDate.Month
                    'find if the minimum shift numbers have been met for the first month
                    For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                        EngrRosterSpace.Rows(inVactnStruct(overCounter).engrRostPosn).Cells(innerCounter).Value = "V"
                        arrayCount = arrayCount + 1
                        manipulateDays = manipulateDays - 1
                    Next
                    localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
                    exportRosterToExcel(localFileName)
                    For counter = 1 To monthDifference
                        'check the roster exists for the months regarding the vacation
                        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month + counter, True) & inStartDate.Year & ".xls"
                        checkRosterFileExists(localFileName, CInt(inStartDate.Month))
                        'loop for the next months
                        checkingCount = 0
                        For innerCounter = 0 To manipulateDays
                            checkingCount = checkingCount + 1
                            If checkingCount > System.DateTime.DaysInMonth(inStartDate.Year, CInt(inStartDate.Month + counter)) Then
                                checkingCount = 1
                                Exit For
                            End If
                            EngrRosterSpace.Rows(inVactnStruct(overCounter).engrRostPosn).Cells(checkingCount).Value = "V"
                            arrayCount = arrayCount + 1
                            manipulateDays = manipulateDays - 1
                            If innerCounter > totalDays Then
                                'exit this for loop to move to the next month
                                Exit For
                            End If
                        Next
                        exportRosterToExcel(localFileName)
                    Next
                End If

                'checks the workdays between non-contigous years
                If inStartDate.Year <> inEndDate.Year Then
                    'get the differences between the vacation
                    monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))
                    'check the first month
                    For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                        EngrRosterSpace.Rows(inVactnStruct(overCounter).engrRostPosn).Cells(innerCounter).Value = "V"
                        arrayCount = arrayCount + 1
                        'manipulateDays = manipulateDays - 1
                    Next
                    localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
                    exportRosterToExcel(localFileName)
                    'continue the loop until the month advance is greater than 12
                    remainingMonths = monthDifference
                    For monthCounter = 1 To monthDifference
                        monthAdvance = inStartDate.Month + monthCounter
                        'resets the months when the count is greater than 12
                        'this signifies that the new year analysis has been started
                        If monthAdvance > 12 Then
                            Exit For
                        End If
                        'check the roster exists for the months regarding the vacation
                        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & inStartDate.Year & ".xls"
                        checkRosterFileExists(localFileName, monthAdvance)
                        'loop for the next months
                        For innerCounter = 0 To manipulateDays - 1
                            EngrRosterSpace.Rows(inVactnStruct(overCounter).engrRostPosn).Cells(innerCounter).Value = "V"
                            arrayCount = arrayCount + 1
                            'manipulateDays = manipulateDays - 1
                            If innerCounter > totalDays Then
                                'exit this for loop to move to the next month
                                Exit For
                            End If
                            remainingMonths = remainingMonths - 1
                        Next

                        exportRosterToExcel(localFileName)
                    Next

                    'go to the next year
                    For monthCounter = 1 To remainingMonths
                        monthAdvance = inStartDate.Month + monthCounter
                        If monthAdvance > 12 Then
                            monthAdvance = monthCounter
                        End If
                        'check the roster exists for the months regarding the vacation
                        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & (inEndDate.Year) & ".xls"
                        checkRosterFileExists(localFileName, monthAdvance)
                        'loop for the next months
                        For innerCounter = 0 To manipulateDays - 1
                            EngrRosterSpace.Rows(inVactnStruct(overCounter).engrRostPosn).Cells(innerCounter + 1).Value = "V"
                            arrayCount = arrayCount + 1
                            'manipulateDays = manipulateDays - 1
                            If innerCounter > totalDays Then
                                'exit this for loop to move to the next month
                                Exit For
                            End If
                        Next
                        exportRosterToExcel(localFileName)
                    Next

                End If

                'write out to the database that the vacation has been added
                vacationAdded(inVactnStruct(overCounter).engrVactnID)
                'subtract the number of days of vacation
                vacationSubtractDays(inVactnStruct(overCounter).engrID, totalDays)
            End If

        Next
        importFromExcelFile(rosterFileName)
    End Sub

    Private Function checkWorkDaysForVacation(ByVal inStartDate As Date, ByVal inEndDate As Date) As Boolean
        'return a boolean depending on if the vacation can be added to the roster
        'this is dependant on the number of engineers working on the days the 
        'vaction is requested
        'there must be a 3/4 success rate for the vacation to be added

        'total days of vacation
        Dim totalDays As Integer
        Dim manipulateDays As Integer

        'variables to hold the differences between the months and years
        Dim monthDifference As Integer
        Dim monthAdvance As Integer
        Dim remainingMonths As Integer

        'number of vacation days
        totalDays = (inEndDate - inStartDate).Days
        'create the array to analyse with the length as the number of vacation days
        Dim workingDaysArray(totalDays) As Boolean

        'variables to hold the running totals and the average
        Dim workingPeriodTotals As Integer
        Dim avgWorkingTotals As Decimal

        'control variables
        Dim arrayCount As Integer
        Dim localFileName As String
        Dim checkingCount As Integer

        'assigments for the variables above
        manipulateDays = totalDays
        'gets the number of months between the two dates
        monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))

        'check the roster exists for the month regarding the vacation
        localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
        checkRosterFileExists(localFileName, CInt(inStartDate.Month))
        countWorkDays()
        'analyse the boolean array
        'different routines depending if the vacation is in the same month
        arrayCount = 0
        If inStartDate.Month = inEndDate.Month Then
            For counter = inStartDate.Day To inEndDate.Day - 1
                If arrayCount > (inEndDate - inStartDate).Days Then
                    Exit For
                End If
                If totalsArray(counter) > minNoEngineers Then
                    workingDaysArray(arrayCount) = True
                Else
                    workingDaysArray(arrayCount) = False
                End If
                arrayCount = arrayCount + 1
            Next
        End If

        'separate if statements to simplify code layout
        'checks the workdays between non-contiguous months in the same year
        If inStartDate.Month <> inEndDate.Month And inStartDate.Year = inEndDate.Year Then
            monthDifference = inEndDate.Month - inStartDate.Month
            'find if the minimum shift numbers have been met for the first month
            For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                If totalsArray(innerCounter - 1) > minNoEngineers Then
                    workingDaysArray(arrayCount) = True
                Else
                    workingDaysArray(arrayCount) = False
                End If
                arrayCount = arrayCount + 1
                manipulateDays = manipulateDays - 1
            Next
            For counter = 1 To monthDifference
                'check the roster exists for the months regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month + counter, True) & inStartDate.Year & ".xls"
                checkRosterFileExists(localFileName, CInt(inStartDate.Month + counter))
                countWorkDays()
                'loop for the next months
                For innerCounter = 1 To manipulateDays - 1
                    checkingCount = innerCounter
                    If checkingCount > System.DateTime.DaysInMonth(inStartDate.Year, CInt(inStartDate.Month + counter)) Then
                        checkingCount = 1
                    End If
                    If totalsArray(checkingCount - 1) > minNoEngineers Then
                        workingDaysArray(arrayCount) = True
                    Else
                        workingDaysArray(arrayCount) = False
                    End If
                    arrayCount = arrayCount + 1
                    manipulateDays = manipulateDays - 1
                    If innerCounter > workingDaysArray.Length - 1 Then
                        'exit this for loop to move to the next month
                        Exit For
                    End If
                Next
            Next
        End If


        'checks the workdays between non-contigous years
        If inStartDate.Year <> inEndDate.Year Then
            'get the differences between the vacation
            monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))
            'check the first month
            For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                If totalsArray(innerCounter - 1) > minNoEngineers Then
                    workingDaysArray(arrayCount) = True
                Else
                    workingDaysArray(arrayCount) = False
                End If
                arrayCount = arrayCount + 1
                manipulateDays = manipulateDays - 1
            Next

            'continue the loop until the month advance is greater than 12
            remainingMonths = monthDifference
            For monthCounter = 1 To monthDifference
                monthAdvance = inStartDate.Month + monthCounter
                'resets the months when the count is greater than 12
                'this signifies that the new year analysis has been started
                If monthAdvance > 12 Then
                    Exit For
                End If
                'check the roster exists for the months regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & (inStartDate.Year) & ".xls"
                checkRosterFileExists(localFileName, monthAdvance)
                'loop for the next months
                For innerCounter = 0 To manipulateDays - 1
                    If totalsArray(innerCounter) > minNoEngineers Then
                        workingDaysArray(arrayCount) = True
                    Else
                        workingDaysArray(arrayCount) = False
                    End If
                    arrayCount = arrayCount + 1
                    manipulateDays = manipulateDays - 1
                    If innerCounter > workingDaysArray.Length - 1 Then
                        'exit this for loop to move to the next month
                        Exit For
                    End If
                    remainingMonths = remainingMonths - 1
                Next
            Next

            'go to the next year
            For monthCounter = 1 To remainingMonths
                monthAdvance = inStartDate.Month + monthCounter
                'check the roster exists for the months regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthCounter, True) & (inEndDate.Year) & ".xls"
                checkRosterFileExists(localFileName, monthCounter)
                'loop for the next months
                For innerCounter = 0 To manipulateDays - 1
                    If totalsArray(innerCounter) > minNoEngineers Then
                        workingDaysArray(arrayCount) = True
                    Else
                        workingDaysArray(arrayCount) = False
                    End If
                    arrayCount = arrayCount + 1
                    manipulateDays = manipulateDays - 1
                    If innerCounter > workingDaysArray.Length - 1 Then
                        'exit this for loop to move to the next month
                        Exit For
                    End If
                Next
            Next

        End If

        'count the number of trues
        For counter = 0 To workingDaysArray.Length - 1
            If workingDaysArray(counter) = True Then
                workingPeriodTotals = workingPeriodTotals + 1
            End If
        Next

        'perform the division
        avgWorkingTotals = workingPeriodTotals / totalDays

        'determine which boolean value to return
        Select Case avgWorkingTotals
            Case Is >= 0.75
                Return True
            Case Else
                Return False
        End Select


    End Function

    Private Sub resetVacationInYear()
        'if the current year isn't equal to the year saved in the textfile
        'then resave the textfile with the new year
        'reset the vacation days using sql.
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        'verification variables
        Dim localCheckDate As Date

        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\VacationYear.txt", OpenMode.Input)
        localCheckDate = DateTime.Parse(LineInput(1))
        FileClose(1)

        'check if the current year is equal to the year in the file
        If localCheckDate.Year <> Now.Year Then
            'write out the new year to the file
            FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\VacationYear.txt", OpenMode.Output)
            PrintLine(1, Now.Date.ToShortDateString)
            FileClose(1)
            'reset the engineer vacation days to 40 <-- though this needs to be expanded in
            'the settings panels to future-proof the program and allow changes
            sql = "UPDATE [Engineers] SET [Engineers].[Vacation_Days_Remaining] = 40 WHERE [Engineers].[Engineer_ID] = ?"
            Dim cmd As New OleDbCommand(sql, connec)
            cmd.CommandType = CommandType.Text
            For counter = 0 To engrIDRosterPostn.Length - 1
                cmd.Parameters.AddWithValue("Engineer_ID", engrIDRosterPostn(counter).engrID)
                cmd.ExecuteNonQuery()
            Next
        End If
        connec.Close()

    End Sub

    Private Sub vacationAdded(ByVal inVactnID As Integer)

        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        'go along the shift roster and total the shifts
        'for each engineer
        'then write each variable to database in parameter query
        sql = "UPDATE [Vacation] SET [Vacation].[Vacation_Added] = TRUE WHERE [Vacation].[Vacation_ID] = ?"
        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@Vacation_ID", inVactnID)
        cmd.ExecuteNonQuery()
        connec.Close()

    End Sub

    Private Sub vacationSubtractDays(ByVal inEngrID As Integer, ByVal inNoDays As Integer)
        'subtract the number of vacation days that have been added
        'from those already in the engineer's record
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim drEngrVactnDays As OleDbDataReader
        Dim dtEngrVactnDays As New DataTable
        Dim sql As String

        'local variable to hold the number of vacation days
        Dim engrRemVactnDays As Integer

        sql = "SELECT [Engineers].[Vacation_Days_Remaining] FROM [Engineers] WHERE [Engineers].[Engineer_ID] = ? "
        cmd = New OleDbCommand(sql, connec)
        cmd.Parameters.AddWithValue("@Engineer_ID", inEngrID)
        connec.Open()
        drEngrVactnDays = cmd.ExecuteReader
        dtEngrVactnDays.Load(drEngrVactnDays)
        connec.Close()

        'assign the value from the dataset
        engrRemVactnDays = dtEngrVactnDays.Rows(0)("Vacation_Days_Remaining")

        'subtract the number of days from the imported data
        engrRemVactnDays = engrRemVactnDays - inNoDays

        'write the value back out
        sql = "UPDATE [Engineers] SET [Engineers].[Vacation_Days_Remaining] = ? WHERE [Engineers].[Engineer_ID] = ?"
        cmd = New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@Vacation_Days_Remaining", engrRemVactnDays)
        cmd.Parameters.AddWithValue("@Engineer_ID", inEngrID)
        cmd.ExecuteNonQuery()
        connec.Close()

    End Sub

    Private Sub MgmtSettChangePasswordBut_Click(sender As Object, e As EventArgs) Handles MgmtSettChangePasswordBut.Click
        'allows the user to change the password
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim cmd As New OleDbCommand
        Dim drMgmtPw As OleDbDataReader
        Dim dtMgmtPw As New DataTable
        Dim sql As String

        'comparison variables
        Dim hashValueFromDB As String
        Dim compareHashValue As String

        'new password variables
        Dim newPassword As String
        Dim checkNewPassword As String
        Dim checkIfSame As String

        'assign the values from the textbox
        newPassword = MgmtSettNewPasswordTB.Text
        checkNewPassword = MgmtSettNewPasswordConfirmTB.Text

        'check if they contain a value
        If IsNothing(newPassword) Or IsNothing(checkNewPassword) Then
            MessageBox.Show("The new password must be a minimum of 8 characters", "LEMS Warning: Incorrect Password Length", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        'get the hash from the database
        sql = "SELECT [Management_Password].[Password] FROM [Management_Password] WHERE [[Management_Password].[ID] = 1"
        cmd = New OleDbCommand(sql, connec)
        connec.Open()
        drMgmtPw = cmd.ExecuteReader
        dtMgmtPw.Load(drMgmtPw)
        connec.Close()

        'copy the hash from the database into a local variable
        hashValueFromDB = dtMgmtPw.Rows(0)("Password")

        'get the value from the textbox
        compareHashValue = generateHashForMgmtPassword(MgmtSettOldPasswordTB.Text)

        'compare the old hash with the new hash
        If compareHashValue = hashValueFromDB Then
            'check if the new password and the confirm new password are the same
            checkIfSame = compareNewPasswords(newPassword, checkNewPassword)
            If checkIfSame = True Then
                'update the password
                sql = "UPDATE [Management_Password] SET [Management_Password].[Password]=?"
                cmd = New OleDbCommand(sql, connec)
                cmd.CommandType = CommandType.Text
                connec.Open()
                cmd.Parameters.AddWithValue("@Password", generateHashForMgmtPassword(newPassword))
                connec.Close()
                'message to alert the user the password has been changed
                MessageBox.Show("The password has been successfully updated", "LEMS Message: Password Updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            'message to alert the user the old password is incorrect
            MessageBox.Show("The original password is incorrect. Enter another", "LEMS Warning: Incorrect Password", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub

    Public Function generateHashForMgmtPassword(ByVal inPassword As String) As String
        'code taken and adapted for use from
        'http://support.microsoft.com/kb/301053
        'after a Google search for 'vb.net hashing'

        Dim mgmtPassValue As String
        Dim tmpSource() As Byte
        Dim tmpHash() As Byte
        Dim mgmtHashedValue As String

        mgmtPassValue = inPassword

        'Create a byte array from source data.
        tmpSource = ASCIIEncoding.ASCII.GetBytes(mgmtPassValue)

        'Compute hash based on source data.
        tmpHash = New MD5CryptoServiceProvider().ComputeHash(tmpSource)

        'get the hex string from the hash
        mgmtHashedValue = ByteArrayToString(tmpHash)

        'return the value from the hash
        Return mgmtHashedValue

    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String
        'returns as 
        'code taken and adapted for use from
        'http://support.microsoft.com/kb/301053
        'after a Google search for 'vb.net hashing'

        Dim i As Integer
        Dim sOutput As New StringBuilder(arrInput.Length)
        For i = 0 To arrInput.Length - 1
            sOutput.Append(arrInput(i).ToString("X2"))
        Next
        Return sOutput.ToString()
    End Function

    Private Function compareNewPasswords(ByVal inNewPass As String, ByVal inNewConfirmPass As String) As Boolean
        'check to see if the new password and the confirm new passwords are the same
        Dim localNewHash As String
        Dim localNewHashConfirm As String

        'hash the values
        localNewHash = generateHashForMgmtPassword(inNewPass)
        localNewHashConfirm = generateHashForMgmtPassword(inNewConfirmPass)

        'compare the values
        If localNewHash = localNewHashConfirm Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub ticShifts()
        'fill the roster with a small number of tic shifts for the engineers that
        'are qualified to be the tic
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dsTicEngrs As New DataSet
        Dim dAdapTicEngrs As New OleDbDataAdapter
        Dim sql As String
        'local control variables
        Dim localEngrIDArray() As Integer
        Dim columnCounter As Integer
        Dim advanceCount As Integer
        Dim randomStart As Integer
        Dim valueString As String

        'select the engineers that can be the tic
        sql = "SELECT [Engineers].[Engineer_ID] FROM [Engineers] WHERE [Can_Be_Tic]"
        Dim cmd As New OleDbCommand(sql, connec)
        dAdapTicEngrs = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapTicEngrs.Fill(dsTicEngrs, "Engineers")

        localEngrIDArray = (From myRow In dsTicEngrs.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Engineer_ID")).ToArray

        'local structure to hold the tic engineer ids
        Dim localEngrTIC(dsTicEngrs.Tables(0).Rows.Count) As tEngrIDAndRostPostn

        'copy the IDs to the structure from the array
        For counter = 0 To dsTicEngrs.Tables(0).Rows.Count - 1
            localEngrTIC(counter).engrID = localEngrIDArray(counter)
            localEngrTIC(counter).engrRostPostn = returnEngrRostPostn(localEngrIDArray(counter))
        Next

        'allocate the tic shifts
        'loop through the datagridview
        'allocate tic shifts for the month
        'by generating a random position
        'add three tic shifts where there are no off days

        Randomize()
        For overCount = 0 To localEngrTIC.Length - 1
            randomStart = CInt(Int((System.DateTime.DaysInMonth(Now.Year, Now.Month)) * Rnd()) + 1)
            advanceCount = 0
            Do
                columnCounter = randomStart
                valueString = EngrRosterSpace.Rows(localEngrTIC(overCount).engrRostPostn).Cells(columnCounter).Value
                If valueString = "R" Or valueString = "V" Or valueString.Contains("TRNG") Or valueString.Contains("MED") Or valueString.Contains("TDY") Or valueString - "SICK" Then
                    columnCounter = columnCounter + 1
                    advanceCount = advanceCount - 1
                Else
                    EngrRosterSpace.Rows(localEngrTIC(overCount).engrRostPostn).Cells(columnCounter).Value = "TIC"
                    advanceCount = advanceCount + 1
                End If
            Loop Until advanceCount = 3
        Next

    End Sub

    Private Sub MgmtItemDSArchiveSearchTB_Click(sender As Object, e As EventArgs) Handles MgmtItemDSArchiveSearchTB.Click
        MgmtItemDSArchiveSearchTB.Clear()
    End Sub

    Private Sub MgmtItemAdvSearchDSArchiveTB_Click(sender As Object, e As EventArgs) Handles MgmtItemAdvSearchDSArchiveTB.Click
        MgmtItemAdvSearchDSArchiveTB.Clear()
    End Sub

    Private Sub MgmtAircraftNoDSArchiveSearchTB_Click(sender As Object, e As EventArgs) Handles MgmtAircraftNoDSArchiveSearchTB.Click
        MgmtAircraftNoDSArchiveSearchTB.Clear()
    End Sub

    Private Sub MgmtShipNoAdvSearchDSArchiveTB_Click(sender As Object, e As EventArgs) Handles MgmtShipNoAdvSearchDSArchiveTB.Click
        MgmtShipNoAdvSearchDSArchiveTB.Clear()
    End Sub

    Private Sub MgmtPrintRostBut_Click(sender As Object, e As EventArgs) Handles MgmtPrintRostBut.Click
        Dim localRosterFileName As String
        localRosterFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(Month(advanceRosterMonth), True) & advanceRosterYear.Year & ".xls"
        printRoster(localRosterFileName)
    End Sub

    Public Sub printRoster(ByVal inFileName As String)
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkBooks As Excel.Workbooks
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim workSheetRange As Excel.Range

        'copied and adapted code
        xlApp.Visible = True
        xlWorkBooks = xlApp.Workbooks
        xlWorkBook = xlWorkBooks.Open(inFileName)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        workSheetRange = xlWorkSheet.UsedRange
        xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        xlWorkSheet.PrintPreview()
        'xlWorkSheet.PrintOutEx()

        'close the excel files
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(workSheetRange)
        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkBooks)
        releaseObject(xlApp)

    End Sub

    Private Sub MgmtSettMinAMTCountTB_KeyDown(sender As Object, e As KeyEventArgs) Handles MgmtSettMinAMTCountTB.KeyDown
        'this routine is used to store the minimum number of engineers
        'in a textfile
        Dim localMinNoString As String
        Dim localMinNoInt As Integer
        If e.KeyCode = Keys.Enter Then
            'gets the value of the textbox
            localMinNoString = MgmtSettMinAMTCountTB.Text
            If Not IsNumeric(localMinNoString) Then
                'display an error message if non-numeric data
                MessageBox.Show("The value entered is not a number. Enter a numeric value.", "LEMS Warning: Invalid character", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                'write the number to the file
                localMinNoInt = CInt(localMinNoString)
                FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\ShiftColours\AMTMinimums.txt", OpenMode.Output)
                PrintLine(1, localMinNoInt)
                FileClose(1)
            End If
        End If

    End Sub

    Private Sub loadMinEngineers()
        'opens and saves the content of the minumum engineers file to a global program variable
        FileOpen(1, "C:\Users\zande_000\Documents\LEMSSystem\AMTMinimums.txt", OpenMode.Input)
        minNoEngineers = CInt(LineInput(1))
        MgmtSettMinAMTCountTB.Text = CStr(minNoEngineers)
        FileClose(1)
    End Sub

    Private Sub createMonthlyHours()
        'this routine fills the monthly hours viewer with the shifts

        'variables to create the columns
        Dim columnToAdd As New DataGridViewTextBoxColumn

        'add the columns to the shift viewer
        columnToAdd.HeaderText = "Name"
        columnToAdd.Name = "MonthlyHoursNameCol"
        MgmtEngrMonthlyHoursDGV.Columns.Add(columnToAdd)

        'add multiple columns for each day
        For colCounter = 1 To System.DateTime.DaysInMonth(Year(Now), Month(Now))
            columnToAdd = New DataGridViewTextBoxColumn
            columnToAdd.HeaderText = "Hours on " & colCounter
            columnToAdd.Name = "MonthlyHoursDateCol" & colCounter
            MgmtEngrMonthlyHoursDGV.Columns.Add(columnToAdd)
            columnToAdd = New DataGridViewTextBoxColumn
            columnToAdd.HeaderText = "Shift"
            columnToAdd.Name = "MonthlyHoursShiftCol" & colCounter
            MgmtEngrMonthlyHoursDGV.Columns.Add(columnToAdd)
        Next

    End Sub

    Private Sub addShiftsToMonthlyHours()
        'this routine adds the shifts from the roster to the monthly hours sheet
        Dim addingValue As Integer

        'loop through the roster and add the shifts
        For rowcounter = 0 To EngrRosterSpace.Rows.Count - 1
            addingValue = 2
            For colCounter = 1 To EngrRosterSpace.Columns.Count - 1
                MgmtEngrMonthlyHoursDGV.Rows(rowcounter).Cells(addingValue).Value = EngrRosterSpace.Rows(rowcounter).Cells(colCounter).Value
                MgmtEngrMonthlyHoursDGV.Rows(rowcounter).Cells(addingValue).Style.BackColor = EngrRosterSpace.Rows(rowcounter).Cells(colCounter).Style.BackColor
                addingValue = addingValue + 2
            Next
        Next

        'resize the columns
        MgmtEngrMonthlyHoursDGV.AutoResizeColumns()

        'export the data to Excel
        exportMonthlyHoursToExcel()

    End Sub

    Private Sub exportMonthlyHoursToExcel()
        'the following code was adapted from:
        'http://vb.net-informations.com/excel-2007/vb.net_excel_2007_create_file.htm
        'and
        'http://stackoverflow.com/questions/6983141/changing-cell-color-of-excel-sheet-via-vb-net

        'Excel file variables
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlworkBooks As Excel.Workbooks
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim misValue As Object = System.Reflection.Missing.Value
        'my variables
        Dim colCount As Integer
        Dim monthlyHoursFilename As String

        'my code
        colCount = System.DateTime.DaysInMonth(Year(Now), Month(Now))

        'copied code, sets up the Excel file
        xlApp.DisplayAlerts = False
        xlworkBooks = xlApp.Workbooks
        xlWorkBook = xlworkBooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        'below is my code
        'this sets the rows headers
        xlWorkSheet.Cells(1, 1) = "Name"
        For counter = 1 To colCount
            xlWorkSheet.Cells(1, counter + 1) = counter
        Next
        'this writes the data to the file from the datagridview
        'also writes the colour of the cell to the file
        For rowCounter = 0 To EngrRosterSpace.RowCount - 2
            For columnCounter = colCount To 1 Step -1
                xlWorkSheet.Cells(rowCounter + 2, columnCounter + 1) = EngrRosterSpace.Rows(rowCounter).Cells(columnCounter).Value
            Next
        Next
        For rowCounter = 0 To EngrRosterSpace.RowCount - 2
            xlWorkSheet.Cells(rowCounter + 2, 1) = EngrRosterSpace.Rows(rowCounter).Cells(0).Value
        Next

        monthlyHoursFilename = "C:\Users\zande_000\Documents\LEMSMonthlyHours" & "\MonthlyHoursFor" & rosterMonth & rosterYear & ".xls"

        'this is the copied code
        xlWorkBook.SaveAs(monthlyHoursFilename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)

        xlWorkBook.Close(True, misValue, misValue)

        'this closes the application as it releases COM objects
        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlworkBooks)
        xlApp.Quit()
        releaseObject(xlApp)

        MsgBox("Monthly Hours File was saved as " & monthlyHoursFilename)
    End Sub

    Private Sub MgmtSettApprvReqs_Click(sender As Object, e As EventArgs) Handles MgmtSettApprvReqs.Click
        'updates the selected request in the datagridview to approved
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        'set the sql string
        sql = "UPDATE [Requests] SET [Requests].[Request_Approved] = TRUE WHERE [Requests].[REQ_ID] = ?"

        'add the parameters to the sql statement
        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.AddWithValue("@REQ_ID", MgmtMngEngrsPendingRqstsDGV.SelectedRows(0).Cells(0).Value)
        connec.Open()
        'execute the query
        cmd.ExecuteNonQuery()
        'close the connection
        connec.Close()

    End Sub

    Private Sub MgmtMngEngrsTDYDoneBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsTDYDoneBut.Click
        'this adds TDY to the database for the selected engineer
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dAdapMostRecntTDY As OleDbDataAdapter
        Dim dsMostRecntTDY As New DataSet
        Dim sql As String
        Dim startDate As Date
        Dim endDate As Date
        Dim stringStartDate As String
        Dim stringEndDate As String
        Dim linkTDYID As Integer

        'carry out the date assignments
        startDate = MgmtMngEngrsTDYStartDateDTP.Value
        endDate = MgmtMngEngrsTDYEndDateDTP.Value

        'check if the start date is before the current date
        If startDate < Now.Date Or endDate < Now.Date Then
            MessageBox.Show("TDY cannot be added retrospectively to rosters.", "LEMS Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'convert the dates into strings for the database
            stringStartDate = EngrLEMSProcess.returnStringDate(startDate)
            stringEndDate = EngrLEMSProcess.returnStringDate(startDate)

            'set up the database connection and set the SQL string
            sql = "INSERT INTO [TDY] SET [TDY].[TDY_Start_Date], [TDY].[TDY_End_Date], [TDY].[TDY_Location] VALUES (?,?,?)"
            Dim cmd As New OleDbCommand(sql, connec)
            'add the parameters
            cmd.Parameters.AddWithValue("@TDY_Start_Date", stringStartDate)
            cmd.Parameters.AddWithValue("@TDY_End_Date", stringEndDate)
            cmd.Parameters.AddWithValue("@TDY_Location", MgmtMngEngrsTDYLocTB.Text)
            connec.Open()
            'add the TDY to the database
            cmd.ExecuteNonQuery()
            connec.Close()

            'get the most recent addition to the TDY table
            sql = "SELECT MAX([TDY].[TDY_ID] AS tdyID FROM [TDY]"
            dAdapMostRecntTDY = New OleDbDataAdapter(sql, connec)
            connec.Open()
            dAdapMostRecntTDY.Fill(dsMostRecntTDY, "TDY")
            connec.Close()

            'copy the value to the variable
            linkTDYID = CInt(dsMostRecntTDY.Tables(0).Rows(0)("tdyID"))

            'update the link table with the engineer ID and the TDY ID
            sql = "UPDATE [EN/TDY] SET [EN/TDY].[EN_ID] = ?, [EN/TDY].[TDY_ID] = ?"
            cmd = New OleDbCommand
            cmd.Parameters.AddWithValue("@EN_ID", CInt(engineerID))
            cmd.Parameters.AddWithValue("@TDY_ID", linkTDYID)
            connec.Open()
            cmd.ExecuteNonQuery()
            connec.Close()

            MgmtMngEngrsTDYPanel.Hide()
        End If
        
    End Sub

    Private Sub TDYAdded(ByVal inTDYID As Integer)

        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim sql As String

        'update the TDY after it has been added to the roster
        sql = "UPDATE [TDY] SET [TDY].[TDY_Added] = TRUE WHERE [TDY].[TDY_ID] = ?"
        Dim cmd As New OleDbCommand(sql, connec)
        cmd.CommandType = CommandType.Text
        connec.Open()
        cmd.Parameters.AddWithValue("@TDY_ID", inTDYID)
        cmd.ExecuteNonQuery()
        connec.Close()

    End Sub

    Private Sub loadTDY()
        'load the engineer vacation here that hasn't been filled
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = "C:\Users\zande_000\Documents\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim cmd As New OleDbCommand
        Dim connec As OleDbConnection = New OleDbConnection(DBconn & DBsource)
        Dim dsNonAddedTDY As New DataSet
        Dim dAdapNonAddedTDY As New OleDbDataAdapter
        Dim sql As String

        'sql string, orders it by preference- 1 being the highest preference and 3 being the lowest
        'so the dataset will contain smallest to largest
        sql = "SELECT * FROM [TDY] INNER JOIN [EN/TDY] ON [TDY].[TDY_ID] = [EN/TDY].[TDY_ID] WHERE [TDY].[TDY_Added] = FALSE"
        cmd = New OleDbCommand(sql, connec)
        dAdapNonAddedTDY = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapNonAddedTDY.Fill(dsNonAddedTDY, "TDY")
        connec.Close()

        'call the structure routine if the dataset is not empty
        If dsNonAddedTDY.Tables(0).Rows.Count > 0 Then
            structureTDY(dsNonAddedTDY)
        End If
    End Sub

    Private Sub structureTDY(ByRef inDataSet As DataSet)
        'put the TDY into a structure so it can be manipulated by the filling routine
        'arrays to create from the dataset and then fill the roster
        Dim localEngrID() As Integer
        Dim localTDYID() As Integer
        Dim localTDYStartDateString() As String
        Dim localTDYEndDateString() As String
        Dim localTDYLoc() As String


        'fill the arrays from the dataset
        localEngrID = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("EN_ID")).ToArray
        localTDYID = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("TDY.TDY_ID")).ToArray
        localTDYStartDateString = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of String)("TDY_Start_Date")).ToArray
        localTDYEndDateString = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of String)("TDY_End_Date")).ToArray
        localTDYLoc = (From myRow In inDataSet.Tables(0).AsEnumerable Select myRow.Field(Of String)("TDY_Location")).ToArray

        'declare the structure used to store the engineer and vacation details
        Dim localEngrTDY(localEngrID.Length - 1) As tEngrTDY

        'fill the structure
        For counter = 0 To localEngrID.Length - 1
            localEngrTDY(counter).engrID = localEngrID(counter)
            localEngrTDY(counter).engrRostPosn = returnEngrRostPostn(localEngrTDY(counter).engrID)
            localEngrTDY(counter).engrTDYID = localTDYID(counter)
            'directly parse the string dates to System.DateTime dates so that the functions
            'associated with DateTime can be used (so the number of days can be caluclated etc)
            localEngrTDY(counter).engrTDYStartDate = DateTime.Parse(localTDYStartDateString(counter))
            localEngrTDY(counter).engrTDYEndDate = DateTime.Parse(localTDYEndDateString(counter))
            localEngrTDY(counter).engrTDYLoc = localTDYLoc(counter)
        Next

        fillRosterWithTDY(localEngrTDY)

    End Sub

    Private Sub fillRosterWithTDY(ByRef inTDYStruct() As tEngrTDY)
        'fill the roster with the TDY for specific engineers

        Dim inStartDate As Date
        Dim inEndDate As Date

        'total days of vacation
        Dim totalDays As Integer
        Dim manipulateDays As Integer

        'variables to hold the differences between the months and years
        Dim monthDifference As Integer
        Dim monthAdvance As Integer
        Dim remainingMonths As Integer

        'create the array to analyse 
        Dim workingDaysArray(totalDays) As Boolean

        'control variables
        Dim arrayCount As Integer
        Dim localFileName As String
        Dim checkingCount As Integer

        For overCounter = 0 To inTDYStruct.Length - 1
            

            inStartDate = inTDYStruct(overCounter).engrTDYStartDate
            inEndDate = inTDYStruct(overCounter).engrTDYEndDate

            totalDays = (inEndDate - inStartDate).Days
            manipulateDays = totalDays
            'gets the number of months between the two dates
            monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))

            'analyse the boolean array
            'different routines depending if the vacation is in the same month
            arrayCount = 0
            If inStartDate.Month = inEndDate.Month Then
                'check the roster exists for the month regarding the vacation
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
                checkRosterFileExists(localFileName, CInt(inStartDate.Month))

                For counter = inStartDate.Day To inEndDate.Day
                    EngrRosterSpace.Rows(inTDYStruct(overCounter).engrRostPosn).Cells(counter).Value = "TDY at " & inTDYStruct(overCounter).engrTDYLoc
                    arrayCount = arrayCount + 1
                Next
                exportRosterToExcel(localFileName)
            End If

            'separate if statements to simplify code layout
            'checks the workdays between non-contiguous months in the same year
            If inStartDate.Month <> inEndDate.Month And inStartDate.Year = inEndDate.Year Then
                monthDifference = inEndDate.Month - inStartDate.Month
                'find if the minimum shift numbers have been met for the first month
                For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                    EngrRosterSpace.Rows(inTDYStruct(overCounter).engrRostPosn).Cells(innerCounter).Value = "TDY at " & inTDYStruct(overCounter).engrTDYLoc
                    arrayCount = arrayCount + 1
                    manipulateDays = manipulateDays - 1
                Next
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
                exportRosterToExcel(localFileName)
                For counter = 1 To monthDifference
                    'check the roster exists for the months regarding the vacation
                    localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month + counter, True) & inStartDate.Year & ".xls"
                    checkRosterFileExists(localFileName, CInt(inStartDate.Month))
                    'loop for the next months
                    checkingCount = 0
                    For innerCounter = 0 To manipulateDays
                        checkingCount = checkingCount + 1
                        If checkingCount > System.DateTime.DaysInMonth(inStartDate.Year, CInt(inStartDate.Month + counter)) Then
                            checkingCount = 1
                            Exit For
                        End If
                        EngrRosterSpace.Rows(inTDYStruct(overCounter).engrRostPosn).Cells(checkingCount).Value = "TDY at " & inTDYStruct(overCounter).engrTDYLoc
                        arrayCount = arrayCount + 1
                        manipulateDays = manipulateDays - 1
                        If innerCounter > totalDays Then
                            'exit this for loop to move to the next month
                            Exit For
                        End If
                    Next
                    exportRosterToExcel(localFileName)
                Next
            End If

            'checks the workdays between non-contigous years
            If inStartDate.Year <> inEndDate.Year Then
                'get the differences between the vacation
                monthDifference = CInt(DateDiff(DateInterval.Month, inStartDate, inEndDate))
                'check the first month
                For innerCounter = inStartDate.Day To System.DateTime.DaysInMonth(inStartDate.Year, inStartDate.Month)
                    EngrRosterSpace.Rows(inTDYStruct(overCounter).engrRostPosn).Cells(innerCounter).Value = "TDY at " & inTDYStruct(overCounter).engrTDYLoc
                    arrayCount = arrayCount + 1
                    'manipulateDays = manipulateDays - 1
                Next
                localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(inStartDate.Month, True) & inStartDate.Year & ".xls"
                exportRosterToExcel(localFileName)
                'continue the loop until the month advance is greater than 12
                remainingMonths = monthDifference
                For monthCounter = 1 To monthDifference
                    monthAdvance = inStartDate.Month + monthCounter
                    'resets the months when the count is greater than 12
                    'this signifies that the new year analysis has been started
                    If monthAdvance > 12 Then
                        Exit For
                    End If
                    'check the roster exists for the months regarding the vacation
                    localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & inStartDate.Year & ".xls"
                    checkRosterFileExists(localFileName, monthAdvance)
                    'loop for the next months
                    For innerCounter = 0 To manipulateDays - 1
                        EngrRosterSpace.Rows(inTDYStruct(overCounter).engrRostPosn).Cells(innerCounter).Value = "TDY at " & inTDYStruct(overCounter).engrTDYLoc
                        arrayCount = arrayCount + 1
                        'manipulateDays = manipulateDays - 1
                        If innerCounter > totalDays Then
                            'exit this for loop to move to the next month
                            Exit For
                        End If
                        remainingMonths = remainingMonths - 1
                    Next

                    exportRosterToExcel(localFileName)
                Next

                'go to the next year
                For monthCounter = 1 To remainingMonths
                    monthAdvance = inStartDate.Month + monthCounter
                    If monthAdvance > 12 Then
                        monthAdvance = monthCounter
                    End If
                    'check the roster exists for the months regarding the vacation
                    localFileName = "C:\Users\zande_000\Documents\LEMSRosters" & "\ShiftRosterFor" & MonthName(monthAdvance, True) & (inEndDate.Year) & ".xls"
                    checkRosterFileExists(localFileName, monthAdvance)
                    'loop for the next months
                    For innerCounter = 0 To manipulateDays - 1
                        EngrRosterSpace.Rows(inTDYStruct(overCounter).engrRostPosn).Cells(innerCounter + 1).Value = "TDY at " & inTDYStruct(overCounter).engrTDYLoc
                        arrayCount = arrayCount + 1
                        'manipulateDays = manipulateDays - 1
                        If innerCounter > totalDays Then
                            'exit this for loop to move to the next month
                            Exit For
                        End If
                    Next
                    exportRosterToExcel(localFileName)
                Next

            End If

            'write out to the database that the TDY has been added
            TDYAdded(inTDYStruct(overCounter).engrTDYID)

        Next
        importFromExcelFile(rosterFileName)
    End Sub

    Private Sub MgmtMngEngrsAddTDYBut_Click(sender As Object, e As EventArgs) Handles MgmtMngEngrsAddTDYBut.Click
        MgmtMngEngrsTDYPanel.Show()
    End Sub

    Private Sub MgmtEngrSignSheetAddChangesBut_Click(sender As Object, e As EventArgs) Handles MgmtEngrSignSheetAddChangesBut.Click
        'this adds the changes to the sign in sheet

        'add the shift change
        If MgmtSignSheetShiftComB.SelectedIndex <> -1 Then
            MgmtEngrSignInSheetDGV.CurrentRow.Cells(1).Value = MgmtSignSheetShiftComB.SelectedItem
        End If

        'swap with a selected ngineer
        If MgmtSwapEngrsListBox.SelectedIndex <> -1 Then
            MgmtEngrSignInSheetDGV.CurrentRow.Cells(0).Value = MgmtSwapEngrsListBox.SelectedItem
        End If

        'reset the check box
        MgmtOverTimeCB.CheckState = CheckState.Unchecked

    End Sub

    Private Sub CDST4DailySheetDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles CDST4DailySheetDGV.CellClick
        'code taken and adapted for use from:
        'http://stackoverflow.com/questions/24320267/column-specific-click-event-in-datagridview-vb-net

        'create a variable to hold the cell index
        Dim cellIndex As Integer

        cellIndex = 1
        'show the panel if the cell index is the index for the AMTs column
        If e.ColumnIndex = 8 Then
            CDST4AddEngrPanel.Show()
        End If

    End Sub

    Private Sub CDST4AddSelectedEngrBut_Click(sender As Object, e As EventArgs) Handles CDST4AddSelectedEngrBut.Click
        'adds the selected engineer to the Daily Sheet datagridview

        'check if there is an engineer selected
        If CDST4AddEngrPanelLB.SelectedIndex = -1 Then
            MessageBox.Show("An engineer was not selected to add to the Daily Sheet. To add an engineer, select one from the list", "LEMS Warning: An Engineer was not selected from the list", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'add the selected engineer to the datagridview
            CDST4DailySheetDGV.CurrentCell.Value = CDST4AddEngrPanelLB.SelectedItem
        End If

    End Sub

    Private Sub CDST3DailySheetDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles CDST3DailySheetDGV.CellClick
        'create a variable to hold the cell index
        Dim cellIndex As Integer

        cellIndex = 1
        'show the panel if the cell index is the index for the AMTs column
        If e.ColumnIndex = 8 Then
            CDST3AddEngrPanel.Show()
        End If
    End Sub

    Private Sub CDST3AddSelectedEngrBut_Click(sender As Object, e As EventArgs) Handles CDST3AddSelectedEngrBut.Click
        'adds the selected engineer to the Daily Sheet datagridview

        'check if there is an engineer selected
        If CDST3AddEngrPanelLB.SelectedIndex = -1 Then
            MessageBox.Show("An engineer was not selected to add to the Daily Sheet. To add an engineer, select one from the list", "LEMS Warning: An Engineer was not selected from the list", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'add the selected engineer to the datagridview
            CDST3DailySheetDGV.CurrentCell.Value = CDST3AddEngrPanelLB.SelectedItem
        End If
    End Sub

    Private Sub CDST4ClearEngrsBut_Click(sender As Object, e As EventArgs) Handles CDST4ClearEngrsBut.Click
        'clear the AMTs cell
        CDST4DailySheetDGV.CurrentCell.Value = ""
    End Sub

    Private Sub CDST3ClearEngrsBut_Click(sender As Object, e As EventArgs) Handles CDST3ClearEngrsBut.Click
        'clear the AMTs cell
        CDST3DailySheetDGV.CurrentCell.Value = ""
    End Sub

    Private Sub CDST4DoneBut_Click(sender As Object, e As EventArgs) Handles CDST4DoneBut.Click
        'hide the panel
        CDST4AddEngrPanel.Hide()
    End Sub

    Private Sub CDST3DoneBut_Click(sender As Object, e As EventArgs) Handles CDST3DoneBut.Click
        'hide the panel
        CDST3AddEngrPanel.Hide()
    End Sub
End Class
'call a new class or another code block for the daily sheet so that is can be typed once
'and reused by the engineer side
'the same needs doing for opening the roster file

'need to hide the panels on startup