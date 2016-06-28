'This is a public module that allows both the engineer side and the management side to query the daily sheet archive
'it returns the query and processes the data according to what the user has specified
'it will either return metrics according to the search
'or the actual daily sheet file that can be viewed in the pdf viewer
Option Explicit On
Imports System.Environment
Imports System.Windows.Forms
Imports System.Globalization
Imports System.IO
Imports System.Data.OleDb
Imports System.Data
Imports System.Text
Imports System.Linq
Imports System.Web.UI.WebControls

Module DailySheetArchiveSearch
    Public resultsReturnString As String
    Public resultsReturnList() As String

    Public Sub mainControl(ByVal searchDSARequestType As String, ByVal smpDSASearchKeywords As String, ByVal smpDSASearchShipNo As String, ByVal smpDSASearchDate As Date, ByVal advDSASearchKeywords As String, ByVal advDSASearchShipNo As String, ByVal advDSASearchLoctn As String, ByVal advDSASearchFlightInNo As String, ByVal advSearchFlightOutNo As String, ByVal advDSASearchAMTs As String, ByVal advDSASearchDate As String)
        'determine which subroutines to go to depending on the state of
        'the controls on the forms
        'this needs to function with both the management and engineer sides!
        'possibly use a traceback or parameter to determine which side called the routine- actually this is not needed as it returns a results string or an array depending on the request
        'there must be a date to string conversion on the management and engineer side when the search button is clicked so that the query can function correctly

        Select Case searchDSARequestType
            Case Is = "smpDSASearch"
                resultsReturnString = simpleArchiveSearch(smpDSASearchKeywords, smpDSASearchShipNo, smpDSASearchDate)
            Case Is = "advDSASearchMets"
                resultsReturnString = advancedArchiveSearchMetrics(advDSASearchKeywords, advDSASearchShipNo, advDSASearchLoctn, advDSASearchFlightInNo, advSearchFlightOutNo, advDSASearchAMTs, advDSASearchDate)
            Case Is = "advDSASearchActDS"
                resultsReturnList = advancedArchiveSearchActualDS(advDSASearchKeywords, advDSASearchShipNo, advDSASearchLoctn, advDSASearchFlightInNo, advSearchFlightOutNo, advDSASearchAMTs, advDSASearchDate)
        End Select
    End Sub

    Private Function simpleArchiveSearch(ByVal smpDSASearchKeywords As String, ByVal smpDSASearchShipNo As String, ByVal smpDSASearchDate As String) As String
        'return simple metrics and list of daily sheets
        'add the parameters based on whether the strings are NOT empty
        Dim dsaSmpSearchResString As String
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim dAdapSmpSearch As OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim dsSmpSearchResults As New DataSet
        Dim dsSmpSearchExtraDateResults As New DataSet
        Dim dsSmpSearchExtraDateResultsEngrs As New DataSet
        Dim sql As String
        Dim sqlTotal As Integer
        Dim tempSqlString As String
        Dim commentsBool As Boolean
        Dim shipBool As Boolean
        Dim dateBool As Boolean

        'local arrays and variables to store data from the archive for processing
        Dim localCommentsArray() As String
        Dim localShipNoArray() As String
        Dim localDateArray() As String
        Dim localFlightInArray() As String
        Dim localFlightOutArray() As String
        Dim localTimeInArray() As String
        Dim localTimeOutArray() As String
        Dim localGateArray() As String
        Dim localLogItemsArray() As Integer
        Dim localTotalLogItems As Integer
        Dim localAmtsOnAircraftArray() As String
        Dim localSheetID As Integer
        Dim tempForename() As String
        Dim tempSurname() As String

        'sets up the SQL for the search
        sql = "SELECT * FROM [Sheet], [Entries] WHERE "
        cmd = New OleDbCommand(sql, connec)
        'checks to see if the parameters are present, if not, they are not added to the sql string
        'the sql string is 'built' depending on the values of the search paramters
        If smpDSASearchKeywords <> Nothing Then
            sql = sql & "[Entries].[Comments] LIKE %?% AND"
            cmd.Parameters.AddWithValue("@Comments", smpDSASearchKeywords)
            sqlTotal = sqlTotal + 1
            commentsBool = True
        End If
        If smpDSASearchShipNo <> Nothing Then
            sql = sql & " [Entries].[Ship] = ? "
            cmd.Parameters.AddWithValue("@Ship", smpDSASearchShipNo)
            sqlTotal = sqlTotal + 2
            shipBool = True
        End If
        If smpDSASearchDate <> Nothing Then
            sql = sql & "AND [Sheet].[Date] = ?"
            cmd.Parameters.AddWithValue("@Sheet_Date", smpDSASearchDate)
            sqlTotal = sqlTotal + 4
            dateBool = True
        End If

        'determine which SQL string to create and removes the AND should 
        Select Case sqlTotal
            Case Is = 1
                'only comments
                tempSqlString = sql.Replace("AND", "")
                sql = tempSqlString
            Case Is = 4
                'only the date
                tempSqlString = sql.Replace("AND", "")
                sql = tempSqlString
        End Select

        dAdapSmpSearch = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapSmpSearch.Fill(dsSmpSearchResults, "Entries")
        connec.Close()

        'check if no results were returned
        If dsSmpSearchResults.Tables(0).Rows.Count < 1 Then
            dsaSmpSearchResString = "No Results"
        Else
            '\\\\\\\\\\\\\\\\\\\\\
            'generate the metrics
            '/////////////////////
            dsaSmpSearchResString = "The following results were returned:" & vbCrLf
            'puts the results of the dataset into local arrays for processing
            localCommentsArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Commments")).ToArray
            localShipNoArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Ship")).ToArray
            localDateArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Sheet_Date")).ToArray
            localFlightInArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Flt_In")).ToArray
            localFlightOutArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Flt_Out")).ToArray
            localTimeInArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Time_In")).ToArray
            localTimeOutArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Time_Out")).ToArray
            localGateArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Gate")).ToArray
            localLogItemsArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Log_Items")).ToArray
            localAmtsOnAircraftArray = (From myRow In dsSmpSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("AMTs")).ToArray
            'used to query the link table between the engineers and the daily sheet table
            localSheetID = CInt(dsSmpSearchResults.Tables(0).Rows(0)("Sheet_ID"))

            Select Case sqlTotal
                Case Is = 1
                    'just comments
                    dsaSmpSearchResString = dsaSmpSearchResString & "The keyword search term " & smpDSASearchKeywords & " appeared " & localCommentsArray.Length & " times."
                Case Is = 2
                    'just ship number
                    dsaSmpSearchResString = dsaSmpSearchResString & "The ship " & smpDSASearchShipNo & " appeared " & localShipNoArray.Length & " times."
                Case Is = 3
                    'comments and ship number
                    dsaSmpSearchResString = dsaSmpSearchResString & "The keyword search term " & smpDSASearchKeywords & " regarding ship no " & smpDSASearchShipNo & " appears " & localCommentsArray.Length & " times."
                Case Is = 4
                    'just the date
                    'returns extra info using the date such as the engineers that are on a particular
                    'daily sheet
                    sql = "SELECT * FROM [EN/SHEET] WHERE [SHEET_ID] = " & localSheetID
                    cmd = New OleDbCommand(sql, connec)
                    dAdapSmpSearch = New OleDbDataAdapter(sql, connec)
                    connec.Open()
                    dAdapSmpSearch.Fill(dsSmpSearchExtraDateResults, "EN/SHEET")
                    connec.Close()

                    Dim localEngrIdents() As Integer = (From myRow In dsSmpSearchExtraDateResults.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("EN_ID")).ToArray
                    Dim queryTempString As String = String.Join(",", localEngrIdents)
                    'gets all of the AMTs who worked on the day using the engineer ID
                    sql = "SELECT [Forename], [Surname] FROM [Engineers] WHERE [Engineers].[Engineer_ID] = ?"
                    cmd = New OleDbCommand(sql, connec)
                    cmd.Parameters.AddWithValue("@Engineer_ID", queryTempString)
                    dAdapSmpSearch = New OleDbDataAdapter(sql, connec)
                    connec.Open()
                    dAdapSmpSearch.Fill(dsSmpSearchExtraDateResultsEngrs, "Engineers")
                    connec.Close()

                    'copies data from the dataset into two temporary arrays
                    tempForename = (From myRow In dsSmpSearchExtraDateResultsEngrs.Tables(0).AsEnumerable Select myRow.Field(Of String)("Forename")).ToArray
                    tempSurname = (From myRow In dsSmpSearchExtraDateResultsEngrs.Tables(0).AsEnumerable Select myRow.Field(Of String)("Surname")).ToArray
                    'the tempoary arrays are then copied into the 
                    Dim localAmtsOnDateArray(tempForename.Length) As String
                    For counter = 0 To tempForename.Length - 1
                        localAmtsOnDateArray(counter) = tempForename(counter)
                        localAmtsOnDateArray(counter) = localAmtsOnDateArray(counter) & " " & tempSurname(counter)
                    Next

                    'joins all of the array elements
                    queryTempString = String.Join(" & ", localAmtsOnDateArray)

                    'gets the total number of log items for the specified day
                    For counter = 0 To localLogItemsArray.Length - 1
                        localTotalLogItems = localTotalLogItems + (localLogItemsArray(counter))
                    Next

                    'compiles the metrics string
                    dsaSmpSearchResString = dsaSmpSearchResString & "Metrics regarding the date of " & smpDSASearchDate & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "A total of " & dsSmpSearchExtraDateResults.Tables(0).Rows.Count & " AMTs worked on " & smpDSASearchDate & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "The AMTs present on " & smpDSASearchDate & " were: " & queryTempString & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "There were a total of " & localShipNoArray.Length & " flights on " & smpDSASearchDate & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "The total number of log items was: " & localTotalLogItems & vbCrLf

                Case Is = 5
                    'comments on a specific day
                    Dim queryTempString As String = String.Join(vbCrLf & "//->", localCommentsArray)
                    dsaSmpSearchResString = dsaSmpSearchResString & "The keyword " & smpDSASearchKeywords & " featured or was similar to other keywords " & localCommentsArray.Length & " times on " & smpDSASearchDate
                    dsaSmpSearchResString = dsaSmpSearchResString & "The following is a list of comments containing the keywords or similar " & queryTempString & vbCrLf
                Case Is = 6
                    'aircraft on a specific day and details
                    dsaSmpSearchResString = dsaSmpSearchResString & "Details for aircraft " & smpDSASearchShipNo & " on " & smpDSASearchDate & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "Time In " & localTimeInArray(0) & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "Time Out " & localTimeInArray(0) & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "Gate " & localGateArray(0) & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "Flight In " & localFlightInArray(0) & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "Flight Out " & localFlightOutArray(0) & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "AMTs " & localAmtsOnAircraftArray(0) & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & "Comments " & localCommentsArray(0) & vbCrLf

                Case Is = 7
                    'comments on specific aircraft on a specific day
                    dsaSmpSearchResString = dsaSmpSearchResString & "Comments for Aircraft " & smpDSASearchShipNo & " on " & smpDSASearchDate & vbCrLf
                    dsaSmpSearchResString = dsaSmpSearchResString & localCommentsArray(0)

            End Select

        End If

        'returns the string to the 
        Return dsaSmpSearchResString

    End Function

    Private Function advancedArchiveSearchMetrics(ByVal advDSASearchKeywords As String, ByVal advDSASearchShipNo As String, ByVal advDSASearchLoctn As String, ByVal advDSASearchFlightInNo As String, ByVal advDSASearchFlightOutNo As String, ByVal advDSASearchAMTs As String, ByVal advDSASearchDate As String) As String
        'return more advanced metrics from the daily sheet
        'private sub so can only be accessed by the control
        Dim dsaAdvSearchResString As String

        'database variables
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim dAdapSmpSearch As OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim dsAdvSearchResults As New DataSet
        Dim dtAdvSearchResultsTable As New DataTable
        Dim sql As String
        Dim queryBuilder As New StringBuilder
        Dim resultsBuilder As New StringBuilder

        'used to determine which parameters have been added
        Dim keywordsBool As Boolean
        Dim shipNoBool As Boolean
        Dim loctnBool As Boolean
        Dim flightInNoBool As Boolean
        Dim flightOutNoBool As Boolean
        Dim amtsBool As Boolean
        Dim dateSearchBool As Boolean

        'local arrays and variables to store data from the archive for processing
        Dim localCommentsArray() As String
        Dim localShipNoArray() As String
        Dim localDateArray() As String
        Dim localFlightInArray() As String
        Dim localFlightOutArray() As String
        Dim localTimeInArray() As String
        Dim localTimeOutArray() As String
        Dim localGateArray() As String
        Dim localLogItemsArray() As Integer
        Dim localTerminalsArray() As Integer
        Dim localTotalLogItems As Integer
        Dim localAmtsOnAircraftArray() As String
        Dim localSheetIDArray() As Integer

        'strings and running totals to put in the final return metrics string
        Dim allComments As String
        Dim allFlightsIn As String
        Dim allFlightsOut As String
        Dim allDates As String
        Dim allGates As String


        'builds a dynamic sql query
        'determines which columns to query based on input parameters
        'uses a string builder to add the expressions
        If advDSASearchKeywords <> Nothing Then
            queryBuilder.Append("[Entries].[Comments] LIKE %?%")
            keywordsBool = True
        End If
        If advDSASearchShipNo <> Nothing Then
            queryBuilder.Append("AND [Entries].[Ship] = ? ")
            shipNoBool = True
        End If
        If advDSASearchLoctn <> Nothing Then
            queryBuilder.Append("AND [Entries].[Terminal] = ? ")
            loctnBool = True
        End If
        If advDSASearchFlightInNo <> Nothing Then
            queryBuilder.Append("AND [Entries].[Flt_In] = ? ")
            flightInNoBool = True
        End If
        If advDSASearchFlightOutNo <> Nothing Then
            queryBuilder.Append("AND [Entries].[Flt_Out] = ? ")
            flightOutNoBool = True
        End If
        If advDSASearchAMTs <> Nothing Then
            queryBuilder.Append("AND [Entries].[AMTs] LIKE %?%")
            amtsBool = True
        End If
        If advDSASearchDate <> Nothing Then
            queryBuilder.Append("AND [Sheet].[Sheet_Date] = ?")
            dateSearchBool = True
        End If
        sql = "SELECT * FROM [EN/SHEETS], [Sheets], [Entries] WHERE "
        'appends the constructed text to the existing sql expression
        sql = sql & queryBuilder.ToString
        cmd = New OleDbCommand(sql, connec)
        'determines which parameters to add to the expression
        If keywordsBool = True Then
            cmd.Parameters.AddWithValue("@Comments", advDSASearchKeywords)
        End If
        If shipNoBool = True Then
            cmd.Parameters.AddWithValue("@Ship", advDSASearchShipNo)
        End If
        If loctnBool = True Then
            cmd.Parameters.AddWithValue("@Terminal", advDSASearchLoctn)
        End If
        If flightInNoBool = True Then
            cmd.Parameters.AddWithValue("@Flt_In", advDSASearchFlightInNo)
        End If
        If flightOutNoBool = True Then
            cmd.Parameters.AddWithValue("@Flt_Out", advDSASearchFlightOutNo)
        End If
        If amtsBool = True Then
            cmd.Parameters.AddWithValue("@AMTs", advDSASearchAMTs)
        End If
        If dateSearchBool = True Then
            cmd.Parameters.AddWithValue("@Sheet_Date", advDSASearchShipNo)
        End If

        'opens the database
        dAdapSmpSearch = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapSmpSearch.Fill(dsAdvSearchResults, "Entries")
        connec.Close()

        'check if any results were returned
        If dsAdvSearchResults.Tables(0).Rows.Count < 1 Then
            resultsBuilder.Append("No Results")
        Else
            resultsBuilder.Append("The following results were returned").AppendLine()
            '\\\\\\\\\\\\\\\\\\\\\
            'generate the metrics
            '/////////////////////

            'puts the results of the dataset into local arrays for processing
            localCommentsArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Commments")).ToArray
            localShipNoArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Ship")).ToArray
            localDateArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Sheet_Date")).ToArray
            localFlightInArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Flt_In")).ToArray
            localFlightOutArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Flt_Out")).ToArray
            localTimeInArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Time_In")).ToArray
            localTimeOutArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Time_Out")).ToArray
            localGateArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Gate")).ToArray
            localLogItemsArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Log_Items")).ToArray
            localAmtsOnAircraftArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("AMTs")).ToArray
            localTerminalsArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Terminal")).ToArray
            'used to query the link table between the engineers and the daily sheet table
            localSheetIDArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of Integer)("Sheet_ID")).ToArray

            'joins the array elements so that they can be put into the string
            allComments = String.Join(vbCrLf, localCommentsArray)
            allFlightsIn = String.Join(", ", localFlightInArray)
            allFlightsOut = String.Join(", ", localFlightOutArray)
            allDates = String.Join(" $$ ", localDateArray)
            allGates = String.Join(", ", localGateArray)

            'calculates the total number of log items for a specific aircraft
            'gets the total number of log items for the specified day
            For counter = 0 To localLogItemsArray.Length - 1
                localTotalLogItems = localTotalLogItems + (localLogItemsArray(counter))
            Next

            'returns a standard metrics response string, almost like a JSON string that has empty fields- not ideal but justify in documentation
            'metrics on a specific keyword
            resultsBuilder.Append("//// KEYWORD METRICS ////").AppendLine()
            resultsBuilder.Append("The keyword " & advDSASearchKeywords & " has " & localCommentsArray.Length & " occurrences.").AppendLine()
            resultsBuilder.Append("The keyword " & advDSASearchKeywords & " has occurrences on " & localSheetIDArray.Length & " Daily Sheets.").AppendLine()
            'metrics on a specific aircraft
            If shipNoBool = True Then
                resultsBuilder.Append("//// AIRCRAFT METRICS ////").AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " has " & localShipNoArray.Length & " daily sheet occurrences.").AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " has the comment " & advDSASearchKeywords & " " & localCommentsArray.Length & " times").AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " has supported all of the following inbound flights " & allFlightsIn).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " has supported all of the following outbound flights " & allFlightsOut).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " supported the inbound flight of " & advDSASearchFlightInNo & " on the following dates: " & allDates).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " supported the outbound flight of " & advDSASearchFlightOutNo & " on the following dates: " & allDates).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " has been at the following gates: " & allGates).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " was at gate " & localGateArray(0) & " on " & advDSASearchDate).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " was supported by the follwing AMTs: " & localAmtsOnAircraftArray(0) & " on " & advDSASearchDate).AppendLine()
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " arrived at " & localTimeInArray(0) & " and departed at " & localTimeOutArray(0) & " on " & advDSASearchDate).AppendLine()
            End If
            If dateSearchBool = False And shipNoBool = True Then
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " has had a total of " & localTotalLogItems & " log items").AppendLine()
            ElseIf dateSearchBool = True And shipNoBool = True Then
                resultsBuilder.Append("Aircraft " & advDSASearchShipNo & " had a total of " & localTotalLogItems & " log items on " & advDSASearchDate).AppendLine()
            End If
            'metrics on a specific engineer
            If amtsBool = True Then
                resultsBuilder.Append("//// ENGINEER METRICS ////").AppendLine()
                resultsBuilder.Append(advDSASearchAMTs & " features " & dsAdvSearchResults.Tables(0).Rows.Count & " times in this query.").AppendLine()
                resultsBuilder.Append(advDSASearchAMTs & " has been related to the comment of " & advDSASearchKeywords.Length & " times.").AppendLine()
                resultsBuilder.Append(advDSASearchAMTs & " has worked on aircraft " & advDSASearchShipNo & " a total of " & dsAdvSearchResults.Tables(0).Rows.Count & " times in this query.").AppendLine()
                resultsBuilder.Append("All daily sheet comments on all aircraft worked by " & advDSASearchAMTs & allComments)
            End If

        End If

        'have an error message appear if neither keywords, ship number or amts field has been used. 
        'the other fields are supplemental and cannot be used exclusively to search

        'assigns the result of the stringbuilder:
        dsaAdvSearchResString = resultsBuilder.ToString

        'returns the string
        Return dsaAdvSearchResString

    End Function

    Private Function advancedArchiveSearchActualDS(ByVal advDSASearchKeywords As String, ByVal advDSASearchShipNo As String, ByVal advDSASearchLoctn As String, ByVal advDSASearchFlightInNo As String, ByVal advDSASearchFlightOutNo As String, ByVal advDSASearchAMTs As String, ByVal advDSASearchDate As String) As String()
        'return the pdf of the daily sheet to be accessible in the viewer
        'private sub so can only be accessed by the control
        'return more advanced metrics from the daily sheet
        'private sub so can only be accessed by the control
        Dim dsaAdvSearchResArray() As String

        'database variables
        Dim DBconn As String = "provider=Microsoft.ACE.oledb.12.0;"
        Dim dbfilename As String = GetFolderPath(SpecialFolder.MyDocuments) & "\LEMSDataBase.accdb"
        Dim DBsource As String = "Data Source=" & dbfilename
        Dim connec As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(DBconn & DBsource)
        Dim dAdapSmpSearch As OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim dsAdvSearchResults As New DataSet
        Dim dtAdvSearchResultsTable As New DataTable
        Dim sql As String
        Dim queryBuilder As New StringBuilder
        Dim resultsBuilder As New StringBuilder

        'used to determine which parameters have been added
        Dim keywordsBool As Boolean
        Dim shipNoBool As Boolean
        Dim loctnBool As Boolean
        Dim flightInNoBool As Boolean
        Dim flightOutNoBool As Boolean
        Dim amtsBool As Boolean
        Dim dateSearchBool As Boolean

        'local array to store all of the filenames
        Dim localFileNameArray() As String
        Dim localReturnNullArray(1) As String

        'builds a dynamic sql query
        'determines which columns to query based on input parameters
        'uses a string builder to add the expressions
        If advDSASearchKeywords <> Nothing Then
            queryBuilder.Append("[Entries].[Comments] LIKE %?%")
            keywordsBool = True
        End If
        If advDSASearchShipNo <> Nothing Then
            queryBuilder.Append("AND [Entries].[Ship] = ? ")
            shipNoBool = True
        End If
        If advDSASearchLoctn <> Nothing Then
            queryBuilder.Append("AND [Entries].[Terminal] = ? ")
            loctnBool = True
        End If
        If advDSASearchFlightInNo <> Nothing Then
            queryBuilder.Append("AND [Entries].[Flt_In] = ? ")
            flightInNoBool = True
        End If
        If advDSASearchFlightOutNo <> Nothing Then
            queryBuilder.Append("AND [Entries].[Flt_Out] = ? ")
            flightOutNoBool = True
        End If
        If advDSASearchAMTs <> Nothing Then
            queryBuilder.Append("AND [Entries].[AMTs] LIKE %?%")
            amtsBool = True
        End If
        If advDSASearchDate <> Nothing Then
            queryBuilder.Append("AND [Sheet].[Sheet_Date] = ?")
            dateSearchBool = True
        End If
        sql = "SELECT * FROM [Sheets], [Entries] WHERE "
        'appends the constructed text to the existing sql expression
        sql = sql & queryBuilder.ToString
        cmd = New OleDbCommand(sql, connec)
        'determines which parameters to add to the expression
        If keywordsBool = True Then
            cmd.Parameters.AddWithValue("@Comments", advDSASearchKeywords)
        End If
        If shipNoBool = True Then
            cmd.Parameters.AddWithValue("@Ship", advDSASearchShipNo)
        End If
        If loctnBool = True Then
            cmd.Parameters.AddWithValue("@Terminal", advDSASearchLoctn)
        End If
        If flightInNoBool = True Then
            cmd.Parameters.AddWithValue("@Flt_In", advDSASearchFlightInNo)
        End If
        If flightOutNoBool = True Then
            cmd.Parameters.AddWithValue("@Flt_Out", advDSASearchFlightOutNo)
        End If
        If amtsBool = True Then
            cmd.Parameters.AddWithValue("@AMTs", advDSASearchAMTs)
        End If
        If dateSearchBool = True Then
            cmd.Parameters.AddWithValue("@Sheet_Date", advDSASearchShipNo)
        End If

        'opens the database
        dAdapSmpSearch = New OleDbDataAdapter(sql, connec)
        connec.Open()
        dAdapSmpSearch.Fill(dsAdvSearchResults, "Entries")
        connec.Close()

        'determines if there are any results to return
        If dsAdvSearchResults.Tables(0).Rows.Count < 1 Then
            'if there is no data, return an array with a null result message in the first element
            localReturnNullArray(0) = "No Results"
            dsaAdvSearchResArray = localReturnNullArray
            Return dsaAdvSearchResArray
        Else
            'if there is data
            'assigns the contents of the filename column to the return array
            localFileNameArray = (From myRow In dsAdvSearchResults.Tables(0).AsEnumerable Select myRow.Field(Of String)("Sheet_Filename")).ToArray
            dsaAdvSearchResArray = localFileNameArray
            Return dsaAdvSearchResArray
        End If

    End Function

End Module
