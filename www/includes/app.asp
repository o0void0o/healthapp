<% Option Explicit %>
<!--#include file="validation.asp" -->
<%
Const READ_COLUMNS = "[Id], [ReadingType], [RecordedAt], [Systolic], [Diastolic], [Pulse], [GlucoseValue], [GlucoseUnit], [MealContext], [Notes]"
Const PAGE_SIZE = 20
Dim pageTitle, activePage

Sub HttpError(status, title, message)
    Response.Clear
    Response.Status = status
    Response.ContentType = "text/html"
    Response.Write "<!doctype html><html lang=""en""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width, initial-scale=1""><title>" & H(title) & " - Health Journal</title><link rel=""stylesheet"" href=""assets/app.css""></head>"
    Response.Write "<body class=""error-page""><main class=""panel error-card""><p class=""eyebrow"">HEALTH JOURNAL</p><h1>" & H(title) & "</h1><p>" & H(message) & "</p><a class=""button"" href=""default.asp"">Back to dashboard</a></main></body></html>"
    Response.End
End Sub

Sub ServerError(operation, errorNumber, detail)
    Dim fileSystem, log, path, loggingError
    Response.AppendToLog " HealthJournal-" & Replace(operation, " ", "-") & "-" & CStr(errorNumber)
    Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
    path = fileSystem.BuildPath(fileSystem.GetParentFolderName(Server.MapPath(".")), "data\errors.log")
    On Error Resume Next
    Set log = fileSystem.OpenTextFile(path, 8, True, -1)
    log.WriteLine IsoTimestamp(Now()) & " | " & operation & " | " & CStr(errorNumber) & " | " & Replace(Replace(Text(detail), vbCr, " "), vbLf, " ")
    log.Close
    loggingError = Err.Number
    On Error GoTo 0
    If loggingError <> 0 Then Response.AppendToLog " HealthJournal-logging-failed-" & CStr(loggingError)
    HttpError "500 Internal Server Error", "We couldn't complete that action", "An error occurred during " & operation & ". No success has been reported. Check that the Access database exists and the data folder is writable. See README.md and the private IIS logs for troubleshooting."
End Sub

Sub InitializeApp()
    Dim address, host, port, allowedHost
    Response.Buffer = True
    Response.Charset = "utf-8"
    Response.CodePage = 65001
    Session.CodePage = 65001
    Session.LCID = 1033
    Session.Timeout = 30
    Response.Expires = -1
    Response.AddHeader "Cache-Control", "no-store, private"
    address = LCase(Request.ServerVariables("REMOTE_ADDR"))
    If address <> "127.0.0.1" And address <> "::1" And address <> "::ffff:127.0.0.1" Then
        HttpError "403 Forbidden", "Local access only", "This personal journal is available only on the computer running it."
    End If
    host = LCase(Request.ServerVariables("HTTP_HOST"))
    port = Request.ServerVariables("SERVER_PORT")
    allowedHost = "localhost:" & port
    If host <> allowedHost Then
        If Not (host = "localhost" And (port = "80" Or port = "443")) Then
            HttpError "403 Forbidden", "Use localhost", "Open this journal using the localhost address shown by Start-HealthApp.ps1."
        End If
    End If
End Sub

Sub RequireGet()
    If Request.ServerVariables("REQUEST_METHOD") <> "GET" And Request.ServerVariables("REQUEST_METHOD") <> "HEAD" Then
        Response.AddHeader "Allow", "GET, HEAD"
        HttpError "405 Method Not Allowed", "Method not allowed", "Use a normal page request for this action."
    End If
End Sub

Function CsrfToken()
    Dim generator, token, errorNumber, detail
    If Len(Text(Session("csrf"))) = 0 Then
        On Error Resume Next
        Set generator = Server.CreateObject("Scriptlet.TypeLib")
        ' The COM strings include trailing NULs after the 38-character UUID.
        token = Left(generator.GUID, 38) & Left(generator.GUID, 38)
        errorNumber = Err.Number
        detail = Err.Description
        On Error GoTo 0
        If errorNumber <> 0 Then ServerError "secure token generation", errorNumber, detail
        token = Replace(Replace(Replace(token, "{", ""), "}", ""), "-", "")
        If Not Matches(token, "^[0-9A-Fa-f]{64}$") Then ServerError "secure token generation", vbObjectError + 103, "Windows returned an invalid GUID token."
        Session("csrf") = token
        Set generator = Nothing
    End If
    CsrfToken = Session("csrf")
End Function

Sub RequirePost()
    Dim origin, expected
    If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
        Response.AddHeader "Allow", "POST"
        HttpError "405 Method Not Allowed", "Method not allowed", "This action requires the journal's form."
    End If
    origin = LCase(Request.ServerVariables("HTTP_ORIGIN"))
    expected = "http://" & LCase(Request.ServerVariables("HTTP_HOST"))
    If LCase(Request.ServerVariables("HTTPS")) = "on" Then expected = "https://" & LCase(Request.ServerVariables("HTTP_HOST"))
    If (origin <> "" And origin <> expected) Or LCase(Request.ServerVariables("HTTP_SEC_FETCH_SITE")) = "cross-site" Then
        HttpError "403 Forbidden", "Request blocked", "Submit this action from the journal on this computer."
    End If
    If Len(Text(Session("csrf"))) = 0 Or Text(Request.Form("csrf")) <> Text(Session("csrf")) Then
        HttpError "403 Forbidden", "Your form has expired", "Reload the form and try again. Cookies must be enabled."
    End If
End Sub

Function OpenDatabase()
    Dim connection, fileSystem, databasePath, errorNumber, detail
    Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
    databasePath = fileSystem.BuildPath(fileSystem.GetParentFolderName(Server.MapPath(".")), "data\health.mdb")
    Set connection = Server.CreateObject("ADODB.Connection")
    connection.ConnectionTimeout = 5
    connection.CommandTimeout = 15
    On Error Resume Next
    connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & databasePath & ";"
    errorNumber = Err.Number
    detail = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 Then ServerError "opening the database", errorNumber, detail
    Set OpenDatabase = connection
End Function

Function NewCommand(connection, sql)
    Dim command
    Set command = Server.CreateObject("ADODB.Command")
    Set command.ActiveConnection = connection
    command.CommandType = 1
    command.CommandTimeout = 15
    command.CommandText = sql
    Set NewCommand = command
End Function

Sub AddParameter(command, dataType, value, size)
    Dim parameter
    ' Reading Parameters.Count here makes Jet auto-derive extra, unbound parameters.
    If size > 0 Then
        Set parameter = command.CreateParameter("", dataType, 1, size, value)
    Else
        Set parameter = command.CreateParameter("", dataType, 1, , value)
    End If
    command.Parameters.Append parameter
End Sub

Function RunReader(command)
    Dim records, errorNumber, detail
    On Error Resume Next
    Set records = command.Execute
    errorNumber = Err.Number
    detail = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 Then ServerError "reading the database", errorNumber, detail
    Set RunReader = records
End Function

Sub RunChange(command)
    Dim affected, errorNumber, detail
    On Error Resume Next
    command.Execute affected, , 128
    errorNumber = Err.Number
    detail = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 Then ServerError "saving the database", errorNumber, detail
    If affected <> 1 Then HttpError "409 Conflict", "Reading not changed", "The reading may have been removed in another tab. Return to history and try again."
End Sub

Function ReadRows(command, maximum)
    Dim records, rows
    rows = Empty
    Set records = RunReader(command)
    If Not records.EOF Then
        If maximum > 0 Then
            rows = records.GetRows(maximum)
        Else
            rows = records.GetRows()
        End If
    End If
    records.Close
    Set records = Nothing
    ReadRows = rows
End Function

Function LoadReading(connection, id)
    Dim command, records, result, field
    Set command = NewCommand(connection, "SELECT * FROM [Readings] WHERE [Id] = ?")
    AddParameter command, 3, id, 0
    Set records = RunReader(command)
    Set result = Nothing
    If Not records.EOF Then
        Set result = CreateObject("Scripting.Dictionary")
        For Each field In records.Fields
            result.Add field.Name, field.Value
        Next
    End If
    records.Close
    Set LoadReading = result
End Function

Function LatestReading(connection, kind)
    Dim command, records, result, field
    Set command = NewCommand(connection, "SELECT TOP 1 * FROM [Readings] WHERE [ReadingType] = ? ORDER BY [RecordedAt] DESC, [Id] DESC")
    AddParameter command, 202, kind, 8
    Set records = RunReader(command)
    Set result = Nothing
    If Not records.EOF Then
        Set result = CreateObject("Scripting.Dictionary")
        For Each field In records.Fields
            result.Add field.Name, field.Value
        Next
    End If
    records.Close
    Set LatestReading = result
End Function

Sub SaveReading(connection, values, id)
    Dim command, unit, context, notes, timestamp, sql
    unit = Null
    context = Null
    notes = Null
    If values("notes") <> "" Then notes = values("notes")
    If values("type") = "glucose" Then
        unit = values("glucose_unit")
        If values("context") <> "" Then context = values("context")
    End If
    timestamp = Now()
    If id = 0 Then
        sql = "INSERT INTO [Readings] ([ReadingType], [RecordedAt], [Systolic], [Diastolic], [Pulse], [GlucoseValue], [GlucoseUnit], [MealContext], [Notes], [CreatedAt], [UpdatedAt]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    Else
        sql = "UPDATE [Readings] SET [ReadingType] = ?, [RecordedAt] = ?, [Systolic] = ?, [Diastolic] = ?, [Pulse] = ?, [GlucoseValue] = ?, [GlucoseUnit] = ?, [MealContext] = ?, [Notes] = ?, [UpdatedAt] = ? WHERE [Id] = ?"
    End If
    Set command = NewCommand(connection, sql)
    AddParameter command, 202, values("type"), 8
    AddParameter command, 7, values("parsed_at"), 0
    AddParameter command, 2, values("parsed_systolic"), 0
    AddParameter command, 2, values("parsed_diastolic"), 0
    AddParameter command, 2, values("parsed_pulse"), 0
    AddParameter command, 5, values("parsed_glucose"), 0
    AddParameter command, 202, unit, 4
    AddParameter command, 202, context, 24
    AddParameter command, 203, notes, 2000
    If id = 0 Then AddParameter command, 7, timestamp, 0
    AddParameter command, 7, timestamp, 0
    If id <> 0 Then AddParameter command, 3, id, 0
    RunChange command
End Sub

Function ReadFilters()
    Dim filters, key, parsed
    Set filters = CreateObject("Scripting.Dictionary")
    For Each key In Array("type", "from", "to")
        filters.Add key, Trim(Text(Request.QueryString(key)))
    Next
    filters.Add "fromDate", Null
    filters.Add "toDate", Null
    If filters("type") <> "" And filters("type") <> "pressure" And filters("type") <> "glucose" Then
        HttpError "400 Bad Request", "Invalid reading type", "Choose all readings, blood pressure, or blood glucose."
    End If
    If filters("from") <> "" Then
        If Not ParseDateOnly(filters("from"), parsed) Then HttpError "400 Bad Request", "Invalid start date", "Use a valid date in YYYY-MM-DD format."
        filters("fromDate") = parsed
    End If
    If filters("to") <> "" Then
        If Not ParseDateOnly(filters("to"), parsed) Then HttpError "400 Bad Request", "Invalid end date", "Use a valid date in YYYY-MM-DD format."
        filters("toDate") = parsed
    End If
    If Not IsNull(filters("fromDate")) And Not IsNull(filters("toDate")) Then
        If filters("fromDate") > filters("toDate") Then HttpError "400 Bad Request", "Check the date range", "The start date must be on or before the end date."
    End If
    Set ReadFilters = filters
End Function

Function FilteredCommand(connection, projection, filters)
    Dim command
    Set command = NewCommand(connection, projection & " FROM [Readings] WHERE 1 = 1")
    If filters("type") <> "" Then
        command.CommandText = command.CommandText & " AND [ReadingType] = ?"
        AddParameter command, 202, filters("type"), 8
    End If
    If Not IsNull(filters("fromDate")) Then
        command.CommandText = command.CommandText & " AND [RecordedAt] >= ?"
        AddParameter command, 7, filters("fromDate"), 0
    End If
    If Not IsNull(filters("toDate")) Then
        command.CommandText = command.CommandText & " AND [RecordedAt] < ?"
        AddParameter command, 7, DateAdd("d", 1, filters("toDate")), 0
    End If
    Set FilteredCommand = command
End Function

Function FilterUrl(filters)
    FilterUrl = "type=" & Server.URLEncode(filters("type")) & "&from=" & Server.URLEncode(filters("from")) & "&to=" & Server.URLEncode(filters("to"))
End Function

Function TrendJson(connection, kind, startDate, targetUnit)
    Dim command, records, result, point
    Set command = NewCommand(connection, "SELECT TOP 180 [RecordedAt], [Systolic], [Diastolic], [GlucoseValue], [GlucoseUnit] FROM [Readings] WHERE [ReadingType] = ? AND [RecordedAt] >= ? ORDER BY [RecordedAt] DESC, [Id] DESC")
    AddParameter command, 202, kind, 8
    AddParameter command, 7, startDate, 0
    Set records = RunReader(command)
    result = ""
    Do Until records.EOF
        point = "{""date"":""" & IsoTimestamp(records("RecordedAt")) & """"
        If kind = "pressure" Then
            point = point & ",""systolic"":" & JsonNumber(records("Systolic")) & ",""diastolic"":" & JsonNumber(records("Diastolic"))
        Else
            point = point & ",""value"":" & JsonNumber(ConvertGlucose(records("GlucoseValue"), records("GlucoseUnit"), targetUnit))
        End If
        If result <> "" Then result = result & ","
        result = result & point & "}"
        records.MoveNext
    Loop
    records.Close
    TrendJson = "[" & result & "]"
End Function

Function Selected(value, expected)
    Selected = ""
    If value = expected Then Selected = " selected"
End Function

Function ReadingSummary(reading)
    If reading("ReadingType") = "pressure" Then
        ReadingSummary = reading("Systolic") & " / " & reading("Diastolic") & " mmHg"
    Else
        ReadingSummary = CompactNumber(reading("GlucoseValue")) & " " & UnitLabel(reading("GlucoseUnit"))
    End If
End Function

InitializeApp
%>
