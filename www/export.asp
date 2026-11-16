<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/app.asp" -->
<%
RequireGet
Dim filters, connection, command, records, cells, index, line
Set filters = ReadFilters()
Set connection = OpenDatabase()
Set command = FilteredCommand(connection, "SELECT " & READ_COLUMNS, filters)
command.CommandText = command.CommandText & " ORDER BY [RecordedAt] DESC, [Id] DESC"
Set records = RunReader(command)
Response.Clear
Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "attachment; filename=""health-readings-" & IsoDate(Date()) & ".csv"""
Response.Buffer = False
Response.Write ChrW(&HFEFF)
Response.Write "Id,Type,RecordedAtLocal,Systolic_mmHg,Diastolic_mmHg,Pulse_bpm,GlucoseValue,GlucoseUnit,Context,Notes" & vbCrLf
Do Until records.EOF
    cells = Array(records("Id").Value, records("ReadingType").Value, IsoTimestamp(records("RecordedAt")), records("Systolic").Value, records("Diastolic").Value, records("Pulse").Value, "", "", ContextLabel(records("MealContext")), records("Notes").Value)
    If records("ReadingType") = "glucose" Then
        cells(6) = CompactNumber(records("GlucoseValue"))
        cells(7) = UnitLabel(records("GlucoseUnit"))
    Else
        cells(8) = ""
    End If
    line = ""
    For index = 0 To UBound(cells)
        If index > 0 Then line = line & ","
        line = line & CsvCell(cells(index))
    Next
    Response.Write line & vbCrLf
    records.MoveNext
Loop
records.Close
connection.Close
%>
