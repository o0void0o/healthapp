<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/app.asp" -->
<%
RequireGet
Dim filters, connection, command, records, total, page, pageCount, pageText, offset, skipIndex
Dim tableRows, showActions, firstItem, lastItem, query
Set filters = ReadFilters()
page = 1
pageText = Text(Request.QueryString("page"))
If pageText <> "" Then
    If Not ParseInteger(pageText, 1, 1000000, page) Then HttpError "400 Bad Request", "Invalid page", "Choose a page using the history navigation."
End If
Set connection = OpenDatabase()
Set command = FilteredCommand(connection, "SELECT COUNT(*) AS [Total]", filters)
Set records = RunReader(command)
total = records("Total").Value
records.Close
pageCount = Int((total + PAGE_SIZE - 1) / PAGE_SIZE)
If pageCount = 0 Then pageCount = 1
If page > pageCount Then HttpError "400 Bad Request", "Page not found", "The requested page is outside the filtered history. Return to the first page."
Set command = FilteredCommand(connection, "SELECT " & READ_COLUMNS, filters)
command.CommandText = command.CommandText & " ORDER BY [RecordedAt] DESC, [Id] DESC"
Set records = RunReader(command)
offset = (page - 1) * PAGE_SIZE
For skipIndex = 1 To offset
    If Not records.EOF Then records.MoveNext
Next
tableRows = Empty
If Not records.EOF Then tableRows = records.GetRows(PAGE_SIZE)
records.Close
connection.Close
Set connection = Nothing
firstItem = 0
lastItem = 0
If total > 0 Then
    firstItem = offset + 1
    lastItem = offset + UBound(tableRows, 2) + 1
End If
query = FilterUrl(filters)
showActions = True
pageTitle = "Reading history"
activePage = "history"
%>
<!--#include file="includes/header.asp" -->
<section class="page-heading"><div><p class="eyebrow">YOUR RECORDS</p><h1>Reading history</h1><p class="lead">Every entry, in one place. Filter, review, or take a copy with you.</p></div><a class="button button-secondary" href="export.asp?<%= H(query) %>">Download CSV</a></section>
<section class="panel filter-panel" aria-label="Filter readings">
    <form action="history.asp" method="get" class="filter-form">
        <div class="field"><label for="type">Reading type</label><select id="type" name="type"><option value="">All readings</option><option value="pressure"<%= Selected(filters("type"), "pressure") %>>Blood pressure</option><option value="glucose"<%= Selected(filters("type"), "glucose") %>>Blood glucose</option></select></div>
        <div class="field"><label for="from">From date</label><input id="from" name="from" type="date" min="1900-01-01" value="<%= H(filters("from")) %>"></div>
        <div class="field"><label for="to">To date <span class="optional-label">Inclusive</span></label><input id="to" name="to" type="date" min="1900-01-01" value="<%= H(filters("to")) %>"></div>
        <div class="filter-actions"><button class="button" type="submit">Apply filters</button><a class="button button-quiet" href="history.asp">Reset</a></div>
    </form>
</section>
<section class="panel history-panel" aria-label="Saved readings">
    <div class="section-heading panel-heading"><div><h2>Saved readings <span class="count-badge"><%= H(total) %></span></h2><p class="muted">Newest measurements first. Glucose is shown in its original units.</p></div><span class="muted result-count">Showing <%= H(firstItem) %>&ndash;<%= H(lastItem) %> of <%= H(total) %></span></div>
    <!--#include file="includes/reading-table.asp" -->
    <% If pageCount > 1 Then %>
    <nav class="pagination" aria-label="History pages">
        <% If page > 1 Then %><a class="button button-secondary button-small" href="history.asp?<%= H(query) %>&amp;page=<%= H(page - 1) %>">&larr; Previous</a><% Else %><span></span><% End If %>
        <span>Page <%= H(page) %> of <%= H(pageCount) %></span>
        <% If page < pageCount Then %><a class="button button-secondary button-small" href="history.asp?<%= H(query) %>&amp;page=<%= H(page + 1) %>">Next &rarr;</a><% Else %><span></span><% End If %>
    </nav>
    <% End If %>
</section>
<p class="history-note">CSV downloads include all matching readings, not just this page. Notes can be viewed in full when editing a reading.</p>
<!--#include file="includes/footer.asp" -->
