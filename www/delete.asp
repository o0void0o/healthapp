<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/app.asp" -->
<%
Dim id, idText, connection, reading, command, isPost, formToken
isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")
If isPost Then
    RequirePost
    idText = Text(Request.Form("id"))
Else
    RequireGet
    idText = Text(Request.QueryString("id"))
End If
If Not ParseInteger(idText, 1, 2147483647, id) Then HttpError "400 Bad Request", "Invalid reading", "Choose an existing reading from history."
Set connection = OpenDatabase()
Set reading = LoadReading(connection, id)
If reading Is Nothing Then HttpError "404 Not Found", "Reading not found", "It may already have been deleted."
If isPost Then
    Set command = NewCommand(connection, "DELETE FROM [Readings] WHERE [Id] = ?")
    AddParameter command, 3, id, 0
    RunChange command
    connection.Close
    Session("flash") = "Reading deleted."
    Response.Status = "303 See Other"
    Response.AddHeader "Location", "history.asp"
    Response.End
End If
connection.Close
formToken = CsrfToken()
pageTitle = "Delete reading"
activePage = "history"
%>
<!--#include file="includes/header.asp" -->
<section class="panel delete-card">
    <p class="eyebrow">CONFIRM DELETION</p><h1>Delete this reading?</h1>
    <p>This permanently removes the entry from your journal. This action cannot be undone.</p>
    <div class="delete-summary"><strong><%= H(ReadingSummary(reading)) %></strong><span><%= H(DisplayDate(reading("RecordedAt"))) %> at <%= H(DisplayTime(reading("RecordedAt"))) %></span><% If Len(Text(reading("Notes"))) > 0 Then %><p class="preserve-lines"><%= H(reading("Notes")) %></p><% End If %></div>
    <form action="delete.asp" method="post" class="form-actions" data-saving-form>
        <input type="hidden" name="csrf" value="<%= H(formToken) %>"><input type="hidden" name="id" value="<%= H(id) %>">
        <button class="button button-danger" type="submit">Delete reading</button><a class="button button-quiet" href="history.asp">Keep reading</a>
    </form>
</section>
<!--#include file="includes/footer.asp" -->
