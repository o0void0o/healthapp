<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/app.asp" -->
<%
Dim connection, values, errors, id, idText, existing, key, isPost, formToken, contextValue
isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")
If isPost Then
    RequirePost
    idText = Trim(Text(Request.Form("id")))
Else
    RequireGet
    idText = Trim(Text(Request.QueryString("id")))
End If
id = 0
If idText <> "" And idText <> "0" Then
    If Not ParseInteger(idText, 1, 2147483647, id) Then HttpError "400 Bad Request", "Invalid reading", "Choose an existing reading from history."
End If
Set values = NewReading()
Set errors = CreateObject("Scripting.Dictionary")
Set connection = OpenDatabase()
Set existing = Nothing
If id > 0 Then
    Set existing = LoadReading(connection, id)
    If existing Is Nothing Then HttpError "404 Not Found", "Reading not found", "It may have been deleted. Return to history to see your current readings."
End If

If isPost Then
    For Each key In Array("type", "recorded_at", "systolic", "diastolic", "pulse", "glucose_value", "glucose_unit", "context")
        values(key) = Trim(Text(Request.Form(key)))
    Next
    values("notes") = Text(Request.Form("notes"))
    Set errors = ValidateReading(values, Now())
    If id > 0 Then
        If values("type") <> existing("ReadingType") Then errors("type") = "The type of an existing reading cannot be changed."
    End If
    If errors.Count = 0 Then
        SaveReading connection, values, id
        connection.Close
        If id = 0 Then Session("flash") = "Reading added. Your journal is up to date."
        If id > 0 Then Session("flash") = "Reading updated."
        Response.Status = "303 See Other"
        Response.AddHeader "Location", "history.asp"
        Response.End
    End If
    Response.Status = "422 Unprocessable Entity"
ElseIf id > 0 Then
    values("type") = existing("ReadingType")
    values("recorded_at") = IsoTimestamp(existing("RecordedAt"))
    values("systolic") = Text(existing("Systolic"))
    values("diastolic") = Text(existing("Diastolic"))
    values("pulse") = Text(existing("Pulse"))
    If Not IsNull(existing("GlucoseValue")) Then values("glucose_value") = CompactNumber(existing("GlucoseValue"))
    values("glucose_unit") = Text(existing("GlucoseUnit"))
    values("context") = Text(existing("MealContext"))
    values("notes") = Text(existing("Notes"))
Else
    values("type") = Text(Request.QueryString("type"))
    If values("type") = "" Then values("type") = "pressure"
    If values("type") <> "pressure" And values("type") <> "glucose" Then HttpError "400 Bad Request", "Invalid reading type", "Choose blood pressure or blood glucose."
    values("recorded_at") = IsoTimestamp(Now())
    values("glucose_unit") = "mmol"
End If
connection.Close
Set connection = Nothing
formToken = CsrfToken()
pageTitle = "Add a reading"
If id > 0 Then pageTitle = "Edit reading"
activePage = "add"

Function InvalidAttribute(field)
    InvalidAttribute = ""
    If errors.Exists(field) Then InvalidAttribute = " aria-invalid=""true"""
End Function
%>
<!--#include file="includes/header.asp" -->
<section class="page-heading"><div><p class="eyebrow">YOUR NEXT ENTRY</p><h1><%= H(pageTitle) %></h1><p class="lead">A few details now. A more complete picture over time.</p></div><a class="text-link" href="history.asp">&larr; Back to history</a></section>
<div class="entry-layout">
<section class="panel entry-panel" aria-label="Reading entry">
    <% If errors.Count > 0 Then %>
    <div class="notice notice-error error-summary" role="alert" tabindex="-1" data-error-summary>
        <h2>Please check your reading</h2><ul><% For Each key In errors.Keys %><li><a href="#<%= H(key) %>"><%= H(errors(key)) %></a></li><% Next %></ul>
    </div>
    <% End If %>
    <% If id = 0 Then %>
    <nav class="entry-tabs" aria-label="Reading type" id="type">
        <a href="reading.asp?type=pressure" class="entry-tab<% If values("type") = "pressure" Then Response.Write " selected-tab" %>"<% If values("type") = "pressure" Then Response.Write " aria-current=""page""" %>>Blood pressure</a>
        <a href="reading.asp?type=glucose" class="entry-tab<% If values("type") = "glucose" Then Response.Write " selected-tab" %>"<% If values("type") = "glucose" Then Response.Write " aria-current=""page""" %>>Blood glucose</a>
    </nav>
    <% Else %>
    <p class="edit-type" id="type">Editing <% If values("type") = "pressure" Then %>blood pressure<% Else %>blood glucose<% End If %> &middot; Entry #<%= H(id) %></p>
    <% End If %>
    <form action="reading.asp" method="post" class="reading-form" data-saving-form>
        <input type="hidden" name="csrf" value="<%= H(formToken) %>">
        <input type="hidden" name="id" value="<%= H(id) %>">
        <input type="hidden" name="type" value="<%= H(values("type")) %>">
        <div class="field"><label for="recorded_at">Date &amp; time <span class="required-label">Required</span></label><input id="recorded_at" name="recorded_at" type="datetime-local" value="<%= H(values("recorded_at")) %>" min="1900-01-01T00:00" required aria-describedby="time-help"<%= InvalidAttribute("recorded_at") %>><p class="field-help" id="time-help">The local date and time the measurement was taken.</p></div>
        <% If values("type") = "pressure" Then %>
        <fieldset><legend>Blood pressure <span class="legend-unit">mmHg</span></legend>
            <div class="two-fields">
                <div class="field"><label for="systolic">Systolic <span class="required-label">Required</span></label><input id="systolic" name="systolic" type="number" inputmode="numeric" min="1" max="400" step="1" value="<%= H(values("systolic")) %>" placeholder="e.g. 120" required<%= InvalidAttribute("systolic") %>><p class="field-help">The upper number.</p></div>
                <div class="field"><label for="diastolic">Diastolic <span class="required-label">Required</span></label><input id="diastolic" name="diastolic" type="number" inputmode="numeric" min="1" max="300" step="1" value="<%= H(values("diastolic")) %>" placeholder="e.g. 80" required<%= InvalidAttribute("diastolic") %>><p class="field-help">The lower number.</p></div>
            </div>
        </fieldset>
        <div class="field half-field"><label for="pulse">Pulse <span class="optional-label">Optional, bpm</span></label><input id="pulse" name="pulse" type="number" inputmode="numeric" min="1" max="300" step="1" value="<%= H(values("pulse")) %>" placeholder="e.g. 72"<%= InvalidAttribute("pulse") %>></div>
        <% Else %>
        <fieldset><legend>Blood glucose</legend>
            <div class="two-fields">
                <div class="field"><label for="glucose_value">Value <span class="required-label">Required</span></label><input id="glucose_value" name="glucose_value" type="number" inputmode="decimal" min="0.01" max="2000" step="0.01" value="<%= H(values("glucose_value")) %>" placeholder="e.g. 5.6" required<%= InvalidAttribute("glucose_value") %>><p class="field-help">Use the value and units shown on your meter.</p></div>
                <div class="field"><label for="glucose_unit">Units <span class="required-label">Required</span></label><select id="glucose_unit" name="glucose_unit" required aria-describedby="unit-help"<%= InvalidAttribute("glucose_unit") %>><option value="mmol"<%= Selected(values("glucose_unit"), "mmol") %>>mmol/L</option><option value="mg"<%= Selected(values("glucose_unit"), "mg") %>>mg/dL</option></select><p class="field-help" id="unit-help">Changing units does not convert the entered number.</p></div>
            </div>
        </fieldset>
        <div class="field"><label for="context">Measurement context <span class="optional-label">Optional</span></label><select id="context" name="context"<%= InvalidAttribute("context") %>><option value="">Not specified</option><% For Each contextValue In Array("fasting", "before-meal", "after-meal", "bedtime", "random") %><option value="<%= H(contextValue) %>"<%= Selected(values("context"), contextValue) %>><%= H(ContextLabel(contextValue)) %></option><% Next %></select></div>
        <% End If %>
        <div class="field"><label for="notes">Notes <span class="optional-label">Optional</span></label><textarea id="notes" name="notes" rows="4" maxlength="2000" placeholder="Anything you would like to remember about this reading..." aria-describedby="notes-help"<%= InvalidAttribute("notes") %>><%= H(values("notes")) %></textarea><p class="field-help" id="notes-help">Up to 2,000 characters. Avoid information you do not want stored on this computer.</p></div>
        <div class="form-actions"><button class="button" type="submit"><% If id > 0 Then %>Save changes<% Else %>Save reading<% End If %></button><a class="button button-quiet" href="history.asp">Cancel</a></div>
    </form>
</section>
<aside class="entry-aside">
    <section class="panel guidance-card"><span class="guidance-icon" aria-hidden="true">i</span><h2>A useful record starts with context.</h2><ul><li>Follow your measuring device's instructions.</li><li>Check the time and, for glucose, the units before saving.</li><li>Add a note if there is something you want to discuss with your healthcare professional.</li></ul></section>
    <p class="aside-note">This journal stores observations. It does not interpret readings, diagnose conditions, or recommend treatment. Input limits are typing checks, not clinical targets.</p>
</aside>
</div>
<!--#include file="includes/footer.asp" -->
