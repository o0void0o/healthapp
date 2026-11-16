<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/app.asp" -->
<%
RequireGet
Dim period, unit, startDate, periodLabel, connection, pressure, glucose
Dim command, records, readingCount, pressureJson, glucoseJson, tableRows, showActions
period = Text(Request.QueryString("range"))
If period = "" Then period = "30"
Select Case period
    Case "7", "30", "90"
        startDate = DateAdd("d", 1 - CInt(period), Date())
        periodLabel = "Last " & period & " days"
    Case "all"
        startDate = DateSerial(1900, 1, 1)
        periodLabel = "All time"
    Case Else
        HttpError "400 Bad Request", "Invalid period", "Choose 7, 30, or 90 days, or all time."
End Select
unit = Text(Request.QueryString("unit"))
If unit = "" Then unit = "mmol"
If unit <> "mmol" And unit <> "mg" Then HttpError "400 Bad Request", "Invalid glucose unit", "Choose mmol/L or mg/dL."

Set connection = OpenDatabase()
Set pressure = LatestReading(connection, "pressure")
Set glucose = LatestReading(connection, "glucose")
Set command = NewCommand(connection, "SELECT COUNT(*) AS [Total] FROM [Readings] WHERE [RecordedAt] >= ?")
AddParameter command, 7, startDate, 0
Set records = RunReader(command)
readingCount = records("Total").Value
records.Close
pressureJson = TrendJson(connection, "pressure", startDate, unit)
glucoseJson = TrendJson(connection, "glucose", startDate, unit)
Set command = NewCommand(connection, "SELECT TOP 6 " & READ_COLUMNS & " FROM [Readings] ORDER BY [RecordedAt] DESC, [Id] DESC")
tableRows = ReadRows(command, 6)
showActions = False
connection.Close
Set connection = Nothing
pageTitle = "Your health, in view"
activePage = "dashboard"
%>
<!--#include file="includes/header.asp" -->
<section class="page-heading">
    <div><p class="eyebrow">YOUR OVERVIEW</p><h1>Your health, in view.</h1><p class="lead">A simple place to keep track of the readings that matter.</p></div>
    <span class="local-badge"><span aria-hidden="true"></span> Local journal</span>
</section>

<section class="stats-grid" aria-label="Reading summary">
    <article class="stat-card">
        <div class="stat-top"><h2>Latest blood pressure</h2><span class="stat-icon pressure-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M3 12h5l3-8 4 16 3-8h3"/></svg></span></div>
        <% If pressure Is Nothing Then %>
        <p class="stat-value">&mdash;<span>mmHg</span></p><p class="stat-detail">No blood pressure readings yet</p>
        <% Else %>
        <p class="stat-value"><%= H(pressure("Systolic")) %><span class="stat-slash">/</span><%= H(pressure("Diastolic")) %><span>mmHg</span></p>
        <p class="stat-detail"><%= H(DisplayDate(pressure("RecordedAt"))) %> at <%= H(DisplayTime(pressure("RecordedAt"))) %><% If Not IsNull(pressure("Pulse")) Then %> &middot; Pulse <%= H(pressure("Pulse")) %> bpm<% End If %></p>
        <% End If %>
    </article>
    <article class="stat-card">
        <div class="stat-top"><h2>Latest blood glucose</h2><span class="stat-icon glucose-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M12 3C10 7 5 11 5 15a7 7 0 0 0 14 0c0-4-5-8-7-12z"/></svg></span></div>
        <% If glucose Is Nothing Then %>
        <p class="stat-value">&mdash;<span><%= H(UnitLabel(unit)) %></span></p><p class="stat-detail">No blood glucose readings yet</p>
        <% Else %>
        <p class="stat-value"><%= H(CompactNumber(ConvertGlucose(glucose("GlucoseValue"), glucose("GlucoseUnit"), unit))) %><span><%= H(UnitLabel(unit)) %></span></p>
        <p class="stat-detail"><%= H(DisplayDate(glucose("RecordedAt"))) %> at <%= H(DisplayTime(glucose("RecordedAt"))) %> &middot; <%= H(ContextLabel(glucose("MealContext"))) %></p>
        <% If glucose("GlucoseUnit") <> unit Then %><p class="conversion-note">Converted from <%= H(CompactNumber(glucose("GlucoseValue"))) %> <%= H(UnitLabel(glucose("GlucoseUnit"))) %></p><% End If %>
        <% End If %>
    </article>
    <article class="stat-card stat-card-accent">
        <div class="stat-top"><h2>Readings in this period</h2><span class="stat-icon count-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M5 19V9m7 10V4m7 15v-7"/></svg></span></div>
        <p class="stat-value"><%= H(readingCount) %><span>entries</span></p>
        <p class="stat-detail"><%= H(periodLabel) %> &middot; One reading at a time</p>
    </article>
</section>

<section class="trend-section" aria-labelledby="trends-title">
    <div class="section-heading">
        <div><h2 id="trends-title">Your trends</h2><p class="muted">Measured values, without clinical interpretation.</p></div>
        <form class="view-controls" action="default.asp" method="get">
            <div><label for="range">Period</label><select name="range" id="range"><option value="7"<%= Selected(period, "7") %>>Last 7 days</option><option value="30"<%= Selected(period, "30") %>>Last 30 days</option><option value="90"<%= Selected(period, "90") %>>Last 90 days</option><option value="all"<%= Selected(period, "all") %>>All time</option></select></div>
            <div><label for="unit">Glucose units</label><select name="unit" id="unit"><option value="mmol"<%= Selected(unit, "mmol") %>>mmol/L</option><option value="mg"<%= Selected(unit, "mg") %>>mg/dL</option></select></div>
            <button class="button button-secondary button-small" type="submit">Update view</button>
        </form>
    </div>
    <div class="charts-grid">
        <article class="panel chart-card">
            <div class="chart-heading"><div><h3>Blood pressure</h3><span class="muted">mmHg</span></div><div class="chart-legend"><span><i class="legend-dot systolic-dot"></i>Systolic</span><span><i class="legend-dot diastolic-dot"></i>Diastolic</span></div></div>
            <div id="pressure-chart" class="chart-container"><p class="chart-placeholder">No blood pressure readings in this period.</p></div>
        </article>
        <article class="panel chart-card">
            <div class="chart-heading"><div><h3>Blood glucose</h3><span class="muted"><%= H(UnitLabel(unit)) %></span></div><div class="chart-legend"><span><i class="legend-dot glucose-dot"></i>Glucose</span></div></div>
            <div id="glucose-chart" class="chart-container"><p class="chart-placeholder">No blood glucose readings in this period.</p></div>
        </article>
    </div>
    <p class="chart-footnote">Each chart shows up to the latest 180 readings in the selected period. Glucose is converted for display only; original values and units are preserved.</p>
    <noscript><p class="notice">Charts need JavaScript. All saved values are available in <a href="history.asp">reading history</a>.</p></noscript>
    <script id="trend-data" type="application/json">{"unit":"<%= UnitLabel(unit) %>","pressure":<%= pressureJson %>,"glucose":<%= glucoseJson %>}</script>
</section>

<section class="panel recent-panel" aria-labelledby="recent-title">
    <div class="section-heading panel-heading"><div><h2 id="recent-title">Recent readings</h2><p class="muted">Your latest entries, across all dates.</p></div><a class="text-link" href="history.asp">View all readings <span aria-hidden="true">&rarr;</span></a></div>
    <!--#include file="includes/reading-table.asp" -->
</section>
<section class="journal-banner">
    <div><h2>Make room for a small daily habit.</h2><p>Keep the date, time, and context together so your records are easier to discuss with your healthcare professional.</p></div>
    <div class="banner-actions"><a class="button button-secondary" href="reading.asp?type=pressure">+ Blood pressure</a><a class="button button-secondary" href="reading.asp?type=glucose">+ Blood glucose</a></div>
</section>
<!--#include file="includes/footer.asp" -->
