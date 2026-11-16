<% If IsEmpty(tableRows) Then %>
<div class="empty-state">
    <span class="empty-symbol" aria-hidden="true">+</span>
    <h3>No readings to show</h3>
    <p>Add your first reading, or try a different date range.</p>
    <a class="button button-secondary" href="reading.asp">Add a reading</a>
</div>
<% Else %>
<div class="table-scroll" role="region" aria-label="Readings table" tabindex="0">
<table class="readings-table">
    <caption class="visually-hidden">Recorded blood pressure and blood glucose readings</caption>
    <thead><tr><th scope="col">Date &amp; time</th><th scope="col">Type</th><th scope="col">Reading</th><th scope="col">Context / pulse</th><th scope="col">Notes</th><% If showActions Then %><th scope="col">Actions</th><% End If %></tr></thead>
    <tbody>
    <%
    Dim tableIndex
    For tableIndex = 0 To UBound(tableRows, 2)
    %>
    <tr>
        <th scope="row" class="date-cell"><time datetime="<%= H(IsoTimestamp(tableRows(2, tableIndex))) %>"><%= H(DisplayDate(tableRows(2, tableIndex))) %><span class="cell-sub"><%= H(DisplayTime(tableRows(2, tableIndex))) %></span></time></th>
        <td><span class="type-badge <%= H(tableRows(1, tableIndex)) %>"><span aria-hidden="true"></span><% If tableRows(1, tableIndex) = "pressure" Then %>Blood pressure<% Else %>Blood glucose<% End If %></span></td>
        <td class="measurement-cell">
            <% If tableRows(1, tableIndex) = "pressure" Then %>
            <strong><%= H(tableRows(3, tableIndex)) %><span class="reading-divider"> / </span><%= H(tableRows(4, tableIndex)) %></strong><span class="cell-unit">mmHg</span>
            <% Else %>
            <strong><%= H(CompactNumber(tableRows(6, tableIndex))) %></strong><span class="cell-unit"><%= H(UnitLabel(tableRows(7, tableIndex))) %></span>
            <% End If %>
        </td>
        <td class="muted">
            <% If tableRows(1, tableIndex) = "pressure" Then %>
                <% If IsNull(tableRows(5, tableIndex)) Then %><span aria-label="No pulse recorded">&mdash;</span><% Else %><%= H(tableRows(5, tableIndex)) %> bpm<% End If %>
            <% Else %>
                <%= H(ContextLabel(tableRows(8, tableIndex))) %>
            <% End If %>
        </td>
        <td><% If Len(Text(tableRows(9, tableIndex))) = 0 Then %><span class="muted" aria-label="No notes">&mdash;</span><% Else %><span class="note-text" title="<%= H(tableRows(9, tableIndex)) %>"><%= H(tableRows(9, tableIndex)) %></span><% End If %></td>
        <% If showActions Then %>
        <td class="row-actions"><a href="reading.asp?id=<%= H(tableRows(0, tableIndex)) %>" aria-label="Edit reading <%= H(tableRows(0, tableIndex)) %>">Edit</a><a class="danger-link" href="delete.asp?id=<%= H(tableRows(0, tableIndex)) %>" aria-label="Delete reading <%= H(tableRows(0, tableIndex)) %>">Delete</a></td>
        <% End If %>
    </tr>
    <% Next %>
    </tbody>
</table>
</div>
<% End If %>
