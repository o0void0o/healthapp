<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="color-scheme" content="light">
    <title><%= H(pageTitle) %> - Health Journal</title>
    <link rel="icon" href="assets/favicon.svg" type="image/svg+xml">
    <link rel="stylesheet" href="assets/app.css">
    <script src="assets/app.js" defer></script>
</head>
<body>
<a class="skip-link" href="#main">Skip to content</a>
<div class="app-shell">
    <aside class="sidebar" aria-label="Main navigation">
        <a class="brand" href="default.asp" aria-label="Health Journal dashboard">
            <svg class="brand-mark" viewBox="0 0 40 40" aria-hidden="true"><rect width="40" height="40" rx="12"/><path d="M8 21h7l3-9 5 17 3-8h6"/></svg>
            <span>health<span class="brand-sub">journal</span></span>
        </a>
        <p class="nav-caption">YOUR SPACE</p>
        <nav class="nav-links">
            <a class="nav-link<% If activePage = "dashboard" Then Response.Write " active" %>" href="default.asp"<% If activePage = "dashboard" Then Response.Write " aria-current=""page""" %>>
                <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></svg>Overview
            </a>
            <a class="nav-link<% If activePage = "history" Then Response.Write " active" %>" href="history.asp"<% If activePage = "history" Then Response.Write " aria-current=""page""" %>>
                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 4h11v17H5V4h3m0-2h8v4H8zM9 11h6m-6 5h6"/></svg>Reading history
            </a>
            <a class="nav-link<% If activePage = "add" Then Response.Write " active" %>" href="reading.asp"<% If activePage = "add" Then Response.Write " aria-current=""page""" %>>
                <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="5"/><path d="M12 7v10M7 12h10"/></svg>Add a reading
            </a>
        </nav>
        <div class="sidebar-note">
            <span class="local-dot" aria-hidden="true"></span>
            <strong>On this computer</strong>
            <p>Your readings are stored locally in a Microsoft Access database.</p>
        </div>
        <p class="sidebar-bottom">A little clarity, every day.</p>
    </aside>
    <div class="workspace">
        <header class="topbar">
            <span class="workspace-label">PERSONAL HEALTH TRACKER</span>
            <div class="topbar-actions">
                <span class="today"><%= H(DisplayDate(Date())) %></span>
                <% If activePage <> "add" Then %><a class="button button-small" href="reading.asp"><span aria-hidden="true">+</span> Add reading</a><% End If %>
            </div>
        </header>
        <main id="main" class="main-content">
        <%
        Dim flashMessage
        flashMessage = Text(Session("flash"))
        Session("flash") = ""
        If flashMessage <> "" Then
        %>
            <div class="notice notice-success" role="status"><%= H(flashMessage) %></div>
        <% End If %>
