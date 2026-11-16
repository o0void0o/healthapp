# Health Journal

A single-person, local web application built with **Classic ASP (VBScript)** and
a real **Microsoft Access `.mdb` database**. No Node.js, Microsoft Access desktop
application, cloud service, external fonts, or JavaScript libraries are needed.

## Start

Double-click **`Start.cmd`**. Keep the console window open while using the app.
It opens **http://localhost:8085/** after the site is ready. Press **Ctrl+C** in
the console to stop it.

Alternatively, in Windows PowerShell:

```powershell
Set-Location 'C:\temp\heathAPP'
.\Start-HealthApp.ps1

# Choose another port, or do not automatically open a browser:
.\Start-HealthApp.ps1 -Port 8090 -NoBrowser
```

The supplied private IIS Express x86 runtime lives in `.runtime`. It is extracted
from Microsoft's signed IIS Express 10 package, without installing IIS or changing
Windows features. The initial database is empty; no demonstration medical records
are included.

If copying just the source files to another 64-bit Windows 10/11 computer, prepare
the database and download the runtime once:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Setup.ps1 -DownloadRuntime
```

The download requires internet access; normal operation does not. Setup preserves
an existing database and verifies its schema. It automatically uses 32-bit Windows
PowerShell for the built-in Microsoft Jet 4.0 OLE DB provider. IIS Express also
needs to be **32-bit**, even on 64-bit Windows.

## Features

- Capture blood pressure, optional pulse, measurement date/time, and notes.
- Capture blood glucose in **mmol/L or mg/dL**, with optional meal context.
- Review latest measurements and 7-, 30-, 90-day, or all-time trend charts.
- Filter history by type and inclusive start/end dates; 20 entries per page.
- Edit records and delete them through an explicit confirmation.
- Download all matching records as UTF-8 CSV, including notes and original units.
- Responsive layouts, keyboard-friendly forms, accessible chart descriptions, and
  printable pages. Capture and history work without JavaScript; charts require it.

The chart display conversion is `mg/dL = mmol/L * 18.0182`. Original glucose
values and units are never replaced by converted values. Display conversions are
rounded; exports retain the original value to the two decimal places accepted
by the entry form. Changing the units in an entry form **does not convert** the
number entered.

Dates are local wall-clock times on the computer running the app, not UTC.
Dates before 1900, impossible dates, and times more than five minutes ahead of
the server clock are rejected. Numeric bounds are broad typing checks, **not
clinical targets**. The app does not classify readings or recommend treatment.

## Files and storage

```text
heathAPP\
  Start.cmd                 Double-click launcher
  Start-HealthApp.ps1        Local IIS Express configuration and readiness checks
  Setup.ps1                 Access schema creation and optional runtime download
  data\
    health.mdb              Your records; outside the web root
  www\
    default.asp             Dashboard and trends
    reading.asp             Add/edit forms
    history.asp             Filtering and pagination
    delete.asp              Confirmed, POST-only deletion
    export.asp              CSV downloads
    includes\               Shared validation, database access, and layout
    assets\                 Local stylesheet, JavaScript, and favicon
    web.config              Request limits and security headers
  tests\                    Validation and live HTTP checks
  .runtime\                 Private web server, generated configuration, and logs
```

`Readings` contains an AutoNumber ID; reading type; local measurement timestamp;
nullable systolic, diastolic, pulse, glucose value, glucose unit, and context;
notes; and created/updated timestamps. Date and type/date indexes support history
queries. Only the fields for the selected measurement type are stored.

## Privacy and limitations

This is a **personal, localhost-only application**, not a shared patient system.
It has **no application login**. Any person able to use this computer or connect
to its localhost services may be able to access the running journal.

- The server binds to `localhost`, and ASP checks both the host and loopback
  address. Do not remove these restrictions or put the app behind a public proxy.
- Database access uses parameterized SQL. State-changing forms require a
  session-bound token, and cross-origin writes are rejected.
- Displayed user text is HTML-encoded. CSV formula-like text is prefixed with an
  apostrophe to prevent spreadsheet formula execution.
- The database is outside `www`; includes, database extensions, and directory
  browsing are blocked. Health pages send `Cache-Control: no-store`.
- A newly created `data` directory grants access only to the current Windows
  user, SYSTEM, and local Administrators. Existing directory permissions are not
  changed. The Access file and CSV downloads are **not encrypted by this app**.
  Protect the Windows account, use device encryption, and secure exported files.
- No telemetry or external web requests are made by the application.
- This is not a medical device, diagnostic tool, emergency service, or a claim of
  compliance with healthcare privacy regulations.

## Backups

Stop the app first, then copy `data\health.mdb` to a secure backup location.
Keep the original file until the backup is verified. Do not copy the temporary
`.ldb` lock file. CSV exports are useful for sharing records, but a database copy
is the full backup. CSV import is not included.

To restore, stop the app, keep a copy of the current database, and replace only
`data\health.mdb` with your verified backup.

## Validation

Pure VBScript validation tests:

```powershell
& "$env:WINDIR\SysWOW64\cscript.exe" //nologo .\tests\validation-tests.vbs
```

With the app running, use another PowerShell window:

```powershell
.\tests\http-tests.ps1
# Or: .\tests\http-tests.ps1 -BaseUri 'http://localhost:8090'
```

The HTTP suite creates uniquely marked temporary records and removes them in a
cleanup block, including after a test failure. It checks capture, persistence,
editing, deletion, optional fields, conversion, filtering, pagination, output
encoding, CSV safety, and rejected invalid requests. Use a disposable database
for tests rather than a valuable personal journal.

## Troubleshooting / using full IIS

- **Port already in use:** close the other instance or start with `-Port 8090`.
- **Script execution blocked:** launch `Start.cmd`; it sets execution policy for
  that PowerShell process only.
- **Database/provider errors:** use the included x86 server. Confirm the current
  Windows user can modify the entire `data` folder, including creating lock files.
  Setup never replaces an existing or incompatible database.
- **Expired form:** reload it, enable cookies, and submit again. Sessions time
  out after 30 minutes of inactivity.
- **Secure-token error:** Classic ASP uses Windows `Scriptlet.TypeLib` to create
  GUID-based session tokens. This component is available on this computer; if an
  organization's policy disables it elsewhere, do not bypass that policy.
- **Server errors:** see the console, private `data\errors.log` (operation, error
  number, and provider message), and `.runtime\logs`. Failures also append a
  `HealthJournal-...` identifier and numeric code to the IIS log. Detailed ASP
  errors are not sent to the browser. Protect logs along with the database.

For full IIS, enable Classic ASP, use a dedicated **32-bit** application pool
(No Managed Code / Classic), point the site at `www`, and bind it only to
`localhost`. Keep `data` beside `www`, never underneath it, and grant the
application-pool identity Modify access to `data`. Configure ASP error logging
and preserve `web.config`. No public or multi-user deployment is supported.
