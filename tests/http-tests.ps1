[CmdletBinding()]
param([string]$BaseUri = 'http://localhost:8085')

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Net.Http
$BaseUri = $BaseUri.TrimEnd('/')
if (([uri]$BaseUri).Host -ne 'localhost') { throw 'These tests only run against a localhost instance.' }
$handler = [Net.Http.HttpClientHandler]::new()
$handler.AllowAutoRedirect = $false
$handler.UseProxy = $false
$handler.CookieContainer = [Net.CookieContainer]::new()
$client = [Net.Http.HttpClient]::new($handler)
$client.Timeout = [TimeSpan]::FromSeconds(15)
$marker = 'healthapp-test-' + [guid]::NewGuid().ToString('N')
$knownIds = [Collections.Generic.HashSet[int]]::new()
$script:checks = 0

function Assert($Condition, [string]$Message) {
    $script:checks++
    if (-not $Condition) { throw "FAIL: $Message" }
}

function Send([string]$Path, [hashtable]$Body = $null, [string]$Origin = $BaseUri) {
    $method = if ($null -eq $Body) { [Net.Http.HttpMethod]::Get } else { [Net.Http.HttpMethod]::Post }
    $request = [Net.Http.HttpRequestMessage]::new($method, "$BaseUri/$Path")
    $response = $null
    try {
        if ($null -ne $Body) {
            $fields = [Collections.Generic.Dictionary[string, string]]::new()
            foreach ($key in $Body.Keys) { $fields.Add($key, [string]$Body[$key]) }
            $request.Content = [Net.Http.FormUrlEncodedContent]::new($fields)
            [void]$request.Headers.TryAddWithoutValidation('Origin', $Origin)
        }
        $response = $client.SendAsync($request).GetAwaiter().GetResult()
        return @{
            Status = [int]$response.StatusCode
            Text = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            Headers = $response.Headers.ToString()
            ContentType = [string]$response.Content.Headers.ContentType
        }
    } finally {
        if ($null -ne $response) { $response.Dispose() }
        $request.Dispose()
    }
}

function Token([string]$Path = 'reading.asp') {
    $response = Send $Path
    Assert ($response.Status -eq 200) "Form loads: $Path (status $($response.Status))"
    $match = [regex]::Match($response.Text, 'name="csrf" value="([0-9A-Fa-f]{64})"')
    Assert $match.Success 'Session CSRF token is present'
    return $match.Groups[1].Value
}

function Csv([string]$Query = '') {
    $response = Send "export.asp$Query"
    Assert ($response.Status -eq 200) 'CSV export succeeds'
    Assert ($response.ContentType -like 'text/csv*') 'CSV content type is correct'
    return @($response.Text.TrimStart([char]0xFEFF) | ConvertFrom-Csv)
}

function Pressure([string]$Notes, [string]$When) {
    return @{
        id = '0'; type = 'pressure'; recorded_at = $When
        systolic = '120'; diastolic = '80'; pulse = '72'; notes = $Notes
        csrf = (Token)
    }
}

function FindTestRow([string]$Notes) {
    $rows = @(Csv | Where-Object { $_.Notes -eq $Notes })
    Assert ($rows.Count -eq 1) "Exactly one saved reading found for $Notes"
    [void]$knownIds.Add([int]$rows[0].Id)
    return $rows[0]
}

try {
    $dashboard = Send ''
    Assert ($dashboard.Status -eq 200) 'Dashboard renders'
    Assert ($dashboard.Text.Contains('Your health, in view.')) 'Dashboard has expected title'
    Assert ($dashboard.Headers -match 'Content-Security-Policy:') 'CSP header is present'
    Assert ($dashboard.Headers -match 'no-store') 'Health pages are not cached'
    Assert ($dashboard.Headers -match 'Referrer-Policy: same-origin') 'Referrer policy preserves the origin on same-site form posts'
    $baseline = @(Csv).Count
    foreach ($asset in @('assets/app.css', 'assets/app.js', 'assets/favicon.svg')) {
        Assert ((Send $asset).Status -eq 200) "Static asset loads: $asset"
    }
    foreach ($path in @('includes/validation.asp', 'data/health.mdb', '../data/health.mdb')) {
        Assert ((Send $path).Status -eq 404) "Private resource is not exposed: $path"
    }

    $when = (Get-Date).AddMinutes(-10).ToString('yyyy-MM-ddTHH:mm')
    $notes = "$marker pressure <script>alert('x')</script>"
    $body = Pressure $notes $when
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Blood pressure saves with POST/redirect/GET'
    $pressure = FindTestRow $notes
    Assert ($pressure.Systolic_mmHg -eq '120' -and $pressure.Diastolic_mmHg -eq '80' -and $pressure.Pulse_bpm -eq '72') 'Blood pressure persists accurately'
    $history = Send 'history.asp'
    Assert ($history.Text.Contains('&lt;script&gt;alert(&#39;x&#39;)&lt;/script&gt;')) 'User notes are HTML encoded'
    Assert (-not $history.Text.Contains("<script>alert('x')</script>")) 'Notes cannot inject executable markup'

    foreach ($item in @(
        @{ value = '5.6'; unit = 'mmol'; context = 'fasting'; suffix = 'mmol' },
        @{ value = '108.11'; unit = 'mg'; context = 'after-meal'; suffix = 'mg' }
    )) {
        $glucoseNotes = "$marker $($item.suffix)"
        $glucoseBody = @{
            id = '0'; type = 'glucose'; recorded_at = $when
            glucose_value = $item.value; glucose_unit = $item.unit
            context = $item.context; notes = $glucoseNotes; csrf = (Token 'reading.asp?type=glucose')
        }
        Assert ((Send 'reading.asp' $glucoseBody).Status -eq 303) "Glucose saves in $($item.unit)"
        $row = FindTestRow $glucoseNotes
        Assert ($row.GlucoseValue -eq $item.value) 'Original glucose precision is retained'
        Assert ($row.GlucoseUnit -eq $(if ($item.unit -eq 'mg') { 'mg/dL' } else { 'mmol/L' })) 'Original glucose unit is retained'
    }
    $converted = Send '?range=all&unit=mmol'
    $jsonText = [regex]::Match($converted.Text, '<script id="trend-data" type="application/json">(.+?)</script>').Groups[1].Value
    $chart = $jsonText | ConvertFrom-Json
    Assert (@($chart.glucose | Where-Object { [Math]::Abs($_.value - 6) -lt 0.001 }).Count -ge 1) 'Mixed glucose units are normalized for charts'

    $body.id = $pressure.Id
    $body.systolic = '128'
    $body.diastolic = '82'
    $body.notes = "$marker edited"
    $body.csrf = Token "reading.asp?id=$($pressure.Id)"
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Existing reading can be edited'
    $updated = FindTestRow "$marker edited"
    Assert ($updated.Id -eq $pressure.Id -and $updated.Systolic_mmHg -eq '128') 'Edit updates the same record'
    $body.notes = ''
    $body.pulse = ''
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Blank optional notes and pulse save as null'
    $emptyOptional = @(Csv | Where-Object { $_.Id -eq $pressure.Id })[0]
    Assert ($emptyOptional.Notes -eq '' -and $emptyOptional.Pulse_bpm -eq '') 'Optional values are cleared, not retained'
    $body.notes = "$marker edited"
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Notes can be restored after clearing'
    $unicodeNotes = 'caf' + [char]0x00E9 + "`r`n" + [char]0x6E2C + [char]0x8A66 + ', "quoted"'
    $body.notes = $unicodeNotes
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Unicode and multiline notes save'
    $unicodeRow = @(Csv | Where-Object { $_.Id -eq $pressure.Id })[0]
    Assert ($unicodeRow.Notes -eq $unicodeNotes) 'Unicode, line breaks, commas, and quotes round-trip through Access and CSV'
    $body.notes = ([string][char]0x6E2C) * 2000
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Maximum-length UTF-8 notes fit within the request limit'
    $unicodeRow = @(Csv | Where-Object { $_.Id -eq $pressure.Id })[0]
    Assert ($unicodeRow.Notes -eq $body.notes) 'Maximum-length Unicode memo is stored without truncation'
    $body.notes = "$marker edited"
    Assert ((Send 'reading.asp' $body).Status -eq 303) 'Marked notes are restored after Unicode checks'

    $invalid = Pressure "$marker invalid" $when
    $invalid.diastolic = '130'
    Assert ((Send 'reading.asp' $invalid).Status -eq 422) 'Reversed blood pressure is rejected'
    $invalid.diastolic = '80'
    $invalid.recorded_at = '2026-02-31T08:30'
    Assert ((Send 'reading.asp' $invalid).Status -eq 422) 'Impossible date is rejected'
    $invalid.recorded_at = (Get-Date).AddMinutes(20).ToString('yyyy-MM-ddTHH:mm')
    Assert ((Send 'reading.asp' $invalid).Status -eq 422) 'Future date is rejected'
    $invalid.recorded_at = $when
    $invalid.csrf = 'wrong-token'
    Assert ((Send 'reading.asp' $invalid).Status -eq 403) 'Invalid CSRF token is rejected'
    $invalid.csrf = Token
    Assert ((Send 'reading.asp' $invalid 'https://example.invalid').Status -eq 403) 'Cross-origin write is rejected'
    Assert ((Send 'reading.asp' $invalid 'null').Status -eq 403) 'Opaque-origin write is rejected'
    $invalid.notes = 'x' * 2001
    Assert ((Send 'reading.asp' $invalid).Status -eq 422) 'Oversized notes are rejected'
    Assert (@(Csv).Count -eq $baseline + 3) 'Invalid requests never add readings'

    foreach ($path in @(
        'reading.asp?id=1%20OR%201%3D1',
        'history.asp?from=2026-02-31',
        'export.asp?from=2026-02-31',
        'history.asp?from=2026-11-16&to=2026-11-15',
        'export.asp?type=invalid',
        '?range=bad',
        '?unit=invalid'
    )) {
        Assert ((Send $path).Status -eq 400) "Invalid query is rejected: $path"
    }
    Assert ((Send 'reading.asp?id=2147483647').Status -eq 404) 'Missing reading returns 404'
    $day = $when.Substring(0, 10)
    $filtered = @(Csv "?type=glucose&from=$day&to=$day")
    Assert (@($filtered | Where-Object { $_.Notes -like "$marker*" }).Count -eq 2) 'Date filter includes the entire end date'
    Assert (@($filtered | Where-Object { $_.Type -ne 'glucose' }).Count -eq 0) 'Type filter is applied to CSV export'

    $formula = Pressure "$marker formula" $when
    $formula.notes = '=2+2 ' + $marker
    Assert ((Send 'reading.asp' $formula).Status -eq 303) 'Formula-like notes save as ordinary text'
    $formulaRows = @(Csv | Where-Object { $_.Notes -eq ("'" + $formula.notes) })
    Assert ($formulaRows.Count -eq 1) 'Spreadsheet formulas are neutralized on export'
    [void]$knownIds.Add([int]$formulaRows[0].Id)

    $pageToken = Token
    for ($index = 1; $index -le 21; $index++) {
        $pageBody = @{
            id = '0'; type = 'pressure'; recorded_at = '2001-02-03T08:30'
            systolic = '120'; diastolic = '80'; pulse = ''
            notes = "$marker page $index"; csrf = $pageToken
        }
        Assert ((Send 'reading.asp' $pageBody).Status -eq 303) "Pagination fixture $index saves"
    }
    $query = 'type=pressure&from=2001-02-03&to=2001-02-03'
    $pageOne = Send "history.asp?$query"
    $pageTwo = Send "history.asp?$query&page=2"
    Assert ($pageOne.Status -eq 200 -and $pageTwo.Status -eq 200) 'Both history pages load'
    Assert ([regex]::Matches($pageOne.Text, '<tr>').Count -eq 21) 'History page contains exactly 20 readings plus its header'
    Assert ($pageTwo.Text.Contains('Page 2 of')) 'History pagination advances to page two'
    $nextLink = [regex]::Match($pageOne.Text, 'href="(history\.asp\?[^"]*&amp;page=2)"')
    $decodedLink = [uri]::UnescapeDataString([Net.WebUtility]::HtmlDecode($nextLink.Groups[1].Value))
    Assert ($nextLink.Success -and $decodedLink.Contains("$query&page=2")) 'Pagination preserves date filters'

    $beforeDelete = @(Csv).Count
    $deleteToken = Token "delete.asp?id=$($pressure.Id)"
    Assert (@(Csv).Count -eq $beforeDelete) 'Viewing delete confirmation does not delete data'
    Assert ((Send 'delete.asp' @{ id = $pressure.Id; csrf = 'invalid' }).Status -eq 403) 'Delete requires a valid CSRF token'
    Assert ((Send 'delete.asp' @{ id = $pressure.Id; csrf = $deleteToken }).Status -eq 303) 'Confirmed delete succeeds'
    [void]$knownIds.Remove([int]$pressure.Id)
    Assert ((Send "reading.asp?id=$($pressure.Id)").Status -eq 404) 'Deleted reading is no longer available'
    Write-Host "$script:checks HTTP checks passed." -ForegroundColor Green
} finally {
    $cleanupFailures = [Collections.Generic.List[string]]::new()
    try {
        foreach ($row in @(Csv)) {
            if ($row.Notes.Contains($marker)) { [void]$knownIds.Add([int]$row.Id) }
        }
        foreach ($id in $knownIds) {
            try {
                $token = Token "delete.asp?id=$id"
                $result = Send 'delete.asp' @{ id = $id; csrf = $token }
                if ($result.Status -ne 303) { throw "Delete returned $($result.Status)" }
            } catch {
                $cleanupFailures.Add("Reading ${id}: $($_.Exception.Message)")
            }
        }
    } finally {
        $client.Dispose()
        $handler.Dispose()
    }
    if ($cleanupFailures.Count -gt 0) { throw "Test cleanup failed: $($cleanupFailures -join '; ')" }
    Write-Host 'Temporary test readings removed.'
}
