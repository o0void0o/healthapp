[CmdletBinding()]
param(
    [ValidateRange(1024, 65535)][int]$Port = 8085,
    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
& (Join-Path $PSScriptRoot 'Setup.ps1')

$candidates = @(
    (Join-Path $PSScriptRoot '.runtime\iis\IIS Express\iisexpress.exe'),
    (Join-Path ${env:ProgramFiles(x86)} 'IIS Express\iisexpress.exe')
)
$iis = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $iis) { throw 'Run .\Setup.ps1 -DownloadRuntime to prepare the private IIS Express x86 runtime.' }

$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
try { $listener.Start() } finally { $listener.Stop() }

$state = Join-Path $PSScriptRoot '.runtime\state'
$logs = Join-Path $PSScriptRoot '.runtime\logs'
$cache = Join-Path $state 'asp-cache'
New-Item -ItemType Directory -Path $state, $logs, $cache -Force | Out-Null
$env:IIS_BIN = Split-Path -Parent $iis
$env:IIS_USER_HOME = $state
$env:IIS_SITES_HOME = $PSScriptRoot
$template = Join-Path $env:IIS_BIN 'AppServer\applicationhost.config'
[xml]$config = Get-Content -LiteralPath $template -Raw
$site = $config.SelectSingleNode('/configuration/system.applicationHost/sites/site')
$site.SetAttribute('name', 'HealthJournal')
$site.SelectSingleNode('application').SetAttribute('applicationPool', 'UnmanagedClassicAppPool')
$site.SelectSingleNode('application/virtualDirectory').SetAttribute('physicalPath', (Join-Path $PSScriptRoot 'www'))
$site.SelectSingleNode('bindings/binding').SetAttribute('bindingInformation', ":${Port}:localhost")
$config.SelectSingleNode('/configuration/system.applicationHost/sites/applicationDefaults').SetAttribute('applicationPool', 'UnmanagedClassicAppPool')
$log = $config.SelectSingleNode('/configuration/system.applicationHost/sites/siteDefaults/logFile')
$log.SetAttribute('directory', $logs)
$log.SetAttribute('enabled', 'true')
$log.SetAttribute('logExtFileFlags', 'Date,Time,ClientIP,UriStem,UriQuery,HttpStatus,HttpSubStatus,Win32Status,TimeTaken')
$asp = $config.SelectSingleNode('/configuration/system.webServer/asp')
$asp.SetAttribute('scriptErrorSentToBrowser', 'false')
$asp.SelectSingleNode('cache').SetAttribute('diskTemplateCacheDirectory', $cache)
$history = $config.SelectSingleNode('/configuration/system.applicationHost/configHistory')
if ($null -eq $history) {
    $history = $config.CreateElement('configHistory')
    [void]$config.SelectSingleNode('/configuration/system.applicationHost').AppendChild($history)
}
$history.SetAttribute('enabled', 'false')
$configPath = Join-Path $state 'applicationhost.config'
$config.Save($configPath)
$configCheck = & (Join-Path $env:IIS_BIN 'appcmd.exe') list config 'HealthJournal/' /section:system.webServer/security/requestFiltering "/apphostconfig:$configPath" 2>&1
if ($LASTEXITCODE -ne 0) { throw "IIS configuration is invalid:`n$($configCheck -join [Environment]::NewLine)" }

$url = "http://localhost:$Port/"
$arguments = '/config:"' + $configPath + '" /site:HealthJournal /systray:false'
$process = $null
try {
    $process = Start-Process -FilePath $iis -ArgumentList $arguments -NoNewWindow -PassThru
    $ready = $false
    for ($attempt = 0; $attempt -lt 40; $attempt++) {
        if ($process.HasExited) { throw "IIS Express exited with code $($process.ExitCode)." }
        try {
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 2
            if ($response.StatusCode -eq 200) { $ready = $true; break }
        } catch [Net.WebException] {
            if ($null -ne $_.Exception.Response) { throw }
        }
        Start-Sleep -Milliseconds 250
    }
    if (-not $ready) { throw 'The site did not become ready. Check .runtime\logs and the console output.' }
    Write-Host "`nHealth Journal is ready: $url" -ForegroundColor Cyan
    Write-Host 'Local computer only. Keep this window open; press Ctrl+C to stop.'
    if (-not $NoBrowser) { Start-Process $url }
    $process.WaitForExit()
} finally {
    if ($null -ne $process -and -not $process.HasExited) { Stop-Process -Id $process.Id }
}
