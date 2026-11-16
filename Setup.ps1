[CmdletBinding()]
param([switch]$DownloadRuntime)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([Environment]::Is64BitProcess) {
    $x86PowerShell = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $x86PowerShell)) {
        throw 'This app requires the 32-bit Windows PowerShell and Microsoft Jet components.'
    }
    $forward = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    if ($DownloadRuntime) { $forward += '-DownloadRuntime' }
    & $x86PowerShell @forward
    if ($LASTEXITCODE -ne 0) { throw 'Health Journal setup failed. See the error above.' }
    return
}

$root = $PSScriptRoot
$data = Join-Path $root 'data'
$runtime = Join-Path $root '.runtime'
$database = Join-Path $data 'health.mdb'
$privateIis = Join-Path $runtime 'iis\IIS Express\iisexpress.exe'
$installedIis = Join-Path ${env:ProgramFiles} 'IIS Express\iisexpress.exe'

if ($DownloadRuntime -and -not (Test-Path -LiteralPath $privateIis) -and -not (Test-Path -LiteralPath $installedIis)) {
    New-Item -ItemType Directory -Path $runtime -Force | Out-Null
    $package = Join-Path $runtime 'iisexpress_x86_en-US.msi'
    $extractLog = Join-Path $runtime 'extract.log'
    $destination = Join-Path $runtime 'iis'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-Host 'Downloading the official Microsoft IIS Express x86 package...'
    Invoke-WebRequest -UseBasicParsing -Uri 'https://download.microsoft.com/download/c/e/8/ce8d18f5-d4c0-45b5-b531-adecd637a1aa/iisexpress_x86_en-US.msi' -OutFile $package
    $signature = Get-AuthenticodeSignature -FilePath $package
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        throw 'Microsoft package signature verification failed. The package was not run.'
    }
    $arguments = '/a "' + $package + '" /qn TARGETDIR="' + $destination + '" /L*v "' + $extractLog + '"'
    $process = Start-Process -FilePath (Join-Path $env:WINDIR 'System32\msiexec.exe') -ArgumentList $arguments -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "IIS Express extraction failed ($($process.ExitCode)). See $extractLog." }
    if (-not (Test-Path -LiteralPath $privateIis)) { throw 'The extracted IIS Express executable could not be found.' }
    Remove-Item -LiteralPath $package, $extractLog
}

if (-not (Test-Path -LiteralPath $data)) {
    New-Item -ItemType Directory -Path $data | Out-Null
    $acl = New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true, $false)
    $identities = @(
        [Security.Principal.WindowsIdentity]::GetCurrent().User,
        (New-Object Security.Principal.SecurityIdentifier('S-1-5-18')),
        (New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544'))
    )
    foreach ($identity in $identities) {
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $data -AclObject $acl
}

$connection = $null
$catalog = $null
$temporary = $null
try {
    if (Test-Path -LiteralPath $database) {
        $connection = New-Object -ComObject ADODB.Connection
        $connection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$database;")
        $check = $connection.Execute('SELECT TOP 1 [Id], [ReadingType], [RecordedAt], [Systolic], [Diastolic], [Pulse], [GlucoseValue], [GlucoseUnit], [MealContext], [Notes], [CreatedAt], [UpdatedAt] FROM [Readings]')
        $check.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($check)
        Write-Host 'Existing Access database verified. Your readings were preserved.'
    } else {
        $temporary = Join-Path $data ('health-' + [guid]::NewGuid().ToString('N') + '.mdb')
        $catalog = New-Object -ComObject ADOX.Catalog
        [void]$catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$temporary;Jet OLEDB:Engine Type=5;")
        $connection = $catalog.ActiveConnection
        [void]$connection.Execute(@'
CREATE TABLE [Readings] (
    [Id] COUNTER CONSTRAINT [PK_Readings] PRIMARY KEY,
    [ReadingType] TEXT(8) NOT NULL,
    [RecordedAt] DATETIME NOT NULL,
    [Systolic] SMALLINT,
    [Diastolic] SMALLINT,
    [Pulse] SMALLINT,
    [GlucoseValue] DOUBLE,
    [GlucoseUnit] TEXT(4),
    [MealContext] TEXT(24),
    [Notes] MEMO,
    [CreatedAt] DATETIME NOT NULL,
    [UpdatedAt] DATETIME NOT NULL
)
'@)
        [void]$connection.Execute('CREATE INDEX [IX_Readings_Date] ON [Readings] ([RecordedAt], [Id])')
        [void]$connection.Execute('CREATE INDEX [IX_Readings_TypeDate] ON [Readings] ([ReadingType], [RecordedAt])')
        Write-Host 'Created a new, empty Microsoft Access database.'
    }
} finally {
    if ($null -ne $connection) {
        if ($connection.State -ne 0) { $connection.Close() }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($connection)
    }
    if ($null -ne $catalog) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($catalog) }
}
if ($null -ne $temporary) { Move-Item -LiteralPath $temporary -Destination $database }
Write-Host "Database: $database"
if (-not (Test-Path -LiteralPath $privateIis) -and -not (Test-Path -LiteralPath $installedIis)) {
    Write-Warning 'No x86 IIS Express runtime found. Run .\Setup.ps1 -DownloadRuntime before starting.'
}
