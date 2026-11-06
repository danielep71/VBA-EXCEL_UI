<#
.SYNOPSIS
Certify the checked-out commit in Excel with one command and write an evidence
bundle.

.DESCRIPTION
For the commit checked out in this repository:

1. preflight: reads the commit SHA and refuses a working tree with changes
   (unless -AllowDirty, which the bundle records); creates a new bundle folder
   next to the checkout; reads Windows and the Trust Center settings without
   changing them, and stops when "Trust access to the VBA project object
   model" is off; writes source-digests.json from the commit;
2. build: creates a development workbook in the bundle with
   tools\Sync-EXCEL_UI-VBA.ps1 -Mode Development -NewWorkbook;
3. open: opens it in a visible Excel instance of its own, which follows the
   Trust Center macro setting, and writes host.json;
4. enable-macros: waits for the operator to click Enable Content when the
   macros are disabled;
5. compile: runs Debug > Compile VBAProject and stops when the project does not
   compile;
6. certify: runs Test_EXCEL_UI_RunReleaseCertification;
7. collect: copies the JSON document and the text report that run wrote to the
   temporary folder into the bundle;
8. close: closes the workbook without saving and quits the Excel instance it
   started.

It then writes session.json and validates the bundle with
tools\check_excel_evidence.py when Python 3.11 or later is available.

Exit codes: 0 the certification passed (and the bundle validated, when Python
was found); 1 it failed or stopped on an error; 2 the operator cancelled it.

See docs\EXCEL_CERTIFICATION.md.

.PARAMETER BundleDir
The bundle folder. It must be outside the checkout and new or empty. By
default, EXCEL_UI-certification-<SHA12>-<yyyyMMdd-HHmmss> next to the checkout.

.PARAMETER AllowDirty
Run even though the working tree has changes. The bundle records it, and the
validator rejects such a bundle: it cannot certify a commit.

.PARAMETER Python
The Python 3.11+ executable that validates the bundle. By default the script
tries py -3, python and python3.

.PARAMETER RepoRoot
The repository. By default, the parent of this script's folder.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-EXCEL_UI.ps1

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-EXCEL_UI.ps1 -BundleDir D:\Evidence\EXCEL_UI-1.1.3
#>
[CmdletBinding()]
param(
    [string]$BundleDir,
    [switch]$AllowDirty,
    [string]$Python,
    [string]$RepoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Latin1 = [System.Text.Encoding]::GetEncoding(28591)

$MsoAutomationSecurityByUI = 2
$CompileControlId = 578
$WorkbookFileName = 'EXCEL_UI-development.xlsm'
$CertificationMacro = 'Test_EXCEL_UI_RunReleaseCertification'
# A read-only public function: it runs only when the macros are enabled.
$ReadinessMacro = 'M_EXCEL_UI.UI_HasExcelUIStateSnapshot'
# The files TST_CertTryWriteEvidence writes; the Ribbon probe's
# .ribbon.json and .ribbon.txt do not match.
$EvidenceNamePattern = '^EXCEL_UI_certification_(\d{8}_\d{6})\.(json|txt)$'
$script:SessionLog = $null
$script:CurrentStep = 'preflight'
$StepNames = @('preflight', 'build', 'open', 'enable-macros', 'compile', 'certify', 'collect', 'close')

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-UtcStamp([datetime]$At = (Get-Date)) {
    return $At.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $Invariant)
}

function Get-LocalStamp([datetime]$At = (Get-Date)) {
    return $At.ToString("yyyy-MM-dd'T'HH:mm:sszzz", $Invariant)
}

function Write-Log([string]$Text) {
    $line = '{0}  {1}' -f (Get-Date).ToString('HH:mm:ss', $Invariant), $Text
    Write-Host $line
    if ($script:SessionLog) {
        [System.IO.File]::AppendAllText($script:SessionLog, $line + "`r`n", $Utf8NoBom)
    }
}

function Set-Step([string]$Name, [string]$Status, [string]$Detail) {
    foreach ($step in $script:Session.steps) {
        if ($step.name -eq $Name) {
            $step.status = $Status
            $step.detail = $Detail
            $step.at_utc = Get-UtcStamp
        }
    }
    $script:CurrentStep = $Name
    Write-Log ('[{0}] {1}: {2}' -f $Status, $Name, $Detail)
}

function Get-StepStatus([string]$Name) {
    foreach ($step in $script:Session.steps) {
        if ($step.name -eq $Name) {
            return [string]$step.status
        }
    }
    return ''
}

function Stop-Cancelled([string]$Message) {
    throw (New-Object System.OperationCanceledException($Message))
}

function Get-InnerMessage($ErrorRecord) {
    $exception = $ErrorRecord.Exception
    while ($null -ne $exception.InnerException) {
        $exception = $exception.InnerException
    }
    return [string]$exception.Message
}

function Get-ComHResult($ErrorRecord) {
    $exception = $ErrorRecord.Exception
    while ($null -ne $exception) {
        if ($exception -is [System.Runtime.InteropServices.COMException]) {
            return $exception.HResult
        }
        $exception = $exception.InnerException
    }
    return 0
}

function Invoke-Excel([scriptblock]$Call) {
    # Excel rejects a call while it is busy (RPC_E_CALL_REJECTED or
    # RPC_E_SERVERCALL_RETRYLATER). A rejected call never started, so it is
    # safe to repeat.
    for ($attempt = 0; ; $attempt++) {
        try {
            return (& $Call)
        }
        catch {
            $code = Get-ComHResult $_
            if (($code -eq -2147418111 -or $code -eq -2147417846) -and $attempt -lt 120) {
                Start-Sleep -Milliseconds 500
                continue
            }
            throw
        }
    }
}

function Invoke-Native([string]$FilePath, [string[]]$Arguments, [string]$WorkingDirectory) {
    # Runs a program and returns its exit code, its standard output as bytes
    # and its standard error as text, without PowerShell re-encoding either.
    $info = New-Object System.Diagnostics.ProcessStartInfo
    $command = Get-Command $FilePath -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $command) {
        throw "$FilePath was not found on the PATH."
    }
    $info.FileName = $command.Source
    $quoted = foreach ($argument in $Arguments) {
        if ($argument -match '[\s"]') { '"' + ($argument -replace '"', '\"') + '"' } else { $argument }
    }
    $info.Arguments = ($quoted -join ' ')
    $info.WorkingDirectory = $WorkingDirectory
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.CreateNoWindow = $true
    $process = [System.Diagnostics.Process]::Start($info)
    $errorText = $process.StandardError.ReadToEndAsync()
    $buffer = New-Object System.IO.MemoryStream
    $process.StandardOutput.BaseStream.CopyTo($buffer)
    $process.WaitForExit()
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        Output   = $buffer.ToArray()
        Error    = [string]$errorText.Result
    }
}

function Invoke-Git([string[]]$Arguments) {
    $result = Invoke-Native -FilePath 'git' -Arguments $Arguments -WorkingDirectory $script:RepoRoot
    if ($result.ExitCode -ne 0) {
        throw ('git {0} failed: {1}' -f ($Arguments -join ' '), $result.Error.Trim())
    }
    # The unary comma returns the byte array as one object. Without it
    # PowerShell enumerates the array, so an empty output (the status of a
    # clean checkout) reaches the caller as $null and other output as Object[].
    return ,$result.Output
}

function Get-Sha256Hex([byte[]]$Bytes) {
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (-join ($algorithm.ComputeHash($Bytes) | ForEach-Object { $_.ToString('x2') }))
    }
    finally {
        $algorithm.Dispose()
    }
}

function Get-FileSha256([string]$Path) {
    return Get-Sha256Hex ([System.IO.File]::ReadAllBytes($Path))
}

function Write-JsonFile([string]$Name, $Value) {
    $json = ConvertTo-Json -InputObject $Value -Depth 10
    [System.IO.File]::WriteAllText((Join-Path $script:BundleDir $Name), $json + "`n", $Utf8NoBom)
}

function Get-RegistryValue([string]$Path, [string]$Name) {
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }
    return $item.$Name
}

function Get-OfficeVersion {
    # Excel.Application.16 -> 16.0, read without starting Excel.
    $current = Get-RegistryValue -Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer' -Name '(default)'
    if ($current -match '\.(\d+)$') {
        return ('{0}.0' -f $Matches[1])
    }
    return '16.0'
}

function Get-SecurityValue([string]$OfficeVersion, [string]$Name) {
    # A policy value takes precedence over the user's own setting.
    foreach ($root in @('HKCU:\Software\Policies\Microsoft\Office', 'HKLM:\Software\Policies\Microsoft\Office',
                        'HKCU:\Software\Microsoft\Office')) {
        $value = Get-RegistryValue -Path "$root\$OfficeVersion\Excel\Security" -Name $Name
        if ($null -ne $value) {
            return $value
        }
    }
    return $null
}

function Get-MacroSetting([string]$OfficeVersion) {
    switch (Get-SecurityValue -OfficeVersion $OfficeVersion -Name 'VBAWarnings') {
        1 { return 'Enable VBA macros' }
        2 { return 'Disable VBA macros with notification' }
        3 { return 'Disable VBA macros except digitally signed macros' }
        4 { return 'Disable VBA macros without notification' }
        default { return 'Disable VBA macros with notification (default; no setting found)' }
    }
}

function Get-VbaProjectAccess([string]$OfficeVersion) {
    if ((Get-SecurityValue -OfficeVersion $OfficeVersion -Name 'AccessVBOM') -eq 1) {
        return 'On'
    }
    return 'Off'
}

function Get-WindowsInfo {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $currentVersion = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $build = [string]$os.BuildNumber
    $revision = Get-RegistryValue -Path $currentVersion -Name 'UBR'
    if ($null -ne $revision) {
        $build = '{0}.{1}' -f $build, $revision
    }
    $displayVersion = Get-RegistryValue -Path $currentVersion -Name 'DisplayVersion'
    return [ordered]@{
        caption         = [string]$os.Caption
        version         = [string]$os.Version
        build           = $build
        display_version = $(if ($null -ne $displayVersion) { [string]$displayVersion } else { $null })
        architecture    = [string]$os.OSArchitecture
    }
}

function Initialize-Native {
    if (-not ('ExcelUiCertificationNative' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelUiCertificationNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWow64Process(IntPtr process, [MarshalAs(UnmanagedType.Bool)] out bool wow64);
}
'@
    }
}

function Get-ExcelProcessId($Excel) {
    try {
        Initialize-Native
        [uint32]$processId = 0
        [void][ExcelUiCertificationNative]::GetWindowThreadProcessId([IntPtr][int64]$Excel.Hwnd, [ref]$processId)
        if ($processId -ne 0) {
            return [int]$processId
        }
    }
    catch {
        Write-Log ('The Excel process could not be identified: ' + (Get-InnerMessage $_))
    }
    return $null
}

function Get-ExcelBitness($ProcessId, [string]$OperatingSystem) {
    # The bitness of the Excel process itself; Application.OperatingSystem
    # ("Windows (64-bit) NT 10.00") is the fallback.
    if ($null -ne $ProcessId) {
        try {
            if (-not [Environment]::Is64BitOperatingSystem) {
                return @('x86', 'process')
            }
            $process = [System.Diagnostics.Process]::GetProcessById($ProcessId)
            $wow64 = $false
            if ([ExcelUiCertificationNative]::IsWow64Process($process.Handle, [ref]$wow64)) {
                if ($wow64) { return @('x86', 'process') } else { return @('x64', 'process') }
            }
        }
        catch {
            Write-Log ('The Excel process bitness could not be read: ' + (Get-InnerMessage $_))
        }
    }
    if ($OperatingSystem -match '\(64-bit\)') { return @('x64', 'Application.OperatingSystem') }
    if ($OperatingSystem -match '\(32-bit\)') { return @('x86', 'Application.OperatingSystem') }
    return @('unknown', 'none')
}

function Get-SourceDigests([string]$Sha) {
    # The development modules of the commit, with SHA-256 of each as a
    # checkout writes it (.gitattributes: *.bas eol=crlf), and of the file the
    # build imported from the working tree.
    $manifestText = $Utf8NoBom.GetString((Invoke-Git @('show', "${Sha}:tools/vba-projects.json")))
    $manifest = $manifestText | ConvertFrom-Json
    $modules = New-Object System.Collections.ArrayList
    foreach ($path in @($manifest.configurations.development.modules)) {
        $text = $Latin1.GetString((Invoke-Git @('show', "${Sha}:$path")))
        $checkedOut = $Latin1.GetBytes([regex]::Replace($text, "(?<!`r)`n", "`r`n"))
        $name = [regex]::Match($text, '(?m)^Attribute VB_Name = "([A-Za-z]\w*)"\r?$').Groups[1].Value
        $file = Join-Path $script:RepoRoot ($path -replace '/', '\')
        $fileDigest = $null
        if (Test-Path -LiteralPath $file -PathType Leaf) {
            $fileDigest = Get-FileSha256 $file
        }
        [void]$modules.Add([ordered]@{
            path               = [string]$path
            name               = [string]$name
            sha256             = Get-Sha256Hex $checkedOut
            checked_out_sha256 = $fileDigest
        })
    }
    return [ordered]@{
        schema        = 1
        commit        = $Sha
        configuration = 'development'
        line_endings  = 'CRLF'
        modules       = $modules
    }
}

function Get-TempFolders {
    $folders = New-Object System.Collections.ArrayList
    $candidates = @($env:TEMP, $env:TMP,
                    [Environment]::GetEnvironmentVariable('TEMP', 'User'),
                    [Environment]::GetEnvironmentVariable('TMP', 'User'),
                    [System.IO.Path]::GetTempPath())
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }
        $expanded = [Environment]::ExpandEnvironmentVariables($candidate)
        if (Test-Path -LiteralPath $expanded -PathType Container) {
            $full = (Get-Item -LiteralPath $expanded).FullName.TrimEnd('\')
            if (-not ($folders -contains $full)) {
                [void]$folders.Add($full)
            }
        }
    }
    return $folders
}

function Find-EvidenceFiles([datetime]$Since) {
    # The certification files written since $Since, by name, in any folder
    # Excel may use as %TEMP%. The same file seen through two folder names
    # counts once.
    $found = @{}
    foreach ($folder in @(Get-TempFolders)) {
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -File -Filter 'EXCEL_UI_certification_*' -ErrorAction SilentlyContinue)) {
            $match = [regex]::Match($file.Name, $EvidenceNamePattern)
            if (-not $match.Success -or $file.LastWriteTime -lt $Since) {
                continue
            }
            $named = [datetime]::ParseExact($match.Groups[1].Value, 'yyyyMMdd_HHmmss', $Invariant)
            if ($named -lt $Since) {
                continue
            }
            $found[$file.Name.ToLowerInvariant()] = $file
        }
    }
    return @($found.Values)
}

function Find-Python {
    $candidates = @()
    if ($Python) {
        $candidates += , @($Python)
    }
    else {
        $candidates += , @('py', '-3')
        $candidates += , @('python')
        $candidates += , @('python3')
    }
    foreach ($candidate in $candidates) {
        if (-not (Get-Command $candidate[0] -CommandType Application -ErrorAction SilentlyContinue)) {
            continue
        }
        $extra = @($candidate | Select-Object -Skip 1)
        $probe = Invoke-Native -FilePath $candidate[0] `
            -Arguments ($extra + @('-c', 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)')) `
            -WorkingDirectory $script:RepoRoot
        if ($probe.ExitCode -eq 0) {
            return [pscustomobject]@{ Exe = $candidate[0]; Arguments = $extra }
        }
    }
    return $null
}

function Release-Com($Object) {
    if ($null -ne $Object) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }
        catch {
        }
    }
}

# ---------------------------------------------------------------------------
# Repository, commit and bundle
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent $PSScriptRoot
}
$script:RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path.TrimEnd('\')

try {
    $sha = $Utf8NoBom.GetString((Invoke-Git @('rev-parse', '--verify', 'HEAD^{commit}'))).Trim()
    $statusText = $Utf8NoBom.GetString((Invoke-Git @('status', '--porcelain', '--untracked-files=no')))
}
catch {
    Write-Host ('ERROR: ' + (Get-InnerMessage $_))
    exit 1
}
if ($sha -notmatch '^[0-9a-f]{40}$') {
    Write-Host "ERROR: cannot read the checked-out commit of $script:RepoRoot."
    exit 1
}
$dirtyPaths = @($statusText -split "`n" | ForEach-Object { $_.TrimEnd("`r") } | Where-Object { $_ -ne '' })
if ($dirtyPaths.Count -gt 0 -and -not $AllowDirty) {
    Write-Host 'ERROR: the working tree has changes. Certify a clean checkout (git status shows none):'
    $dirtyPaths | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$started = Get-Date
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $script:RepoRoot) `
        ('EXCEL_UI-certification-{0}-{1}' -f $sha.Substring(0, 12), $started.ToString('yyyyMMdd-HHmmss', $Invariant))
}
$script:BundleDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BundleDir).TrimEnd('\')
if (($script:BundleDir + '\').StartsWith($script:RepoRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Host "ERROR: the bundle folder must be outside the checkout: $script:BundleDir"
    exit 1
}
if (Test-Path -LiteralPath $script:BundleDir) {
    if (@(Get-ChildItem -LiteralPath $script:BundleDir -Force).Count -gt 0) {
        Write-Host "ERROR: the bundle folder is not empty: $script:BundleDir. A bundle is never overwritten."
        exit 1
    }
}
else {
    [void](New-Item -ItemType Directory -Path $script:BundleDir)
}
$script:SessionLog = Join-Path $script:BundleDir 'session.log'
$workbookPath = Join-Path $script:BundleDir $WorkbookFileName

$script:Session = [ordered]@{
    schema         = 1
    tool           = 'tools/Certify-EXCEL_UI.ps1'
    commit         = $sha
    dirty          = ($dirtyPaths.Count -gt 0)
    allow_dirty    = [bool]$AllowDirty
    dirty_paths    = @($dirtyPaths | ForEach-Object { [string]$_ })
    started_utc    = Get-UtcStamp $started
    started_local  = Get-LocalStamp $started
    finished_utc   = $null
    finished_local = $null
    steps          = @($StepNames | ForEach-Object {
            [ordered]@{ name = [string]$_; status = 'NOT_RUN'; detail = ''; at_utc = $null }
        })
    workbook       = $null
    components     = @()
    evidence       = @()
    result         = 'FAIL'
    exit_code      = 1
    message        = ''
}

Write-Log "VBA Excel UI certification of commit $sha"
Write-Log "Bundle: $script:BundleDir"
if ($script:Session.dirty) {
    Write-Log 'WARNING: -AllowDirty: the working tree has changes. This bundle cannot certify the commit.'
}

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------
$excel = $null
$workbook = $null
$vbe = $null
$excelProcessId = $null
$cancelled = $false
$certificationPassed = $false

try {
    # --- preflight ----------------------------------------------------------
    Set-Step 'preflight' 'RUNNING' 'Reading Windows, the Trust Center settings and the commit source'
    $windows = Get-WindowsInfo
    $officeVersion = Get-OfficeVersion
    $macroSetting = Get-MacroSetting $officeVersion
    $vbaProjectAccess = Get-VbaProjectAccess $officeVersion
    Write-Log ('Windows: {0} {1} build {2} ({3})' -f $windows.caption, $windows.version, $windows.build, $windows.architecture)
    Write-Log "Macro setting: $macroSetting; Trust access to the VBA project object model: $vbaProjectAccess"
    if ($vbaProjectAccess -ne 'On') {
        throw ("'Trust access to the VBA project object model' is off. Building the workbook needs it: " +
               'turn it on in File > Options > Trust Center > Trust Center Settings > Macro Settings, ' +
               'if your policy allows. This script never changes it.')
    }
    $syncScript = Join-Path $script:RepoRoot 'tools\Sync-EXCEL_UI-VBA.ps1'
    if (-not (Test-Path -LiteralPath $syncScript -PathType Leaf)) {
        throw "The synchronizer is missing: $syncScript"
    }
    Write-JsonFile 'source-digests.json' (Get-SourceDigests $sha)
    Set-Step 'preflight' 'PASS' "Commit $sha; source-digests.json written"

    # --- build --------------------------------------------------------------
    Set-Step 'build' 'RUNNING' 'tools\Sync-EXCEL_UI-VBA.ps1 -Mode Development -NewWorkbook'
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript `
            -WorkbookPath $workbookPath -Mode Development -NewWorkbook -Force -NoBackup 2>&1 |
            ForEach-Object { Write-Log ('  sync: ' + [string]$_) }
        $syncCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($syncCode -eq 2) {
        Stop-Cancelled 'The synchronizer was cancelled.'
    }
    if ($syncCode -ne 0 -or -not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
        throw "The synchronizer exited with code $syncCode; see session.log."
    }
    $script:Session.workbook = [ordered]@{ file = $WorkbookFileName; sha256 = Get-FileSha256 $workbookPath }
    Set-Step 'build' 'PASS' "$WorkbookFileName built from the checkout"

    # --- open ---------------------------------------------------------------
    Set-Step 'open' 'RUNNING' 'Starting Excel'
    $excel = New-Object -ComObject Excel.Application
    if ((Invoke-Excel { $excel.Workbooks.Count }) -ne 0) {
        # Never act on an instance this script did not start.
        $attached = $excel
        $excel = $null
        Release-Com $attached
        throw 'Excel returned an instance that already has workbooks open; close Excel and run again.'
    }
    $excelProcessId = Get-ExcelProcessId $excel
    Invoke-Excel { $excel.Visible = $true }
    Invoke-Excel { $excel.AutomationSecurity = $MsoAutomationSecurityByUI }
    $workbook = Invoke-Excel { $excel.Workbooks.Open($workbookPath) }
    $operatingSystem = [string](Invoke-Excel { $excel.OperatingSystem })
    $bitness = Get-ExcelBitness $excelProcessId $operatingSystem
    $clickToRun = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $hostRecord = [ordered]@{
        schema       = 1
        windows      = $windows
        excel        = [ordered]@{
            version               = [string](Invoke-Excel { $excel.Version })
            build                 = [string](Invoke-Excel { $excel.Build })
            operating_system      = $operatingSystem
            bitness               = [string]$bitness[0]
            bitness_source        = [string]$bitness[1]
            click_to_run_version  = Get-RegistryValue -Path $clickToRun -Name 'VersionToReport'
            click_to_run_platform = Get-RegistryValue -Path $clickToRun -Name 'Platform'
        }
        trust_center = [ordered]@{
            macro_setting       = $macroSetting
            vba_project_access  = $vbaProjectAccess
            automation_security = 'msoAutomationSecurityByUI'
        }
    }
    Write-JsonFile 'host.json' $hostRecord
    Set-Step 'open' 'PASS' ('Excel {0} build {1}, {2}; process {3}' -f $hostRecord.excel.version, $hostRecord.excel.build,
                            $hostRecord.excel.bitness, $excelProcessId)

    # --- enable-macros ------------------------------------------------------
    Set-Step 'enable-macros' 'RUNNING' 'Waiting for the workbook macros to run'
    $qualified = "'" + $workbook.Name.Replace("'", "''") + "'!"
    $ready = $false
    while (-not $ready) {
        try {
            [void](Invoke-Excel { $excel.Run($qualified + $ReadinessMacro) })
            $ready = $true
        }
        catch {
            Write-Host ''
            Write-Host 'The workbook macros do not run. In the Excel window, click Enable Content in the'
            Write-Host 'message bar (or File > Info > Enable Content), then press Enter here.'
            Write-Host 'If Excel shows a compile error instead, close the message and type F.'
            $answer = Read-Host 'Enter to try again, F for a compile error, Q to stop'
            if ($answer -match '^\s*q') {
                Stop-Cancelled 'The operator stopped: the macros were not enabled.'
            }
            if ($answer -match '^\s*f') {
                throw 'The project does not compile: Excel reported a compile error.'
            }
        }
    }
    Set-Step 'enable-macros' 'PASS' 'The workbook macros run'

    # --- compile ------------------------------------------------------------
    # Debug > Compile VBAProject acts on the active project and is greyed out
    # once that project is compiled, so the VBE is shown and its state read.
    Set-Step 'compile' 'RUNNING' 'Debug > Compile VBAProject'
    $vbe = Invoke-Excel { $excel.VBE }
    Invoke-Excel { $vbe.MainWindow.Visible = $true }
    try {
        Invoke-Excel { $vbe.ActiveVBProject = $workbook.VBProject }
    }
    catch {
        Write-Log ('The workbook project could not be made the active one: ' + (Get-InnerMessage $_))
    }
    $control = Invoke-Excel { $vbe.CommandBars.FindControl(1, $CompileControlId) }
    if ($null -eq $control) {
        throw 'Debug > Compile VBAProject was not found in the VBE.'
    }
    if (Invoke-Excel { $control.Enabled }) {
        Invoke-Excel { $control.Execute() }
    }
    Release-Com $control
    $compiled = $false
    for ($attempt = 0; $attempt -lt 10 -and -not $compiled; $attempt++) {
        Start-Sleep -Milliseconds 500
        $control = Invoke-Excel { $vbe.CommandBars.FindControl(1, $CompileControlId) }
        $compiled = -not (Invoke-Excel { $control.Enabled })
        Release-Com $control
    }
    $compileDetail = 'Debug > Compile VBAProject completed and is greyed out'
    if (-not $compiled) {
        # The command can stay enabled without an error; the operator looks.
        Write-Host ''
        Write-Host 'Debug > Compile VBAProject is still enabled. In the VBE, choose Debug > Compile'
        Write-Host 'VBAProject yourself. Type Y if it completes with no error, N if it reports one.'
        $answer = Read-Host 'Compiled with no error? (Y/N, Q to stop)'
        if ($answer -match '^\s*q') {
            Stop-Cancelled 'The operator stopped at the compile step.'
        }
        if ($answer -notmatch '^\s*y') {
            throw 'The project does not compile: the VBE reported an error.'
        }
        $compileDetail = 'Debug > Compile VBAProject completed with no error, as the operator saw in the VBE'
    }
    $components = New-Object System.Collections.ArrayList
    foreach ($component in @(Invoke-Excel { $workbook.VBProject.VBComponents })) {
        [void]$components.Add([ordered]@{ name = [string]$component.Name; type = [int]$component.Type })
        Release-Com $component
    }
    $script:Session.components = $components
    Invoke-Excel { $vbe.MainWindow.Visible = $false }
    Invoke-Excel { $workbook.Activate() }
    Set-Step 'compile' 'PASS' $compileDetail

    # --- certify ------------------------------------------------------------
    Set-Step 'certify' 'RUNNING' "$CertificationMacro (this takes a while; Excel opens and closes windows)"
    Write-Host 'Do not use Excel until the run ends. If Excel shows a run-time error message,'
    Write-Host 'choose End: the evidence is written before the error is raised.'
    $runStart = Get-Date
    $runStart = $runStart.AddTicks(-($runStart.Ticks % [TimeSpan]::TicksPerSecond))
    try {
        [void](Invoke-Excel { $excel.Run($qualified + $CertificationMacro) })
        Set-Step 'certify' 'PASS' "$CertificationMacro returned without raising"
    }
    catch {
        Set-Step 'certify' 'FAIL' ("$CertificationMacro raised: " + (Get-InnerMessage $_))
    }

    # --- collect ------------------------------------------------------------
    Set-Step 'collect' 'RUNNING' 'Collecting the files the run wrote to the temporary folder'
    $files = @(Find-EvidenceFiles $runStart)
    $json = @($files | Where-Object { $_.Extension -eq '.json' })
    $text = @($files | Where-Object { $_.Extension -eq '.txt' })
    if ($json.Count -ne 1 -or $text.Count -ne 1) {
        $names = ($files | ForEach-Object { $_.FullName }) -join ', '
        if ($names -eq '') { $names = 'none' }
        throw ('Expected one certification JSON and one report written since {0}; found {1} and {2}: {3}. Searched: {4}' -f
               $runStart.ToString('yyyy-MM-dd HH:mm:ss', $Invariant), $json.Count, $text.Count, $names,
               (@(Get-TempFolders) -join ', '))
    }
    $evidence = New-Object System.Collections.ArrayList
    foreach ($pair in @(@($json[0], 'certification.json'), @($text[0], 'certification.txt'))) {
        $target = Join-Path $script:BundleDir $pair[1]
        Copy-Item -LiteralPath $pair[0].FullName -Destination $target
        [void]$evidence.Add([ordered]@{ file = $pair[1]; source_name = [string]$pair[0].Name; sha256 = Get-FileSha256 $target })
        Write-Log ('Collected {0} as {1}' -f $pair[0].FullName, $pair[1])
    }
    $script:Session.evidence = $evidence
    $result = Get-Content -LiteralPath (Join-Path $script:BundleDir 'certification.json') -Raw | ConvertFrom-Json
    $certificationPassed = ($result.passed -eq $true -and $result.complete -eq $true)
    $line = @(Select-String -LiteralPath (Join-Path $script:BundleDir 'certification.txt') -Pattern '^RESULT: ')
    $verdict = $(if ($line.Count -gt 0) { $line[0].Line } else { 'no RESULT line in the report' })
    Set-Step 'collect' 'PASS' $verdict
}
catch [System.OperationCanceledException] {
    $cancelled = $true
    $script:Session.message = [string]$_.Exception.Message
    Set-Step $script:CurrentStep 'CANCELLED' $script:Session.message
}
catch {
    $script:Session.message = Get-InnerMessage $_
    Set-Step $script:CurrentStep 'FAIL' $script:Session.message
}
finally {
    # --- close --------------------------------------------------------------
    if ($null -ne $excel) {
        Set-Step 'close' 'RUNNING' 'Closing the workbook and the Excel instance this script started'
        $closeProblems = @()
        if ($null -ne $vbe) {
            try { $vbe.MainWindow.Visible = $false } catch { }
            Release-Com $vbe
        }
        if ($null -ne $workbook) {
            try {
                Invoke-Excel { $workbook.Close($false) }
            }
            catch {
                $closeProblems += 'the workbook did not close: ' + (Get-InnerMessage $_)
            }
            Release-Com $workbook
        }
        try {
            # Workbooks a failed run left open are closed without saving.
            Invoke-Excel { $excel.DisplayAlerts = $false }
            Invoke-Excel { $excel.Quit() }
        }
        catch {
            $closeProblems += 'Excel did not quit: ' + (Get-InnerMessage $_)
        }
        Release-Com $excel
        $excel = $null
        $workbook = $null
        $vbe = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        $closeDetail = 'Workbook closed without saving; Excel quit'
        if ($null -ne $excelProcessId) {
            $process = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
            if ($null -ne $process -and -not $process.WaitForExit(30000)) {
                $closeDetail += "; process $excelProcessId is still running, close it yourself"
            }
        }
        if ($closeProblems.Count -gt 0) {
            Set-Step 'close' 'FAIL' ($closeProblems -join '; ')
        }
        else {
            Set-Step 'close' 'PASS' $closeDetail
        }
    }
}

# ---------------------------------------------------------------------------
# Record and validate
# ---------------------------------------------------------------------------
$finished = Get-Date
$script:Session.finished_utc = Get-UtcStamp $finished
$script:Session.finished_local = Get-LocalStamp $finished
$allPassed = (@($script:Session.steps | Where-Object { $_.status -ne 'PASS' }).Count -eq 0)
if ($cancelled) {
    $script:Session.result = 'CANCELLED'
    $script:Session.exit_code = 2
}
elseif ($allPassed -and $certificationPassed) {
    $script:Session.result = 'PASS'
    $script:Session.exit_code = 0
}
else {
    $script:Session.result = 'FAIL'
    $script:Session.exit_code = 1
}
Write-JsonFile 'session.json' $script:Session
Write-Log ('Run finished: {0}' -f $script:Session.result)

$exitCode = [int]$script:Session.exit_code
if (-not $cancelled) {
    Write-Host ''
    $pythonCommand = Find-Python
    $validator = Join-Path $script:RepoRoot 'tools\check_excel_evidence.py'
    if ($null -eq $pythonCommand) {
        Write-Host 'Python 3.11 or later was not found, so the bundle was not validated here. On any'
        Write-Host 'machine with a clone of this repository, run:'
        Write-Host "  python3 tools/check_excel_evidence.py <bundle> --commit $sha"
    }
    else {
        $arguments = @($pythonCommand.Arguments) + @($validator, $script:BundleDir, '--commit', $sha, '--root', $script:RepoRoot)
        & $pythonCommand.Exe @arguments
        if ($LASTEXITCODE -ne 0 -and $exitCode -eq 0) {
            $exitCode = 1
        }
    }
}

Write-Host ''
Write-Host "Bundle: $script:BundleDir"
switch ($exitCode) {
    0 { Write-Host "RESULT: PASS - commit $sha is certified on this host." }
    2 { Write-Host 'RESULT: CANCELLED - keep the bundle only as a record of the attempt.' }
    default { Write-Host 'RESULT: FAIL - see the steps above and session.log. Keep the bundle as it is.' }
}
exit $exitCode
