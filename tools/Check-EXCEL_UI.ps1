<#
.SYNOPSIS
Run certification, the regression runners and the demo checks in Excel with
one command, and write one result bundle with a summary.

.DESCRIPTION
For the commit checked out in this repository:

1. preflight: reads the commit SHA and refuses a working tree with changes;
   creates a new bundle folder next to the checkout; reads Windows and the
   Trust Center settings without changing them, and stops when "Trust access
   to the VBA project object model" is off;
2. certification: runs tools\Certify-EXCEL_UI.ps1 unchanged, with its bundle
   in the certification folder of this bundle;
3. open-development: opens the development workbook of that bundle in an
   Excel instance of its own, which follows the Trust Center macro setting;
4. one step per runner: runs it through TST_RunRunnerWithLog, which writes
   what the runner prints to runner-<name>.log and returns its verdict;
5. close-development: closes that workbook without saving and quits the
   Excel instance;
6. demo-build: creates EXCEL_UI-demo.xlsm with
   tools\Sync-EXCEL_UI-VBA.ps1 -Mode Demo -NewWorkbook;
7. demo-open and demo-compile: opens it in another Excel instance of its own,
   waits for Enable Content and runs Debug > Compile VBAProject;
8. demo-cleanup-checks: runs Demo_Test_RunCleanupChecks. A scenario the host
   could not arrange, because the default printer driver cannot lay out a
   sheet, is skipped, not failed. The step is then SKIP, which is not a pass;
9. demo-create-sheet-1 and demo-create-sheet-2: time two runs of
   Demo_CreateDemoSheet, and warn above -SlowBuildSeconds;
10. default-printer: records the name and status of the Windows default
    printer;
11. demo-titlebar: selects only chkTitleBar on the demo sheet, runs the HIDE
    and SHOW actions, and reads the title bar back after each;
12. close-demo: closes the demo workbook without saving and quits that Excel
    instance.

A step that fails does not stop the steps that do not depend on it. The
script writes summary.txt, with one line per step and a final RESULT line,
and session.json, and prints the summary at the end. RESULT is PASS when every
step passed, warned or was recorded as INFO; FAIL when a step failed; and
INCOMPLETE when no step failed but a step is SKIP, so that part of what it
checks did not run.

Functions and settings it shares with tools\Certify-EXCEL_UI.ps1 are read
from that script's syntax tree, without running its body, so both scripts
open, run and close Excel the same way.

Exit codes: 0 every step passed, warned or was recorded as INFO; 1 a step
failed or the run stopped on an error; 2 the operator cancelled it; 3 no step
failed but a step was SKIP (RESULT: INCOMPLETE).

Unattended: with -Unattended the whole run needs no key press and no click,
from the start to the final RESULT line. The switch goes to
tools\Certify-EXCEL_UI.ps1 too. Every Excel instance this script (or that
script) starts is set to msoAutomationSecurityLow before it opens its workbook,
so the macros of the workbooks built from the checked-out commit run without
the Enable Content message; no other instance, the Trust Center, the registry
and the Trusted Locations are left as they are. session.json records
automation_security as "msoAutomationSecurityLow (-Unattended)". Each wait is
bounded: a workbook whose macros do not answer within 60 seconds, or a compile
command that stays enabled for 30 seconds, fails its step with a message that
says why. Without the switch the script asks, as described above.

See docs\EXCEL_CHECKS.md.

.PARAMETER BundleDir
The bundle folder. It must be outside the checkout and new or empty. By
default, EXCEL_UI-checks-<SHA12>-<yyyyMMdd-HHmmss> next to the checkout.

.PARAMETER Runners
The runners to run, in order, instead of the default list. Names may be
separated by commas. Test_EXCEL_UI_RunRibbonSdiProbe is recorded as INFO.

.PARAMETER IncludeRibbonProbe
Also run Test_EXCEL_UI_RunRibbonSdiProbe, after the other runners. It asserts
nothing, so its step is INFO whatever it reports.

.PARAMETER SlowBuildSeconds
A Demo_CreateDemoSheet run that takes longer than this many seconds is a
WARN, not a FAIL. Default 10.

.PARAMETER Python
The Python 3.11+ executable that tools\Certify-EXCEL_UI.ps1 validates its
bundle with. By default that script looks for one.

.PARAMETER Unattended
Run without any prompt, as described above. It is passed to
tools\Certify-EXCEL_UI.ps1. Without it the script asks you to click Enable
Content and to confirm a compile that Excel does not report as finished.

.PARAMETER RepoRoot
The repository. By default, the parent of this script's folder.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Check-EXCEL_UI.ps1

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Check-EXCEL_UI.ps1 -Runners Test_EXCEL_UI_RunCore -SlowBuildSeconds 20
#>
[CmdletBinding()]
param(
    [string]$BundleDir,
    [string[]]$Runners,
    [switch]$IncludeRibbonProbe,
    [ValidateRange(0.1, 3600)]
    [double]$SlowBuildSeconds = 10,
    [string]$Python,
    [string]$RepoRoot,
    [switch]$Unattended
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Latin1 = [System.Text.Encoding]::GetEncoding(28591)

# The runners TST_RunRunnerWithLog accepts, in the order they run by default.
$DefaultRunners = @(
    'Test_EXCEL_UI_RunAll',
    'Test_EXCEL_UI_RunCore',
    'Test_EXCEL_UI_RunTitleBarOnly',
    'Test_EXCEL_UI_RunSnapshotIdentity',
    'Test_EXCEL_UI_RunTitleBarSdiIdentity',
    'Test_EXCEL_UI_RunCertificationSelfTest',
    'Test_EXCEL_UI_RunOwnershipCleanupChecks'
)
# A characterization that asserts nothing: run only when asked, recorded as INFO.
$ProbeRunner = 'Test_EXCEL_UI_RunRibbonSdiProbe'
$ActiveFrameSettleSeconds = 10
$RunnerMacro = 'TST_RunRunnerWithLog'

# The demo macros, run through Demo_Test_RunLogged of M_DEMO_CLEANUP_TESTS.
$DemoLoggedMacro = 'Demo_Test_RunLogged'
$DemoTitleBarStateMacro = 'Demo_Test_TitleBarState'
$DemoCleanupMacro = 'Demo_Test_RunCleanupChecks'
$DemoBuildMacro = 'Demo_CreateDemoSheet'
$DemoClearMacro = 'Demo_ClearAllUI'
$DemoHideMacro = 'Demo_HideSelectedUI'
$DemoShowMacro = 'Demo_ShowSelectedUI'
# DEMO_SHEET_NAME and CB_TITLEBAR of demo/M_EXCEL_UI_DEMO.bas.
$DemoSheetName = 'DEMO_UI'
$TitleBarCheckBox = 'chkTitleBar'
$XlOn = 1

$DemoWorkbookFileName = 'EXCEL_UI-demo.xlsm'
$CertificationFolder = 'certification'
$CertifyScriptName = 'Certify-EXCEL_UI.ps1'

# What this script takes from tools\Certify-EXCEL_UI.ps1.
$SharedFunctions = @(
    'Get-UtcStamp', 'Get-LocalStamp', 'Write-Log', 'Stop-Cancelled', 'Get-InnerMessage', 'Get-ComHResult',
    'Invoke-Excel', 'Invoke-Native', 'Invoke-Git', 'Read-EvidencePolicy', 'Write-JsonFile', 'Get-RegistryValue',
    'Get-OfficeVersion', 'Get-SecurityValue', 'Get-MacroSetting', 'Get-VbaProjectAccess', 'Get-WindowsInfo',
    'Initialize-Native', 'Get-ExcelProcessId', 'Get-ExcelBitness', 'Stop-QuitProcess', 'Release-Com',
    'Get-AutomationSecurityLabel', 'Set-ExcelAutomationSecurity', 'Read-OperatorAnswer', 'Wait-ExcelMacros',
    'Wait-CompileCommand', 'Resolve-CompileFallback'
)
$SharedValues = @('MsoAutomationSecurityByUI', 'MsoAutomationSecurityLow', 'CompileControlId', 'ReadinessMacro',
                  'MacrosWaitSeconds', 'MacrosPollSeconds', 'CompileWaitSeconds')

$script:SessionLog = $null
$script:Steps = @()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Import-CertifyDefinitions {
    # Returns the named function definitions of the certification script as
    # one script block, and the constant values of the named script
    # variables. Nothing in that script runs. A name it no longer defines
    # stops this script, so the two cannot drift apart silently.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$Functions,
        [string[]]$Variables = @()
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "The certification script is missing: $Path"
    }
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$parseErrors)
    if (@($parseErrors).Count -gt 0) {
        throw ('{0} does not parse: {1}' -f $Path, $parseErrors[0].Message)
    }
    $definitions = @{}
    $values = @{}
    foreach ($statement in $ast.EndBlock.Statements) {
        if ($statement -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
            $definitions[$statement.Name] = $statement.Extent.Text
        }
        elseif ($statement -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $statement.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $statement.Right -is [System.Management.Automation.Language.CommandExpressionAst] -and
                $statement.Right.Expression -is [System.Management.Automation.Language.ConstantExpressionAst]) {
            $values[$statement.Left.VariablePath.UserPath] = $statement.Right.Expression.Value
        }
    }
    $missing = @($Functions | Where-Object { -not $definitions.ContainsKey($_) }) +
               @($Variables | Where-Object { -not $values.ContainsKey($_) })
    if ($missing.Count -gt 0) {
        throw ('{0} no longer defines {1}.' -f $Path, ($missing -join ', '))
    }
    $selected = @{}
    foreach ($name in $Variables) {
        $selected[$name] = $values[$name]
    }
    return [pscustomobject]@{
        Definitions = [scriptblock]::Create((@($Functions | ForEach-Object { $definitions[$_] }) -join "`n`n"))
        Values      = $selected
    }
}

function Resolve-RunnerList {
    # The runners to run, in order and without repeats. Names may be given
    # as one comma-separated argument, as powershell.exe -File passes them.
    param(
        [string[]]$Requested,
        [switch]$IncludeProbe,
        [Parameter(Mandatory = $true)][string[]]$Default,
        [Parameter(Mandatory = $true)][string]$Probe
    )
    $names = @($Requested | ForEach-Object { ([string]$_) -split ',' } | ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' })
    if ($names.Count -eq 0) {
        $names = @($Default)
    }
    if ($IncludeProbe) {
        $names += $Probe
    }
    $known = @($Default) + $Probe
    $list = New-Object System.Collections.ArrayList
    foreach ($name in $names) {
        $match = @($known | Where-Object { $_ -ieq $name })
        if ($match.Count -eq 0) {
            throw ('Unknown runner {0}. Known runners: {1}.' -f $name, ($known -join ', '))
        }
        if (-not ($list -contains $match[0])) {
            [void]$list.Add($match[0])
        }
    }
    return , $list.ToArray()
}

function New-CheckStep([string]$Name, [string]$Log) {
    return [ordered]@{ name = $Name; status = 'NOT_RUN'; detail = ''; log = $Log; seconds = $null; at_utc = $null }
}

function Find-CheckStep([string]$Name) {
    foreach ($step in $script:Steps) {
        if ($step.name -eq $Name) {
            return $step
        }
    }
    throw "The step '$Name' is not in this run."
}

function Set-CheckStep {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$Detail = '',
        $Seconds = $null
    )
    $step = Find-CheckStep $Name
    $step.status = $Status
    $step.detail = $Detail
    $step.at_utc = Get-UtcStamp
    if ($null -ne $Seconds) {
        $step.seconds = [math]::Round([double]$Seconds, 2)
    }
    Write-Log ('[{0}] {1}: {2}' -f $Status, $Name, $Detail)
}

function Get-CheckStepStatus([string]$Name) {
    return [string](Find-CheckStep $Name).status
}

function Skip-CheckStep([string[]]$Names, [string]$Reason) {
    # A step that cannot run because one it depends on failed is a FAIL: the
    # run did not check what the step checks.
    foreach ($name in $Names) {
        if ((Get-CheckStepStatus $name) -eq 'NOT_RUN') {
            Set-CheckStep $name 'FAIL' ('not run: ' + $Reason)
        }
    }
}

function Register-StepError([string]$Name, $ErrorRecord) {
    # Records a step's error as its FAIL. A cancellation is passed on, so
    # that the run stops.
    if ($ErrorRecord.Exception -is [System.OperationCanceledException]) {
        throw $ErrorRecord
    }
    Set-CheckStep $Name 'FAIL' (Get-InnerMessage $ErrorRecord)
}

function ConvertFrom-Verdict([string]$Text) {
    # Reads the line TST_RunRunnerWithLog and Demo_Test_RunLogged return:
    # "RESULT: PASS", "RESULT: FAIL | ..." or, from Demo_Test_RunLogged only,
    # "RESULT: INCOMPLETE | ..." for a run that skipped scenarios and failed
    # none, and " | LOG WRITE FAILED: ..." when the log could not be written.
    # Only PASS is Passed; INCOMPLETE is not a pass.
    $marker = ' | LOG WRITE FAILED: '
    $line = [string]$Text
    $logFailure = ''
    $index = $line.IndexOf($marker, [System.StringComparison]::Ordinal)
    if ($index -ge 0) {
        $logFailure = $line.Substring($index + $marker.Length)
        $line = $line.Substring(0, $index)
    }
    if ($line -match '^RESULT: (PASS|FAIL|INCOMPLETE)( \| .*)?$') {
        return [pscustomobject]@{ HasVerdict = $true; Passed = ($Matches[1] -eq 'PASS')
                                  Incomplete = ($Matches[1] -eq 'INCOMPLETE'); ResultLine = $line
                                  LogFailure = $logFailure }
    }
    return [pscustomobject]@{ HasVerdict = $false; Passed = $false; Incomplete = $false; ResultLine = $line
                              LogFailure = $logFailure }
}

function Get-LastLogLine([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }
    $lines = @([System.IO.File]::ReadAllLines($Path, $Utf8NoBom) | Where-Object { $_.Trim() -ne '' })
    if ($lines.Count -eq 0) {
        return $null
    }
    return [string]$lines[-1]
}

function Get-MacroOutcome {
    # The outcome of one logged macro run: its verdict, and whether its log
    # holds the lines it printed and ends with that verdict.
    param(
        [string]$Returned,
        [string]$Raised,
        [string]$LogPath
    )
    if ($Raised) {
        return [pscustomobject]@{ Passed = $false; Incomplete = $false; Detail = 'Application.Run raised: ' + $Raised }
    }
    $verdict = ConvertFrom-Verdict $Returned
    if (-not $verdict.HasVerdict) {
        return [pscustomobject]@{ Passed = $false; Incomplete = $false; Detail = "no verdict was returned: '$Returned'" }
    }
    # A verdict that cannot be trusted is a failure, even an INCOMPLETE one.
    if ($verdict.LogFailure) {
        return [pscustomobject]@{ Passed = $false; Incomplete = $false
                                  Detail = $verdict.ResultLine + '; the log could not be written: ' + $verdict.LogFailure }
    }
    $last = Get-LastLogLine $LogPath
    if ($last -ne $verdict.ResultLine) {
        return [pscustomobject]@{ Passed = $false; Incomplete = $false
                                  Detail = $verdict.ResultLine + '; the log does not end with that verdict' }
    }
    return [pscustomobject]@{ Passed = $verdict.Passed; Incomplete = $verdict.Incomplete; Detail = $verdict.ResultLine }
}

function Get-OutcomeStatus {
    # The step status of a macro outcome: PASS for a pass, SKIP for a run that
    # skipped scenarios and failed none, otherwise FAIL. SKIP is not a pass.
    param([Parameter(Mandatory = $true)]$Outcome)
    if ($Outcome.Passed) {
        return 'PASS'
    }
    if ($Outcome.Incomplete) {
        return 'SKIP'
    }
    return 'FAIL'
}

function Get-TimingStatus([double]$Seconds, [double]$Threshold) {
    if ($Seconds -gt $Threshold) {
        return 'WARN'
    }
    return 'PASS'
}

function Format-Seconds([double]$Seconds) {
    return $Seconds.ToString('0.0', $Invariant)
}

function Get-RunResult {
    # PASS when every step passed, warned or was recorded as INFO. A step
    # still NOT_RUN counts as a failure unless the run was cancelled. A step
    # that is SKIP is not a pass: with no failure, the run is INCOMPLETE.
    param(
        [Parameter(Mandatory = $true)]$Steps,
        [bool]$Cancelled
    )
    if ($Cancelled) {
        return [pscustomobject]@{ Result = 'FAIL'; ExitCode = 2 }
    }
    $failed = @($Steps | Where-Object { @('PASS', 'WARN', 'INFO', 'SKIP') -notcontains $_.status })
    if ($failed.Count -gt 0) {
        return [pscustomobject]@{ Result = 'FAIL'; ExitCode = 1 }
    }
    $skipped = @($Steps | Where-Object { $_.status -eq 'SKIP' })
    if ($skipped.Count -gt 0) {
        return [pscustomobject]@{ Result = 'INCOMPLETE'; ExitCode = 3 }
    }
    return [pscustomobject]@{ Result = 'PASS'; ExitCode = 0 }
}

function Format-Summary {
    # summary.txt: the commit, the host, one line per step and the RESULT line.
    param(
        [Parameter(Mandatory = $true)][string]$Commit,
        [string]$HostLine,
        [Parameter(Mandatory = $true)]$Steps,
        [Parameter(Mandatory = $true)][string]$Result
    )
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("Commit: $Commit")
    if ($HostLine) {
        [void]$lines.Add("Host: $HostLine")
    }
    foreach ($step in $Steps) {
        $line = '{0} {1}: {2}' -f ([string]$step.status).PadRight(4), $step.name, $step.detail
        if ($step.log) {
            $line += ' [' + $step.log + ']'
        }
        [void]$lines.Add($line.TrimEnd())
    }
    [void]$lines.Add("RESULT: $Result")
    return , $lines.ToArray()
}

function ConvertTo-ArgumentLine([string[]]$Arguments) {
    # One command line, each argument quoted as Invoke-Native quotes it.
    $quoted = foreach ($argument in $Arguments) {
        if ($argument -eq '' -or $argument -match '[\s"]') { '"' + ($argument -replace '"', '\"') + '"' } else { $argument }
    }
    return ($quoted -join ' ')
}

function Invoke-Interactive([string]$FilePath, [string[]]$Arguments) {
    # Runs a program in this console, without redirecting its input or
    # output, so that its prompts reach the operator, and returns its exit
    # code.
    $process = Start-Process -FilePath $FilePath -ArgumentList (ConvertTo-ArgumentLine $Arguments) -NoNewWindow `
        -Wait -PassThru
    return [int]$process.ExitCode
}

function Add-LogLine([string]$Path, [string]$Text) {
    [System.IO.File]::AppendAllText($Path, $Text + "`r`n", $Utf8NoBom)
}

function Get-DefaultPrinter {
    # The Windows default printer as CIM reports it, or $null when there is
    # none. Excel lays sheets out through its driver, so a slow or offline
    # printer explains a slow demo build.
    $printer = @(Get-CimInstance -ClassName Win32_Printer -Filter 'Default = TRUE') | Select-Object -First 1
    if ($null -eq $printer) {
        return $null
    }
    $statusNames = @{ 1 = 'Other'; 2 = 'Unknown'; 3 = 'Idle'; 4 = 'Printing'; 5 = 'Warmup'; 6 = 'Stopped printing';
                      7 = 'Offline' }
    $status = $null
    $statusText = 'not reported'
    if ($null -ne $printer.PrinterStatus) {
        $status = [int]$printer.PrinterStatus
        $statusText = $(if ($statusNames.ContainsKey($status)) { $statusNames[$status] } else { 'code ' + $status })
    }
    return [ordered]@{
        name                = [string]$printer.Name
        work_offline        = [bool]$printer.WorkOffline
        network             = [bool]$printer.Network
        printer_status      = $status
        printer_status_text = $statusText
    }
}

function New-ExcelState([string]$Label) {
    return @{ Label = $Label; Excel = $null; Workbook = $null; Vbe = $null; ProcessId = $null; ProcessStart = $null
              Qualified = ''; Closed = $false }
}

function Open-ExcelWorkbook($State, [string]$Path) {
    # Starts an Excel instance of its own, as the certification script does,
    # and opens the workbook with the Trust Center macro setting.
    $State.Excel = New-Object -ComObject Excel.Application
    if ((Invoke-Excel { $State.Excel.Workbooks.Count }) -ne 0) {
        # Never act on an instance this script did not start.
        $attached = $State.Excel
        $State.Excel = $null
        Release-Com $attached
        throw 'Excel returned an instance that already has workbooks open; close Excel and run again.'
    }
    $State.ProcessId = Get-ExcelProcessId $State.Excel
    if ($null -ne $State.ProcessId) {
        $State.ProcessStart = (Get-Process -Id $State.ProcessId -ErrorAction SilentlyContinue).StartTime
    }
    Invoke-Excel { $State.Excel.Visible = $true }
    [void](Set-ExcelAutomationSecurity $State.Excel)
    $State.Workbook = Invoke-Excel { $State.Excel.Workbooks.Open($Path) }
    $State.Qualified = "'" + ([string]$State.Workbook.Name).Replace("'", "''") + "'!"
}

function Wait-WorkbookMacros($State) {
    # Waits until a read-only public function of the workbook runs, which it
    # does only once its macros are enabled, as the certification script does.
    Wait-ExcelMacros -Excel $State.Excel -Qualified $State.Qualified -WorkbookName ([string]$State.Workbook.Name)
}

function Invoke-ProjectCompile($State) {
    # Debug > Compile VBAProject acts on the active project and is greyed out
    # once that project is compiled, so the VBE is shown and its state read.
    $State.Vbe = Invoke-Excel { $State.Excel.VBE }
    Invoke-Excel { $State.Vbe.MainWindow.Visible = $true }
    try {
        Invoke-Excel { $State.Vbe.ActiveVBProject = $State.Workbook.VBProject }
    }
    catch {
        Write-Log ('The workbook project could not be made the active one: ' + (Get-InnerMessage $_))
    }
    $control = Invoke-Excel { $State.Vbe.CommandBars.FindControl(1, $CompileControlId) }
    if ($null -eq $control) {
        throw 'Debug > Compile VBAProject was not found in the VBE.'
    }
    if (Invoke-Excel { $control.Enabled }) {
        Invoke-Excel { $control.Execute() }
    }
    Release-Com $control
    $detail = 'Debug > Compile VBAProject completed and is greyed out'
    if (-not (Wait-CompileCommand $State.Vbe)) {
        # The command can stay enabled without an error; the operator looks,
        # or, with -Unattended, the step fails.
        $detail = Resolve-CompileFallback
    }
    Invoke-Excel { $State.Vbe.MainWindow.Visible = $false }
    Invoke-Excel { $State.Workbook.Activate() }
    return $detail
}

function Get-ExcelHost($State) {
    $operatingSystem = [string](Invoke-Excel { $State.Excel.OperatingSystem })
    $bitness = Get-ExcelBitness $State.ProcessId $operatingSystem
    return [ordered]@{
        version          = [string](Invoke-Excel { $State.Excel.Version })
        build            = [string](Invoke-Excel { $State.Excel.Build })
        operating_system = $operatingSystem
        bitness          = [string]$bitness[0]
    }
}

function Invoke-LoggedMacro {
    # Runs one macro through a logging entry point of the workbook and
    # returns its outcome and how long it took. The workbook is activated
    # first, as it is when you run a macro from its own window. A COM error,
    # such as a macro that is not found, is a failed outcome, not an exception.
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$EntryPoint,
        [Parameter(Mandatory = $true)][string]$Macro,
        [Parameter(Mandatory = $true)][string]$LogPath
    )
    $returned = ''
    $raised = ''
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        Invoke-Excel { $State.Workbook.Activate() }
        $returned = [string](Invoke-Excel { $State.Excel.Run($State.Qualified + $EntryPoint, $Macro, $LogPath) })
    }
    catch {
        $raised = Get-InnerMessage $_
    }
    $stopwatch.Stop()
    $outcome = Get-MacroOutcome -Returned $returned -Raised $raised -LogPath $LogPath
    return [pscustomobject]@{ Passed = $outcome.Passed; Incomplete = $outcome.Incomplete; Detail = $outcome.Detail
                              Seconds = $stopwatch.Elapsed.TotalSeconds }
}

function Wait-ActiveFrame($State) {
    # The identity runners refuse to start while Application.ActiveWindow and
    # Application.Hwnd name different frames. A runner that opens and closes
    # windows can leave the instance in that state, and you never meet it when
    # you start a runner from the Excel window itself. Bring the instance to
    # the foreground, activate the workbook and wait until both name one frame.
    # Returns an empty string when they agree, otherwise what was last read.
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic
        [void][Microsoft.VisualBasic.Interaction]::AppActivate([int]$State.ProcessId)
    }
    catch { }
    $deadline = (Get-Date).AddSeconds($ActiveFrameSettleSeconds)
    $lastRead = ''
    do {
        try {
            Invoke-Excel { $State.Workbook.Activate() }
            $frameWindow = Invoke-Excel { $State.Excel.ActiveWindow }
            $windowHandle = [int64](Invoke-Excel { $frameWindow.Hwnd })
            Release-Com $frameWindow
            $applicationHandle = [int64](Invoke-Excel { $State.Excel.Hwnd })
            if ($windowHandle -eq $applicationHandle) {
                return ''
            }
            $lastRead = 'ActiveWindow.Hwnd {0}, Application.Hwnd {1}' -f $windowHandle, $applicationHandle
        }
        catch {
            $lastRead = Get-InnerMessage $_
        }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    return $lastRead
}

function Read-TitleBarState($State) {
    try {
        return [string](Invoke-Excel { $State.Excel.Run($State.Qualified + $DemoTitleBarStateMacro) })
    }
    catch {
        return 'UNKNOWN | Application.Run raised: ' + (Get-InnerMessage $_)
    }
}

function Close-ExcelInstance($State) {
    # Closes the workbook without saving, closes any workbook a failed run
    # left open, quits the instance and, as the certification script does,
    # ends its process only when it is still the instance this script started.
    # Returns the status and the detail of the close step.
    if ($null -eq $State -or $State.Closed -or $null -eq $State.Excel) {
        return $null
    }
    $State.Closed = $true
    $problems = @()
    if ($null -ne $State.Vbe) {
        try { $State.Vbe.MainWindow.Visible = $false } catch { }
        Release-Com $State.Vbe
        $State.Vbe = $null
    }
    if ($null -ne $State.Workbook) {
        try {
            Invoke-Excel { $State.Workbook.Close($false) }
        }
        catch {
            $problems += 'the workbook did not close: ' + (Get-InnerMessage $_)
        }
        Release-Com $State.Workbook
        $State.Workbook = $null
    }
    try {
        Invoke-Excel { $State.Excel.DisplayAlerts = $false }
        Invoke-Excel { $State.Excel.Quit() }
    }
    catch {
        $problems += 'Excel did not quit: ' + (Get-InnerMessage $_)
    }
    Release-Com $State.Excel
    $State.Excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    $detail = 'Workbook closed without saving; Excel quit'
    if ($null -ne $State.ProcessId -and $null -ne $State.ProcessStart) {
        switch (Stop-QuitProcess -ProcessId $State.ProcessId -ExpectedStartTime $State.ProcessStart) {
            'stopped' { $detail += ('; process {0} did not exit after Quit and was ended' -f $State.ProcessId) }
            'running' {
                $problems += ('process {0} is still running, close it yourself' -f $State.ProcessId)
            }
        }
    }
    if ($problems.Count -gt 0) {
        return [pscustomobject]@{ Status = 'FAIL'; Detail = ($problems -join '; ') }
    }
    return [pscustomobject]@{ Status = 'PASS'; Detail = $detail }
}

function Complete-CloseStep([string]$Name, $State) {
    $closed = Close-ExcelInstance $State
    if ($null -ne $closed) {
        Set-CheckStep $Name $closed.Status $closed.Detail
    }
}

function Get-CertificationWorkbook([string]$Folder, $Policy) {
    # The workbook the certification run built, as its session record names
    # it, or $null when there is none.
    $sessionPath = Join-Path $Folder ([string]$Policy.required_files.session)
    if (-not (Test-Path -LiteralPath $sessionPath -PathType Leaf)) {
        return $null
    }
    $record = [System.IO.File]::ReadAllText($sessionPath, $Utf8NoBom) | ConvertFrom-Json
    if ($null -eq $record.workbook -or -not $record.workbook.file) {
        return $null
    }
    $path = Join-Path $Folder ([string]$record.workbook.file)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return $null
    }
    return $path
}

function Get-ReportResultLine([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ''
    }
    $line = @(Select-String -LiteralPath $Path -Pattern '^RESULT: ' | Select-Object -First 1)
    if ($line.Count -eq 0) {
        return ''
    }
    return [string]$line[0].Line
}

# ---------------------------------------------------------------------------
# Shared definitions
# ---------------------------------------------------------------------------
$CertifyScript = Join-Path $PSScriptRoot $CertifyScriptName
try {
    $shared = Import-CertifyDefinitions -Path $CertifyScript -Functions $SharedFunctions -Variables $SharedValues
}
catch {
    Write-Host ('ERROR: ' + $_.Exception.Message)
    exit 1
}
. $shared.Definitions
$MsoAutomationSecurityByUI = $shared.Values['MsoAutomationSecurityByUI']
$MsoAutomationSecurityLow = $shared.Values['MsoAutomationSecurityLow']
$CompileControlId = $shared.Values['CompileControlId']
$ReadinessMacro = $shared.Values['ReadinessMacro']
$MacrosWaitSeconds = $shared.Values['MacrosWaitSeconds']
$MacrosPollSeconds = $shared.Values['MacrosPollSeconds']
$CompileWaitSeconds = $shared.Values['CompileWaitSeconds']

# ---------------------------------------------------------------------------
# Repository, commit and bundle
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent $PSScriptRoot
}
$script:RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path.TrimEnd('\')
$policyPath = Join-Path (Join-Path $script:RepoRoot '.github') 'excel-evidence-policy.json'

try {
    $policy = Read-EvidencePolicy $policyPath
    $runnerList = Resolve-RunnerList -Requested $Runners -IncludeProbe:$IncludeRibbonProbe -Default $DefaultRunners `
        -Probe $ProbeRunner
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
if ($dirtyPaths.Count -gt 0) {
    Write-Host 'ERROR: the working tree has changes. Check a clean checkout (git status shows none):'
    $dirtyPaths | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$started = Get-Date
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $script:RepoRoot) `
        ('EXCEL_UI-checks-{0}-{1}' -f $sha.Substring(0, 12), $started.ToString('yyyyMMdd-HHmmss', $Invariant))
}
$script:BundleDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BundleDir).TrimEnd('\')
if (($script:BundleDir + '\').StartsWith($script:RepoRoot + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
    ($script:BundleDir + '/').StartsWith($script:RepoRoot + '/', [System.StringComparison]::Ordinal)) {
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
$certificationDir = Join-Path $script:BundleDir $CertificationFolder
$demoWorkbookPath = Join-Path $script:BundleDir $DemoWorkbookFileName
$demoBuildLog = 'demo-build.log'
$titleBarLog = 'demo-titlebar.log'

$runnerSteps = @($runnerList | ForEach-Object { New-CheckStep $_ ('runner-{0}.log' -f $_) })
$script:Steps = @(
    (New-CheckStep 'preflight' ''),
    (New-CheckStep 'certification' ('{0}/session.log' -f $CertificationFolder)),
    (New-CheckStep 'open-development' '')
) + $runnerSteps + @(
    (New-CheckStep 'close-development' ''),
    (New-CheckStep 'demo-build' $demoBuildLog),
    (New-CheckStep 'demo-open' ''),
    (New-CheckStep 'demo-compile' ''),
    (New-CheckStep 'demo-cleanup-checks' 'demo-cleanup-checks.log'),
    (New-CheckStep 'demo-create-sheet-1' 'demo-create-sheet-1.log'),
    (New-CheckStep 'demo-create-sheet-2' 'demo-create-sheet-2.log'),
    (New-CheckStep 'default-printer' ''),
    (New-CheckStep 'demo-titlebar' $titleBarLog),
    (New-CheckStep 'close-demo' '')
)
$runnerNames = @($runnerList)
$demoChecks = @('demo-cleanup-checks', 'demo-create-sheet-1', 'demo-create-sheet-2', 'demo-titlebar')

$script:Session = [ordered]@{
    schema           = 1
    tool             = 'tools/Check-EXCEL_UI.ps1'
    commit           = $sha
    started_utc      = Get-UtcStamp $started
    started_local    = Get-LocalStamp $started
    finished_utc     = $null
    finished_local   = $null
    duration_seconds = $null
    runners          = $runnerNames
    slow_build_seconds = $SlowBuildSeconds
    host             = [ordered]@{ windows = $null; excel = $null; trust_center = $null }
    certification    = [ordered]@{ folder = $CertificationFolder; exit_code = $null; result_line = '' }
    default_printer  = $null
    steps            = $script:Steps
    result           = 'FAIL'
    exit_code        = 1
    message          = ''
}

Write-Log "VBA Excel UI checks of commit $sha"
Write-Log "Bundle: $script:BundleDir"
Write-Log ('Runners: ' + ($runnerNames -join ', '))

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------
$devState = New-ExcelState 'development'
$demoState = New-ExcelState 'demo'
$cancelled = $false

try {
    # --- preflight ----------------------------------------------------------
    try {
        Set-CheckStep 'preflight' 'RUNNING' 'Reading Windows and the Trust Center settings'
        $officeVersion = Get-OfficeVersion
        $macroSetting = Get-MacroSetting $officeVersion
        $vbaProjectAccess = Get-VbaProjectAccess $officeVersion
        $script:Session.host.trust_center = [ordered]@{
            macro_setting       = $macroSetting
            vba_project_access  = $vbaProjectAccess
            automation_security = (Get-AutomationSecurityLabel)
        }
        Write-Log "Macro setting: $macroSetting; Trust access to the VBA project object model: $vbaProjectAccess"
        if ($vbaProjectAccess -ne 'On') {
            throw ("'Trust access to the VBA project object model' is off. Building the workbooks needs it: " +
                   'turn it on in File > Options > Trust Center > Trust Center Settings > Macro Settings, ' +
                   'if your policy allows. This script never changes it.')
        }
        $windows = Get-WindowsInfo
        $script:Session.host.windows = $windows
        $syncScript = Join-Path $script:RepoRoot 'tools\Sync-EXCEL_UI-VBA.ps1'
        if (-not (Test-Path -LiteralPath $syncScript -PathType Leaf)) {
            throw "The synchronizer is missing: $syncScript"
        }
        Set-CheckStep 'preflight' 'PASS' ('Commit {0}; {1} {2} build {3}' -f $sha, $windows.caption, $windows.version,
                                          $windows.build)
    }
    catch {
        Register-StepError 'preflight' $_
    }

    if ((Get-CheckStepStatus 'preflight') -ne 'PASS') {
        Skip-CheckStep @($script:Steps | ForEach-Object { $_.name }) 'the preflight failed'
    }
    else {
        # --- certification --------------------------------------------------
        # Run unchanged as a child process with the console of this one, so
        # its prompts reach the operator, or with the same -Unattended, so it
        # has none.
        Set-CheckStep 'certification' 'RUNNING' "tools\$CertifyScriptName; it builds, compiles and certifies"
        $certifyArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $CertifyScript,
                              '-BundleDir', $certificationDir, '-RepoRoot', $script:RepoRoot)
        if ($Python) {
            $certifyArguments += @('-Python', $Python)
        }
        if ($Unattended) {
            $certifyArguments += '-Unattended'
        }
        $certifyCode = Invoke-Interactive -FilePath 'powershell.exe' -Arguments $certifyArguments
        $reportLine = Get-ReportResultLine (Join-Path $certificationDir ([string]$policy.required_files.certification_report))
        $script:Session.certification.exit_code = $certifyCode
        $script:Session.certification.result_line = $reportLine
        $certificationDetail = 'exit {0}; {1}' -f $certifyCode,
            $(if ($reportLine) { $reportLine } else { 'no RESULT line in the certification report' })
        switch ($certifyCode) {
            0 { Set-CheckStep 'certification' 'PASS' $certificationDetail }
            2 {
                Set-CheckStep 'certification' 'CANCELLED' $certificationDetail
                Stop-Cancelled 'The certification run was cancelled.'
            }
            default { Set-CheckStep 'certification' 'FAIL' $certificationDetail }
        }

        # --- development workbook and runners -------------------------------
        $devWorkbook = Get-CertificationWorkbook $certificationDir $policy
        if ($null -eq $devWorkbook) {
            Skip-CheckStep (@('open-development') + $runnerNames + @('close-development')) `
                'the certification bundle holds no development workbook'
        }
        else {
            try {
                Set-CheckStep 'open-development' 'RUNNING' "Opening $devWorkbook"
                Open-ExcelWorkbook $devState $devWorkbook
                Wait-WorkbookMacros $devState
                $script:Session.host.excel = Get-ExcelHost $devState
                Set-CheckStep 'open-development' 'PASS' ('Excel {0} build {1}, {2}; process {3}' -f
                    $script:Session.host.excel.version, $script:Session.host.excel.build,
                    $script:Session.host.excel.bitness, $devState.ProcessId)
            }
            catch {
                Register-StepError 'open-development' $_
            }
            if ((Get-CheckStepStatus 'open-development') -eq 'PASS') {
                Write-Host 'Do not use Excel until the runners end. They open and close windows.'
                foreach ($step in $runnerSteps) {
                    Set-CheckStep $step.name 'RUNNING' ('{0} via {1}' -f $step.name, $RunnerMacro)
                    $unsettled = Wait-ActiveFrame $devState
                    $outcome = Invoke-LoggedMacro -State $devState -EntryPoint $RunnerMacro -Macro $step.name `
                        -LogPath (Join-Path $script:BundleDir $step.log)
                    $status = $(if ($step.name -eq $ProbeRunner) { 'INFO' } elseif ($outcome.Passed) { 'PASS' } else { 'FAIL' })
                    $runnerDetail = $outcome.Detail
                    if ($unsettled) {
                        $runnerDetail += ' | the active frame did not settle before the run: ' + $unsettled
                    }
                    Set-CheckStep $step.name $status $runnerDetail -Seconds $outcome.Seconds
                }
            }
            else {
                Skip-CheckStep $runnerNames 'the development workbook did not open'
            }
            Complete-CloseStep 'close-development' $devState
            Skip-CheckStep @('close-development') 'no Excel instance was started'
        }

        # --- demo workbook --------------------------------------------------
        try {
            Set-CheckStep 'demo-build' 'RUNNING' 'tools\Sync-EXCEL_UI-VBA.ps1 -Mode Demo -NewWorkbook'
            $buildLogPath = Join-Path $script:BundleDir $demoBuildLog
            $previousPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try {
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript `
                    -WorkbookPath $demoWorkbookPath -Mode Demo -NewWorkbook -Force -NoBackup 2>&1 |
                    ForEach-Object {
                        $text = [string]$_
                        Add-LogLine $buildLogPath $text
                        Write-Log ('  sync: ' + $text)
                    }
                $syncCode = $LASTEXITCODE
            }
            finally {
                $ErrorActionPreference = $previousPreference
            }
            if ($syncCode -ne 0 -or -not (Test-Path -LiteralPath $demoWorkbookPath -PathType Leaf)) {
                throw "The synchronizer exited with code $syncCode."
            }
            Set-CheckStep 'demo-build' 'PASS' "$DemoWorkbookFileName built from the checkout"
        }
        catch {
            Register-StepError 'demo-build' $_
        }

        if ((Get-CheckStepStatus 'demo-build') -eq 'PASS') {
            try {
                Set-CheckStep 'demo-open' 'RUNNING' "Opening $DemoWorkbookFileName"
                Open-ExcelWorkbook $demoState $demoWorkbookPath
                Wait-WorkbookMacros $demoState
                $demoHost = Get-ExcelHost $demoState
                if ($null -eq $script:Session.host.excel) {
                    $script:Session.host.excel = $demoHost
                }
                Set-CheckStep 'demo-open' 'PASS' ('Excel {0} build {1}, {2}; process {3}' -f $demoHost.version,
                                                  $demoHost.build, $demoHost.bitness, $demoState.ProcessId)
            }
            catch {
                Register-StepError 'demo-open' $_
            }
        }
        else {
            Skip-CheckStep @('demo-open') 'the demo workbook was not built'
        }

        if ((Get-CheckStepStatus 'demo-open') -eq 'PASS') {
            try {
                Set-CheckStep 'demo-compile' 'RUNNING' 'Debug > Compile VBAProject'
                Set-CheckStep 'demo-compile' 'PASS' (Invoke-ProjectCompile $demoState)
            }
            catch {
                Register-StepError 'demo-compile' $_
            }
        }
        else {
            Skip-CheckStep @('demo-compile') 'the demo workbook did not open'
        }

        if ((Get-CheckStepStatus 'demo-compile') -eq 'PASS') {
            # --- demo-cleanup-checks ----------------------------------------
            $step = Find-CheckStep 'demo-cleanup-checks'
            Set-CheckStep $step.name 'RUNNING' $DemoCleanupMacro
            $outcome = Invoke-LoggedMacro -State $demoState -EntryPoint $DemoLoggedMacro -Macro $DemoCleanupMacro `
                -LogPath (Join-Path $script:BundleDir $step.log)
            Set-CheckStep $step.name (Get-OutcomeStatus $outcome) $outcome.Detail -Seconds $outcome.Seconds

            # --- demo-create-sheet-1 and -2 ---------------------------------
            foreach ($name in @('demo-create-sheet-1', 'demo-create-sheet-2')) {
                $step = Find-CheckStep $name
                Set-CheckStep $name 'RUNNING' $DemoBuildMacro
                $outcome = Invoke-LoggedMacro -State $demoState -EntryPoint $DemoLoggedMacro -Macro $DemoBuildMacro `
                    -LogPath (Join-Path $script:BundleDir $step.log)
                $timing = '{0} took {1} s (threshold {2} s)' -f $DemoBuildMacro, (Format-Seconds $outcome.Seconds),
                    (Format-Seconds $SlowBuildSeconds)
                if ($outcome.Passed) {
                    Set-CheckStep $name (Get-TimingStatus $outcome.Seconds $SlowBuildSeconds) $timing -Seconds $outcome.Seconds
                }
                else {
                    Set-CheckStep $name 'FAIL' ($outcome.Detail + '; ' + $timing) -Seconds $outcome.Seconds
                }
            }
        }
        else {
            Skip-CheckStep $demoChecks 'the demo workbook did not compile'
        }

        # --- default-printer ------------------------------------------------
        try {
            Set-CheckStep 'default-printer' 'RUNNING' 'Win32_Printer where Default = TRUE'
            $printer = Get-DefaultPrinter
            $script:Session.default_printer = $printer
            if ($null -eq $printer) {
                Set-CheckStep 'default-printer' 'INFO' 'no default printer'
            }
            else {
                Set-CheckStep 'default-printer' 'INFO' ('{0}; work offline {1}; status {2}; network {3}' -f
                    $printer.name, $printer.work_offline, $printer.printer_status_text, $printer.network)
            }
        }
        catch {
            Set-CheckStep 'default-printer' 'INFO' ('not read: ' + (Get-InnerMessage $_))
        }

        # --- demo-titlebar --------------------------------------------------
        $sheetBuilt = @('PASS', 'WARN') -contains (Get-CheckStepStatus 'demo-create-sheet-1') -or
                      @('PASS', 'WARN') -contains (Get-CheckStepStatus 'demo-create-sheet-2')
        if ((Get-CheckStepStatus 'demo-titlebar') -eq 'NOT_RUN' -and -not $sheetBuilt) {
            Skip-CheckStep @('demo-titlebar') "$DemoBuildMacro did not build the demo sheet"
        }
        elseif ((Get-CheckStepStatus 'demo-titlebar') -eq 'NOT_RUN') {
            try {
                Set-CheckStep 'demo-titlebar' 'RUNNING' "Only $TitleBarCheckBox selected; HIDE, then SHOW"
                $logPath = Join-Path $script:BundleDir $titleBarLog
                $outcome = Invoke-LoggedMacro -State $demoState -EntryPoint $DemoLoggedMacro -Macro $DemoClearMacro `
                    -LogPath $logPath
                if (-not $outcome.Passed) {
                    throw "$DemoClearMacro`: $($outcome.Detail)"
                }
                # The click on the check box: only the title bar is selected.
                $sheet = Invoke-Excel { $demoState.Workbook.Worksheets.Item($DemoSheetName) }
                $shape = Invoke-Excel { $sheet.Shapes.Item($TitleBarCheckBox) }
                $format = Invoke-Excel { $shape.ControlFormat }
                Invoke-Excel { $format.Value = $XlOn }
                Release-Com $format
                Release-Com $shape
                Release-Com $sheet
                Add-LogLine $logPath "CHECK $TitleBarCheckBox selected; the other check boxes cleared"
                $before = Read-TitleBarState $demoState
                Add-LogLine $logPath "CHECK title bar before HIDE: $before"
                $outcome = Invoke-LoggedMacro -State $demoState -EntryPoint $DemoLoggedMacro -Macro $DemoHideMacro `
                    -LogPath $logPath
                if (-not $outcome.Passed) {
                    throw "$DemoHideMacro`: $($outcome.Detail)"
                }
                $afterHide = Read-TitleBarState $demoState
                Add-LogLine $logPath "CHECK title bar after HIDE: $afterHide"
                $outcome = Invoke-LoggedMacro -State $demoState -EntryPoint $DemoLoggedMacro -Macro $DemoShowMacro `
                    -LogPath $logPath
                if (-not $outcome.Passed) {
                    throw "$DemoShowMacro`: $($outcome.Detail)"
                }
                $afterShow = Read-TitleBarState $demoState
                Add-LogLine $logPath "CHECK title bar after SHOW: $afterShow"
                $titleBarDetail = 'before {0}; after HIDE {1}; after SHOW {2}' -f $before, $afterHide, $afterShow
                $titleBarStatus = $(if ($afterHide -eq 'HIDDEN' -and $afterShow -eq 'VISIBLE') { 'PASS' } else { 'FAIL' })
                Add-LogLine $logPath "RESULT: $titleBarStatus | $titleBarDetail"
                Set-CheckStep 'demo-titlebar' $titleBarStatus $titleBarDetail
            }
            catch {
                Register-StepError 'demo-titlebar' $_
            }
        }

        Complete-CloseStep 'close-demo' $demoState
        Skip-CheckStep @('close-demo') 'no Excel instance was started'
    }
}
catch [System.OperationCanceledException] {
    $cancelled = $true
    $script:Session.message = [string]$_.Exception.Message
    $current = @($script:Steps | Where-Object { $_.status -eq 'RUNNING' -or $_.status -eq 'CANCELLED' })
    if ($current.Count -gt 0) {
        Set-CheckStep $current[0].name 'CANCELLED' $script:Session.message
    }
    else {
        Write-Log ('CANCELLED: ' + $script:Session.message)
    }
}
catch {
    $script:Session.message = Get-InnerMessage $_
    Write-Log ('ERROR: ' + $script:Session.message)
    $current = @($script:Steps | Where-Object { $_.status -eq 'RUNNING' })
    if ($current.Count -gt 0) {
        Set-CheckStep $current[0].name 'FAIL' $script:Session.message
    }
}
finally {
    # Never leave an Excel instance this script started running or modal.
    Complete-CloseStep 'close-development' $devState
    Complete-CloseStep 'close-demo' $demoState
}

# ---------------------------------------------------------------------------
# Record and summarize
# ---------------------------------------------------------------------------
foreach ($step in @($script:Steps | Where-Object { $_.status -eq 'RUNNING' })) {
    Set-CheckStep $step.name 'FAIL' ('stopped: ' + $step.detail)
}
if ($cancelled) {
    foreach ($step in @($script:Steps | Where-Object { $_.status -eq 'NOT_RUN' })) {
        $step.detail = 'not run: the run was cancelled'
    }
}
else {
    Skip-CheckStep @($script:Steps | ForEach-Object { $_.name }) 'the run stopped before this step'
}
# A step names only a log that exists.
foreach ($step in @($script:Steps | Where-Object { $_.log })) {
    if (-not (Test-Path -LiteralPath (Join-Path $script:BundleDir $step.log) -PathType Leaf)) {
        $step.log = ''
    }
}

$finished = Get-Date
$script:Session.finished_utc = Get-UtcStamp $finished
$script:Session.finished_local = Get-LocalStamp $finished
$script:Session.duration_seconds = [math]::Round(($finished - $started).TotalSeconds, 1)
$runResult = Get-RunResult -Steps $script:Steps -Cancelled $cancelled
$script:Session.result = $(if ($cancelled) { 'CANCELLED' } else { $runResult.Result })
$script:Session.exit_code = $runResult.ExitCode
Write-JsonFile 'session.json' $script:Session

$hostLine = ''
$hostRecord = $script:Session.host
if ($null -ne $hostRecord.excel) {
    $hostLine = 'Excel {0} build {1}, {2}' -f $hostRecord.excel.version, $hostRecord.excel.build, $hostRecord.excel.bitness
}
if ($null -ne $hostRecord.windows) {
    $windowsText = '{0} {1} build {2}' -f $hostRecord.windows.caption, $hostRecord.windows.version, $hostRecord.windows.build
    $hostLine = $(if ($hostLine) { "$hostLine; $windowsText" } else { $windowsText })
}
$summary = Format-Summary -Commit $sha -HostLine $hostLine -Steps $script:Steps -Result $runResult.Result
[System.IO.File]::WriteAllText((Join-Path $script:BundleDir 'summary.txt'), (($summary -join "`r`n") + "`r`n"), $Utf8NoBom)
Write-Log ('Run finished: {0}' -f $script:Session.result)

Write-Host ''
Write-Host "Bundle: $script:BundleDir"
Write-Host 'Summary (summary.txt):'
Write-Host ''
$summary | ForEach-Object { Write-Host $_ }
exit ([int]$runResult.ExitCode)
