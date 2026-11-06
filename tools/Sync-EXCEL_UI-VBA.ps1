<#
.SYNOPSIS
Import the VBA modules of this repository into a local .xlsm workbook.

.DESCRIPTION
Updates VBA code only. A full synchronization imports the standard modules
that tools\vba-projects.json lists for one configuration, and asks which one
unless -Mode gives the answer:

- Production: the four modules of src\.
- Development: the production modules and the regression module of test\,
  for a workbook that runs the regression suite and release certification.
- Demo: the production modules and the modules of demo\.

Every module the configuration lists is replaced from the source. Modules
the configuration does not list stay in the workbook: a synchronization
never removes one, and it names the repository modules it leaves in place.
The script does not write cells, formulas, names, formatting, worksheets or
workbook structure, and it checks that the sheet names and code names are
unchanged before it saves.

Every .bas file in src\, test\ and demo\ must be listed in at least one
configuration of the manifest; an unlisted module stops the script before
Excel starts.

-NewWorkbook creates a blank .xlsm at -WorkbookPath first, in an Excel
instance of its own, replacing a file of that name, and then synchronizes
it. It is refused with -SourceDir. tools\Sync-EXCEL_UI-VBA.cmd started
without a workbook uses it to build test-results\EXCEL_UI_Development.xlsm.

With -SourceDir only the modules found in that folder are imported, for
example modules downloaded for a review, without copying them into the
repository. A file is identified only by the Attribute VB_Name it declares,
never by its file name: a file without exactly one valid declaration is
ignored. A file is taken when that VB_Name is a module of the repository
with the same extension, or when it is a standard module (.bas) whose name
matches a project_name_patterns entry of the manifest that the repository's
own module names match (^M_), such as a module added by the pull request
under review; it is listed apart as a new module. When several files carry
the same VB_Name, for example "M_EXCEL_UI (1).bas", the newest is taken. The
script lists the chosen and ignored files and asks for confirmation unless
-Force is given. Every other module in the workbook is left as it is.

Files are copied to a temporary folder with CRLF line endings and imported
from there; the repository and the -SourceDir folder are never changed.

Exit codes: 0 success, 1 error, 2 cancelled or nothing to import (the
workbook was not changed).

.REQUIREMENTS
- Windows desktop Excel.
- Excel > Options > Trust Center > Macro Settings >
  "Trust access to the VBA project object model" enabled.
- VBA project not password-locked.
- The workbook closed in every Excel window.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath "C:\work\EXCEL_UI.xlsm"

Asks for the configuration, then synchronizes the workbook.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath ".\test-results\EXCEL_UI_Development.xlsm" -Mode Development -NewWorkbook

Creates a blank development workbook and fills it.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath "C:\work\EXCEL_UI.xlsm" -SourceDir "$env:USERPROFILE\Downloads"

Imports the modules found in the Downloads folder.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Production", "Development", "Demo")]
    [string]$Mode,

    [Parameter(Mandatory = $false)]
    [switch]$NewWorkbook,

    [Parameter(Mandatory = $false)]
    [string]$SourceDir,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$NoBackup,

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$vbext_ct_StdModule = 1
$xlCalculationManual = -4135
$msoAutomationSecurityForceDisable = 3
$xlOpenXMLWorkbookMacroEnabled = 52

$ManifestRelativePath = "tools/vba-projects.json"
$SourceDirectories = @("src", "test", "demo")
$Modes = @("Production", "Development", "Demo")

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)
    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PathValue).Path)
}

function Get-DeclaredVbaName {
    # The identity a file declares: exactly one Attribute VB_Name line
    # holding a valid VBA name. The file name is never used instead, so an
    # "M_EXCEL_UI.bas" without the attribute cannot replace M_EXCEL_UI.
    # Returns the name and, when there is none, why. A list, not a set: a
    # declaration repeated identically, or in another case, is a second one.
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"\s*$') {
            $names.Add($Matches[1])
        }
        elseif ($line -match '^\s*Attribute\s+VB_Name\b') {
            return [pscustomobject]@{ Name = ""; Problem = "malformed Attribute VB_Name line" }
        }
    }

    if ($names.Count -eq 0) {
        return [pscustomobject]@{ Name = ""; Problem = "no Attribute VB_Name line; the file name is not used as its identity" }
    }
    if ($names.Count -gt 1) {
        return [pscustomobject]@{ Name = ""; Problem = "several Attribute VB_Name lines" }
    }

    $name = $names[0]
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,30}$') {
        return [pscustomobject]@{ Name = ""; Problem = "VB_Name '$name' is not a valid VBA name" }
    }

    return [pscustomobject]@{ Name = $name; Problem = "" }
}

function Copy-WithCrlf {
    # ISO-8859-1 maps every byte to one character and back, so the cp1252
    # bytes of a module are kept exactly; only the line endings change.
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    $encoding = [System.Text.Encoding]::GetEncoding(28591)
    $text = $encoding.GetString([System.IO.File]::ReadAllBytes($SourcePath))
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
    [System.IO.File]::WriteAllBytes($TargetPath, $encoding.GetBytes($text))
}

function ConvertTo-SyncMode {
    # The answer to the configuration question: Production, Development,
    # Demo, or $null for anything else.
    param([AllowNull()][AllowEmptyString()][string]$Answer)

    switch -Regex (([string]$Answer).Trim()) {
        '^(?i)(1|p|prod|production)$' { return "Production" }
        '^(?i)(2|d|dev|development)$' { return "Development" }
        '^(?i)(3|demo)$' { return "Demo" }
    }
    return $null
}

function Read-VbaProjectManifest {
    # tools/vba-projects.json. This synchronizer imports standard modules
    # only, so a manifest that lists document modules is refused.
    param([Parameter(Mandatory = $true)][string]$RepoPath)

    $manifestPath = Join-Path $RepoPath $ManifestRelativePath
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Module manifest not found: $manifestPath"
    }

    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $manifest.PSObject.Properties["configurations"]) {
        throw "No configurations in $manifestPath"
    }
    if ($null -ne $manifest.PSObject.Properties["document_modules"] -and @($manifest.document_modules).Count -gt 0) {
        throw "$manifestPath lists document modules; this synchronizer imports standard modules only."
    }
    return $manifest
}

function Get-ManifestSources {
    # The modules of one configuration of tools/vba-projects.json, in the
    # manifest's order. Every listed file must exist and be a standard module.
    param(
        [Parameter(Mandatory = $true)][string]$RepoPath,
        [Parameter(Mandatory = $true)][ValidateSet("Production", "Development", "Demo")][string]$Mode
    )

    $manifest = Read-VbaProjectManifest -RepoPath $RepoPath
    $key = $Mode.ToLowerInvariant()
    $property = $manifest.configurations.PSObject.Properties[$key]
    if ($null -eq $property) {
        throw "Configuration '$key' not found in $ManifestRelativePath"
    }

    $files = @(
        foreach ($entry in @($property.Value.modules)) {
            $relative = [string]$entry
            if ([System.IO.Path]::GetExtension($relative).ToLowerInvariant() -ne ".bas") {
                throw "Module listed in $ManifestRelativePath is not a standard module (.bas): $relative"
            }
            $path = Join-Path $RepoPath $relative
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "Module listed in $ManifestRelativePath not found: $relative"
            }
            Get-Item -LiteralPath $path
        }
    )
    return $files
}

function Get-RepositorySources {
    # Every .bas file under src/, test/ and demo/, sorted by path.
    param([Parameter(Mandatory = $true)][string]$RepoPath)

    $files = @(
        foreach ($directory in $SourceDirectories) {
            $root = Join-Path $RepoPath $directory
            if (-not (Test-Path -LiteralPath $root -PathType Container)) {
                throw "Required source directory not found: $root"
            }
            Get-ChildItem -LiteralPath $root -File -Recurse |
                Where-Object { $_.Extension.ToLowerInvariant() -eq ".bas" }
        }
    )
    return @($files | Sort-Object FullName)
}

function Get-RepositoryNames {
    # The repository modules by lower-case VB_Name: @{ name = FileInfo }.
    # Each must declare one valid VB_Name, and no name may appear twice.
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SourceFiles)

    $names = @{}
    foreach ($file in $SourceFiles) {
        $declared = Get-DeclaredVbaName -SourcePath $file.FullName
        if (-not [string]::IsNullOrEmpty($declared.Problem)) {
            throw "Repository module $($file.FullName): $($declared.Problem)"
        }
        $key = $declared.Name.ToLowerInvariant()
        if ($names.ContainsKey($key)) {
            throw "Duplicate VBA module name '$($declared.Name)' in the repository source."
        }
        $names[$key] = $file
    }
    return $names
}

function Get-UnlistedSources {
    # The repository modules that no configuration of the manifest lists.
    param(
        [Parameter(Mandatory = $true)][string]$RepoPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SourceFiles
    )

    $listed = @{}
    foreach ($mode in $Modes) {
        foreach ($file in @(Get-ManifestSources -RepoPath $RepoPath -Mode $mode)) {
            $listed[$file.FullName.ToLowerInvariant()] = $true
        }
    }
    return @($SourceFiles | Where-Object { -not $listed.ContainsKey($_.FullName.ToLowerInvariant()) })
}

function Assert-ManifestComplete {
    # Stops when a module of src/, test/ or demo/ is in no configuration, so
    # that a new module is never left out of every workbook unnoticed.
    param(
        [Parameter(Mandatory = $true)][string]$RepoPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SourceFiles
    )

    $unlisted = @(Get-UnlistedSources -RepoPath $RepoPath -SourceFiles $SourceFiles)
    if ($unlisted.Count -gt 0) {
        $paths = @($unlisted | ForEach-Object { $_.FullName })
        throw ("Module not listed in any configuration of $ManifestRelativePath`: " + ($paths -join ", "))
    }
}

function Get-NewModulePatterns {
    # The project_name_patterns of the manifest that name modules: those
    # that at least one repository module name matches. The others name
    # procedures, not modules.
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Patterns,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ModuleNames
    )

    return @(
        foreach ($pattern in $Patterns) {
            foreach ($name in $ModuleNames) {
                if ([regex]::IsMatch($name, $pattern)) {
                    $pattern
                    break
                }
            }
        }
    )
}

function Test-NewModuleName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if ([regex]::IsMatch($Name, $pattern)) { return $true }
    }
    return $false
}

function Copy-ToStage {
    # Copies each file to StagePath as <VB_Name>.bas with CRLF line endings
    # and returns the copies.
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Files,
        [Parameter(Mandatory = $true)][string]$StagePath
    )

    $null = New-Item -ItemType Directory -Path $StagePath -Force
    return @(
        foreach ($file in $Files) {
            $declared = Get-DeclaredVbaName -SourcePath $file.FullName
            $target = Join-Path $StagePath ($declared.Name + $file.Extension.ToLowerInvariant())
            Copy-WithCrlf -SourcePath $file.FullName -TargetPath $target
            Get-Item -LiteralPath $target
        }
    )
}

function Select-FolderSources {
    # The modules of FolderPath to import, staged in StagePath; an empty list
    # when none is usable. Prints what is taken and what is ignored, and why.
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)][hashtable]$RepositoryNames,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$NewModulePatterns,
        [Parameter(Mandatory = $true)][string]$StagePath
    )

    $picked = @{}
    $newNames = @{}
    $ignored = New-Object System.Collections.Generic.List[string]
    $candidates = @(
        Get-ChildItem -LiteralPath $FolderPath -File |
            Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm") } |
            Sort-Object Name
    )

    foreach ($file in $candidates) {
        $extension = $file.Extension.ToLowerInvariant()
        $declared = Get-DeclaredVbaName -SourcePath $file.FullName
        if (-not [string]::IsNullOrEmpty($declared.Problem)) {
            $ignored.Add("$($file.Name) ($($declared.Problem))")
            continue
        }
        $name = $declared.Name
        $key = $name.ToLowerInvariant()
        if ($RepositoryNames.ContainsKey($key)) {
            $repoExtension = $RepositoryNames[$key].Extension.ToLowerInvariant()
            if ($extension -ne $repoExtension) {
                $ignored.Add("$($file.Name) ('$name' is a $repoExtension module in the repository)")
                continue
            }
        }
        elseif ($extension -eq ".bas" -and (Test-NewModuleName -Name $name -Patterns $NewModulePatterns)) {
            $newNames[$key] = $true
        }
        else {
            $ignored.Add("$($file.Name) (VB_Name '$name' is not a repository module)")
            continue
        }

        if ($picked.ContainsKey($key)) {
            if ($file.LastWriteTime -gt $picked[$key].LastWriteTime) {
                $ignored.Add("$($picked[$key].Name) (older copy of '$name')")
                $picked[$key] = $file
            }
            else {
                $ignored.Add("$($file.Name) (older copy of '$name')")
            }
        }
        else {
            $picked[$key] = $file
        }
    }

    if ($picked.Count -gt 0) {
        Write-Host "Modules to import from $FolderPath"
        foreach ($key in ($picked.Keys | Sort-Object)) {
            if ($newNames.ContainsKey($key)) { continue }
            $file = $picked[$key]
            Write-Host ("  {0,-30} {1,-36} {2:yyyy-MM-dd HH:mm}" -f (Get-DeclaredVbaName -SourcePath $file.FullName).Name, $file.Name, $file.LastWriteTime)
        }
        if ($newNames.Count -gt 0) {
            Write-Host "New modules, not in the repository checkout (imported as new standard modules):"
            foreach ($key in ($picked.Keys | Sort-Object)) {
                if (-not $newNames.ContainsKey($key)) { continue }
                $file = $picked[$key]
                Write-Host ("  {0,-30} {1,-36} {2:yyyy-MM-dd HH:mm}" -f (Get-DeclaredVbaName -SourcePath $file.FullName).Name, $file.Name, $file.LastWriteTime)
            }
        }
    }
    if ($ignored.Count -gt 0) {
        Write-Host "Ignored:"
        foreach ($line in $ignored) { Write-Host "  $line" }
    }
    Write-Host ""

    if ($picked.Count -eq 0) {
        return @()
    }
    $files = @(foreach ($key in ($picked.Keys | Sort-Object)) { $picked[$key] })
    return @(Copy-ToStage -Files $files -StagePath $StagePath)
}

function Get-LeftInPlaceModules {
    # The repository modules the workbook holds that the chosen
    # configuration does not list: a synchronization leaves them in place.
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$WorkbookModules,
        [Parameter(Mandatory = $true)][hashtable]$RepositoryNames,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ModeNames
    )

    $inMode = @{}
    foreach ($name in $ModeNames) { $inMode[$name.ToLowerInvariant()] = $true }
    return @(
        foreach ($name in $WorkbookModules) {
            $key = $name.ToLowerInvariant()
            if ($RepositoryNames.ContainsKey($key) -and -not $inMode.ContainsKey($key)) { $name }
        }
    )
}

function Find-VbaComponent {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($component in $VBProject.VBComponents) {
        if ([string]::Equals([string]$component.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $component
        }
    }
    return $null
}

function Get-SheetSnapshot {
    param([Parameter(Mandatory = $true)]$Workbook)

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in $Workbook.Sheets) {
        $codeName = ""
        try { $codeName = [string]$sheet.CodeName }
        catch { $codeName = "<no-codename>" }
        $items.Add("$($sheet.Name)|$codeName")
    }
    return @($items)
}

function Assert-SheetSnapshotUnchanged {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Before,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$After
    )

    if ($Before.Count -ne $After.Count) {
        throw ("Workbook sheet count changed from {0} to {1}. Save aborted." -f $Before.Count, $After.Count)
    }
    for ($i = 0; $i -lt $Before.Count; $i++) {
        if ($Before[$i] -cne $After[$i]) {
            throw ("Workbook sheet structure changed: sheet {0} was '{1}' (name|CodeName) and is now '{2}'. Save aborted." -f
                ($i + 1), $Before[$i], $After[$i])
        }
    }
}

function Import-StandardModule {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $existing = Find-VbaComponent -VBProject $VBProject -Name $Name
    if ($null -ne $existing) {
        $VBProject.VBComponents.Remove($existing)
    }

    $imported = $VBProject.VBComponents.Import($SourcePath)
    if (-not [string]::Equals([string]$imported.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Module '$Name' was imported as '$($imported.Name)'. Save aborted."
    }
    Write-Host ("{0,-32} standard module imported" -f $imported.Name)
}

function New-BlankWorkbook {
    # Saves a new, blank workbook as an .xlsm at TargetPath, in an Excel
    # instance of its own. A file already at TargetPath is replaced.
    param([Parameter(Mandatory = $true)][string]$TargetPath)

    $null = New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($TargetPath))
    if (Test-Path -LiteralPath $TargetPath -PathType Leaf) {
        Remove-Item -LiteralPath $TargetPath -Force
    }

    $builder = $null
    $book = $null
    try {
        $builder = New-Object -ComObject Excel.Application
        $builder.Visible = $false
        $builder.DisplayAlerts = $false
        $builder.EnableEvents = $false
        $builder.AutomationSecurity = $msoAutomationSecurityForceDisable
        $book = $builder.Workbooks.Add()
        $book.SaveAs($TargetPath, $xlOpenXMLWorkbookMacroEnabled)
        $book.Close($false)
        $book = $null
    }
    finally {
        if ($null -ne $book) {
            try { $book.Close($false) } catch {}
        }
        if ($null -ne $builder) {
            try { $builder.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($builder) } catch {}
        }
    }
}

$excel = $null
$workbook = $null
$saved = $false
$backupPath = $null
$stagePath = $null
$syncMode = $null
$restore = $null

try {
    # Resolve the repository root after parameter binding: $PSScriptRoot is
    # not reliable as a parameter default in every Windows PowerShell mode.
    if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrWhiteSpace($scriptPath)) {
            throw "Cannot determine the script path. Supply -RepoRoot explicitly."
        }
        $RepoRoot = Split-Path -Parent (Split-Path -Parent $scriptPath)
    }
    $repoFullPath = Resolve-FullPath -PathValue $RepoRoot

    # A workbook made by -NewWorkbook need not exist yet.
    if ($NewWorkbook) {
        $workbookFullPath = [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkbookPath))
    }
    else {
        if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
            throw "Workbook not found: $WorkbookPath"
        }
        $workbookFullPath = Resolve-FullPath -PathValue $WorkbookPath
    }
    if ([System.IO.Path]::GetExtension($workbookFullPath).ToLowerInvariant() -ne ".xlsm") {
        throw "WorkbookPath must point to an .xlsm file."
    }

    $folderImport = -not [string]::IsNullOrWhiteSpace($SourceDir)
    if ($folderImport -and $NewWorkbook) {
        throw "Use -NewWorkbook only with a full synchronization, not with -SourceDir."
    }
    if ($folderImport -and -not [string]::IsNullOrWhiteSpace($Mode)) {
        throw "Use -Mode only with a full synchronization, not with -SourceDir."
    }

    # Everything below up to the backup reads files only: a refusal leaves
    # the workbook as it was.
    $manifest = Read-VbaProjectManifest -RepoPath $repoFullPath
    $repositorySources = @(Get-RepositorySources -RepoPath $repoFullPath)
    if ($repositorySources.Count -eq 0) {
        throw "No VBA modules found in src, test or demo of $repoFullPath"
    }
    $repositoryNames = Get-RepositoryNames -SourceFiles $repositorySources
    Assert-ManifestComplete -RepoPath $repoFullPath -SourceFiles $repositorySources

    $stagePath = Join-Path ([System.IO.Path]::GetTempPath()) ("excel-ui-vba-sync-" + (Get-Date -Format "yyyyMMdd-HHmmss") + "-" + $PID)
    $modeNames = @()

    if (-not $folderImport) {
        $syncMode = $Mode
        if ([string]::IsNullOrWhiteSpace($syncMode)) {
            Write-Host "Which configuration should the workbook hold?"
            Write-Host "  1 = Production:  the modules of src"
            Write-Host "  2 = Development: src and the regression module of test"
            Write-Host "  3 = Demo:        src and the modules of demo"
            $syncMode = ConvertTo-SyncMode -Answer (Read-Host "Production, development or demo? [1/2/3]")
            if ($null -eq $syncMode) {
                Write-Host "Cancelled: the workbook was not changed."
                exit 2
            }
        }

        $modeSources = @(Get-ManifestSources -RepoPath $repoFullPath -Mode $syncMode)
        $modeNames = @($modeSources | ForEach-Object { (Get-DeclaredVbaName -SourcePath $_.FullName).Name })
        $importFiles = @(Copy-ToStage -Files $modeSources -StagePath $stagePath)
        $sourceLabel = $repoFullPath
        Write-Host "Mode: $syncMode ($($importFiles.Count) modules)"
        Write-Host ""
    }
    else {
        if (-not (Test-Path -LiteralPath $SourceDir -PathType Container)) {
            throw "SourceDir is not a folder: $SourceDir"
        }
        $sourceDirFull = Resolve-FullPath -PathValue $SourceDir
        $patternProperty = $manifest.PSObject.Properties["project_name_patterns"]
        $patterns = @()
        if ($null -ne $patternProperty) { $patterns = @($patternProperty.Value | ForEach-Object { [string]$_ }) }
        $repositoryModuleNames = @($repositoryNames.Values | ForEach-Object { (Get-DeclaredVbaName -SourcePath $_.FullName).Name })
        $newModulePatterns = @(Get-NewModulePatterns -Patterns $patterns -ModuleNames $repositoryModuleNames)

        $importFiles = @(Select-FolderSources -FolderPath $sourceDirFull -RepositoryNames $repositoryNames `
                -NewModulePatterns $newModulePatterns -StagePath $stagePath)
        $sourceLabel = $sourceDirFull
        if ($importFiles.Count -eq 0) {
            Write-Host "Nothing to import: no usable module in $sourceDirFull. The workbook was not changed."
            exit 2
        }
        if (-not $Force) {
            $answer = Read-Host "Import these $($importFiles.Count) module(s) into the workbook? [Y/N]"
            if ($answer -notmatch '^\s*(y|yes)\s*$') {
                Write-Host "Cancelled: the workbook was not changed."
                exit 2
            }
        }
    }

    # -NewWorkbook: the workbook is made anew, so there is nothing to back up.
    if ($NewWorkbook) {
        New-BlankWorkbook -TargetPath $workbookFullPath
        Write-Host "New workbook: $workbookFullPath"
    }
    elseif (-not $NoBackup) {
        $directory = [System.IO.Path]::GetDirectoryName($workbookFullPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
        $extension = [System.IO.Path]::GetExtension($workbookFullPath)
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupPath = Join-Path $directory "$baseName.pre-vba-sync.$timestamp$extension"
        Copy-Item -LiteralPath $workbookFullPath -Destination $backupPath -Force
        Write-Host "Backup:     $backupPath"
    }

    Write-Host "Workbook:   $workbookFullPath"
    Write-Host "Repository: $repoFullPath"
    Write-Host "Source:     $sourceLabel"
    Write-Host "Modules:    $($importFiles.Count)"
    Write-Host ""

    # An Excel instance of the script's own, with events, macros, alerts and
    # external-link updates off; its settings are put back before it quits.
    $excel = New-Object -ComObject Excel.Application
    $restore = [ordered]@{
        DisplayAlerts      = $excel.DisplayAlerts
        EnableEvents       = $excel.EnableEvents
        ScreenUpdating     = $excel.ScreenUpdating
        AskToUpdateLinks   = $excel.AskToUpdateLinks
        AutomationSecurity = $excel.AutomationSecurity
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Some Excel builds reject Application.Calculation changes until a
    # workbook is open (HRESULT 0x800A03EC), so the workbook is opened first.
    # UpdateLinks 0: external links are not updated.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)
    if ($workbook.ReadOnly) {
        throw "Workbook opened read-only. Close it in every Excel window and run the script again."
    }

    # The calculation mode is saved with the workbook, so the one it was
    # opened with is put back before the save.
    $calculation = $excel.Calculation
    $calculateBeforeSave = $excel.CalculateBeforeSave
    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw (
            "Cannot access the VBA project. Enable Excel > Options > Trust Center > Trust Center Settings > " +
            "Macro Settings > 'Trust access to the VBA project object model', and check that the VBA " +
            "project is not password-locked. The workbook was not changed."
        )
    }

    # The sheets are recorded only now. Excel leaves the CodeName of the
    # sheets of a new workbook empty until its VBA project has been read, so a
    # snapshot taken earlier would differ from the one after the import even
    # though no sheet changed. Reading every component's name first makes
    # Excel assign the CodeNames.
    foreach ($component in $vbProject.VBComponents) { $null = [string]$component.Name }
    $before = Get-SheetSnapshot -Workbook $workbook

    # Refuse before the first change when a name to import belongs to a
    # component that is not a standard module, such as a sheet or a class.
    foreach ($file in $importFiles) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $existing = Find-VbaComponent -VBProject $vbProject -Name $name
        if ($null -ne $existing -and [int]$existing.Type -ne $vbext_ct_StdModule) {
            throw "The workbook already holds '$($existing.Name)', which is not a standard module. The workbook was not changed."
        }
    }

    # A full synchronization never removes modules: repository modules that
    # this configuration does not list stay, and the user is told.
    if (-not $folderImport) {
        $workbookModules = @(
            foreach ($component in $vbProject.VBComponents) {
                if ([int]$component.Type -eq $vbext_ct_StdModule) { [string]$component.Name }
            }
        )
        $leftInPlace = @(Get-LeftInPlaceModules -WorkbookModules $workbookModules -RepositoryNames $repositoryNames -ModeNames $modeNames)
        if ($leftInPlace.Count -gt 0) {
            Write-Warning ("The workbook holds modules that the $syncMode configuration does not list; " +
                "they are left in place: " + ($leftInPlace -join ", ") +
                ". For a workbook with exactly this configuration, start from a new workbook (-NewWorkbook).")
        }
    }

    foreach ($file in $importFiles) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        Import-StandardModule -VBProject $vbProject -SourcePath $file.FullName -Name $name
    }

    $after = Get-SheetSnapshot -Workbook $workbook
    Assert-SheetSnapshotUnchanged -Before $before -After $after

    $excel.Calculation = $calculation
    $excel.CalculateBeforeSave = $calculateBeforeSave

    $workbook.Save()
    $saved = $true

    # The backup is a safety copy for this run only: it is removed after a
    # successful save, so repeated synchronizations do not accumulate files.
    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Remove-Item -LiteralPath $backupPath -Force
        Write-Host "Backup removed: $backupPath"
    }

    Write-Host ""
    Write-Host "SUCCESS: VBA synchronized. No cell, sheet or other module was changed."
}
catch {
    Write-Error $_ -ErrorAction Continue
    if ($null -ne $workbook -and -not $saved) {
        Write-Warning "Synchronization failed. The workbook is closed without saving."
    }
    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Write-Warning "Safety backup kept: $backupPath"
    }
    exit 1
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($stagePath) -and (Test-Path -LiteralPath $stagePath -PathType Container)) {
        try { Remove-Item -LiteralPath $stagePath -Recurse -Force } catch {}
    }

    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
    }

    if ($null -ne $excel) {
        if ($null -ne $restore) {
            foreach ($setting in $restore.Keys) {
                try { $excel.$setting = $restore[$setting] } catch {}
            }
        }
        try { $excel.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
