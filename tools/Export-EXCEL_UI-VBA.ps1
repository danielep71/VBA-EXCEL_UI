<#
.SYNOPSIS
Export every VBA component of a workbook to a new folder, for comparison
with the repository source.

.DESCRIPTION
Opens the workbook read-only in an Excel instance of its own, with events and
macros disabled, exports every VBA component with the VBE's own Export
(standard modules as .bas, class and document modules as .cls, forms as .frm
with their .frx), and closes the workbook without saving. The folder must be
new or empty, so that an earlier export is never mixed in.

Beside the components the script writes export-manifest.json: the file name
and SHA-256 of the workbook, hashed before it is opened and checked again
after it is closed, and the VB_Name, type, file and SHA-256 of every
exported file. tools/compare_vba_export.py compares the export with the
source of one commit and, with --workbook, checks the workbook and the files
against the manifest. The manifest is a record, not a signature.

Exit codes: 0 success, 1 error.

.REQUIREMENTS
- Windows desktop Excel.
- Excel > Options > Trust Center > Macro Settings >
  "Trust access to the VBA project object model" enabled.
- VBA project not password-locked.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Export-EXCEL_UI-VBA.ps1 -WorkbookPath ".\test-results\EXCEL_UI_Development.xlsm" -OutDir ".\test-results\export"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$msoAutomationSecurityForceDisable = 3
$ExportManifestName = "export-manifest.json"
$ExportFormat = "excel-ui-vba-export-1"

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-ExportKind {
    # The file extension and the type name of a component type, or $null
    # for a type the VBE does not export.
    param([Parameter(Mandatory = $true)][int]$ComponentType)

    switch ($ComponentType) {
        1 { return [pscustomobject]@{ Extension = ".bas"; Type = "standard" } }
        2 { return [pscustomobject]@{ Extension = ".cls"; Type = "class" } }
        3 { return [pscustomobject]@{ Extension = ".frm"; Type = "form" } }
        100 { return [pscustomobject]@{ Extension = ".cls"; Type = "document" } }
    }
    return $null
}

$excel = $null
$workbook = $null

try {
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        throw "Workbook not found: $WorkbookPath"
    }
    $workbookFullPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $WorkbookPath).Path)
    $outFullPath = [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir))

    if (Test-Path -LiteralPath $outFullPath -PathType Container) {
        if (@(Get-ChildItem -LiteralPath $outFullPath -Force).Count -gt 0) {
            throw "OutDir is not empty: $outFullPath. Use a new folder."
        }
    }
    elseif (Test-Path -LiteralPath $outFullPath) {
        throw "OutDir is a file: $outFullPath. Use a new folder."
    }
    else {
        $null = New-Item -ItemType Directory -Path $outFullPath
    }

    $workbookHash = Get-Sha256 -Path $workbookFullPath

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Read-only, links not updated: the export never changes the workbook.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $true)

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw (
            "Cannot access the VBA project. Enable Excel > Options > Trust Center > Trust Center Settings > " +
            "Macro Settings > 'Trust access to the VBA project object model', and check that the VBA " +
            "project is not password-locked."
        )
    }

    $count = 0
    $components = New-Object System.Collections.Generic.List[object]
    foreach ($component in $vbProject.VBComponents) {
        $kind = Get-ExportKind -ComponentType $component.Type
        if ($null -eq $kind) {
            Write-Warning "Not exported: $($component.Name) (component type $($component.Type))"
            continue
        }
        $file = $component.Name + $kind.Extension
        $target = Join-Path $outFullPath $file
        $component.Export($target)
        $components.Add([ordered]@{
                name   = [string]$component.Name
                type   = $kind.Type
                file   = $file
                sha256 = Get-Sha256 -Path $target
            })
        # A form's binary properties go to a .frx beside its .frm.
        $resource = [System.IO.Path]::ChangeExtension($target, ".frx")
        if ($kind.Type -eq "form" -and (Test-Path -LiteralPath $resource -PathType Leaf)) {
            $components.Add([ordered]@{
                    name   = [string]$component.Name
                    type   = "form-resource"
                    file   = $component.Name + ".frx"
                    sha256 = Get-Sha256 -Path $resource
                })
        }
        Write-Host ("  {0,-32} {1}" -f $file, $kind.Type)
        $count++
    }

    # Close before hashing again: the manifest records the workbook only if
    # the export left it unchanged.
    $workbook.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook = $null
    if ((Get-Sha256 -Path $workbookFullPath) -ne $workbookHash) {
        throw "The workbook changed while it was exported: $workbookFullPath. Export it again."
    }

    $manifest = [ordered]@{
        format     = $ExportFormat
        workbook   = [ordered]@{
            file   = [System.IO.Path]::GetFileName($workbookFullPath)
            sha256 = $workbookHash
        }
        components = $components.ToArray()
    }
    [System.IO.File]::WriteAllText(
        (Join-Path $outFullPath $ExportManifestName),
        ($manifest | ConvertTo-Json -Depth 4),
        (New-Object System.Text.UTF8Encoding($false)))

    Write-Host ""
    Write-Host "Workbook: $workbookFullPath"
    Write-Host "SHA-256:  $workbookHash"
    Write-Host "Exported: $count component(s) to $outFullPath, with $ExportManifestName"
    Write-Host "Next:     python tools/compare_vba_export.py --export `"$outFullPath`" --mode production|development|demo --workbook `"$workbookFullPath`""
}
catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
