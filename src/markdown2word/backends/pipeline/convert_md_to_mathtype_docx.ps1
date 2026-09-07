[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [Alias("Markdown", "Md")]
    [string]$MarkdownPath,

    [Parameter(Mandatory = $false)]
    [Alias("Template")]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [Alias("Output")]
    [string]$OutputPath,

    [string]$Stage1Path,
    [string]$VisioDirectory,
    [switch]$WriteVisioSources,
    [string]$Stage1BackendDirectory,
    [string]$MathTypeBackendDirectory,
    [string]$PythonPath = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe",
    [string]$PandocPath = "$env:LOCALAPPDATA\Pandoc\pandoc.exe",
    [switch]$ExternalInlineDialogGuard,
    [switch]$AddEquationNumbers,
    [ValidateSet("Simple", "ChapterDash", "ChapterDot")]
    [string]$EquationNumberFormat = "Simple",
    [switch]$StartWordIfNeeded,
    [switch]$SkipTextIntegrityCheck,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$env:PYTHONIOENCODING = "utf-8"
Set-StrictMode -Version Latest

function Resolve-InputFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label file not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-InputDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "$Label directory not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Normalize-DocxPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([System.IO.Path]::GetExtension($Path) -ne ".docx") {
        $Path += ".docx"
    }
    return [System.IO.Path]::GetFullPath($Path)
}

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "== $Message ==" -ForegroundColor Cyan
}

$backendRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Stage1BackendDirectory)) {
    $Stage1BackendDirectory = Join-Path $backendRoot "stage1"
}
if ([string]::IsNullOrWhiteSpace($MathTypeBackendDirectory)) {
    $MathTypeBackendDirectory = Join-Path $backendRoot "mathtype"
}

$MarkdownPath = Resolve-InputFile -Path $MarkdownPath -Label "Markdown"
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Resolve-InputFile -Path $TemplatePath -Label "Word template"
}
$PythonPath = Resolve-InputFile -Path $PythonPath -Label "Python"
$PandocPath = Resolve-InputFile -Path $PandocPath -Label "Pandoc"
$Stage1BackendDirectory = Resolve-InputDirectory -Path $Stage1BackendDirectory -Label "stage-1 backend"
$MathTypeBackendDirectory = Resolve-InputDirectory -Path $MathTypeBackendDirectory -Label "MathType backend"

$stage1Runner = Resolve-InputFile -Path (Join-Path $Stage1BackendDirectory "run_md_to_template_docx.ps1") -Label "Stage-1 runner"
$stage1Audit = Resolve-InputFile -Path (Join-Path $Stage1BackendDirectory "audit_stage1_docx.py") -Label "Stage-1 audit"
$mathTypeRunner = Resolve-InputFile -Path (Join-Path $MathTypeBackendDirectory "convert_word_latex_to_mathtype.ps1") -Label "MathType converter"

$OutputPath = Normalize-DocxPath -Path $OutputPath
$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if ($outputDirectory) {
    [void][System.IO.Directory]::CreateDirectory($outputDirectory)
}
if ([string]::IsNullOrWhiteSpace($Stage1Path)) {
    $Stage1Path = Join-Path $outputDirectory ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath) + "_stage1_latex.docx")
} else {
    $Stage1Path = Normalize-DocxPath -Path $Stage1Path
}
if ([string]::IsNullOrWhiteSpace($VisioDirectory)) {
    $VisioDirectory = Join-Path $outputDirectory ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath) + "_visio_sources")
} else {
    $VisioDirectory = [System.IO.Path]::GetFullPath($VisioDirectory)
}
$mathTypeOnlyPath = $OutputPath
$stageCount = 2

Write-Host "Markdown template MathType pipeline" -ForegroundColor Green
Write-Host "Markdown      : $MarkdownPath"
Write-Host "Template      : $(if ($TemplatePath) { $TemplatePath } else { 'bundled default template' })"
Write-Host "Stage-1 DOCX  : $Stage1Path"
Write-Host "Final DOCX    : $OutputPath"
Write-Host "Visio sources : $(if ($WriteVisioSources) { $VisioDirectory } else { 'disabled (placeholder only)' })"

Write-Step "Stage 1/${stageCount}: Markdown to template DOCX"
$stage1Args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $stage1Runner,
    "-MarkdownPath", $MarkdownPath,
    "-OutputPath", $Stage1Path,
    "-BackendDirectory", $Stage1BackendDirectory,
    "-PythonPath", $PythonPath,
    "-PandocPath", $PandocPath
)
if ($TemplatePath) {
    $stage1Args += @("-TemplatePath", $TemplatePath)
}
if ($WriteVisioSources) {
    $stage1Args += @("-WriteVisioSources", "-VisioDirectory", $VisioDirectory)
}
& powershell @stage1Args
if ($LASTEXITCODE -ne 0) {
    throw "Stage 1 failed with exit code $LASTEXITCODE."
}
if (-not (Test-Path -LiteralPath $Stage1Path -PathType Leaf)) {
    throw "Stage 1 finished but did not create: $Stage1Path"
}

$auditJson = (& $PythonPath $stage1Audit $Stage1Path | Out-String)
if ($LASTEXITCODE -ne 0) {
    throw "Stage-1 audit failed with exit code $LASTEXITCODE. Stage-1 DOCX retained: $Stage1Path"
}
try {
    $stage1Summary = $auditJson | ConvertFrom-Json
} catch {
    throw "Stage-1 audit returned invalid JSON. Stage-1 DOCX retained: $Stage1Path. Error=$($_.Exception.Message)"
}
$formulaCount = [int]$stage1Summary.inline_latex + [int]$stage1Summary.display_latex

if ($formulaCount -eq 0) {
    Write-Step "Stage 2/${stageCount}: No LaTeX formulas; MathType skipped"
    if ($Stage1Path -ne $mathTypeOnlyPath) {
        if ((Test-Path -LiteralPath $mathTypeOnlyPath) -and (-not $Force)) {
            throw "Output already exists: $mathTypeOnlyPath. Use -Force to replace it."
        }
        Copy-Item -LiteralPath $Stage1Path -Destination $mathTypeOnlyPath -Force
    }
    Write-Host "No formulas found. Stage-1 DOCX promoted to final output; subsequent Visio processing can continue." -ForegroundColor Yellow
} else {
    Write-Step "Stage 2/${stageCount}: Word LaTeX to MathType"
    $stage2Args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $mathTypeRunner,
        "-InputPath", $Stage1Path,
        "-OutputPath", $mathTypeOnlyPath
    )
    if ($Force) { $stage2Args += "-Force" }
    if ($SkipTextIntegrityCheck) { $stage2Args += "-SkipTextIntegrityCheck" }
    if ($StartWordIfNeeded) { $stage2Args += "-StartWordIfNeeded" }
    if ($ExternalInlineDialogGuard) { $stage2Args += "-ExternalInlineDialogGuard" }
    if ($AddEquationNumbers) {
        $stage2Args += @("-AddEquationNumbers", "-EquationNumberFormat", $EquationNumberFormat)
    }

    & powershell @stage2Args
    if ($LASTEXITCODE -ne 0) {
        throw "Stage 2 failed with exit code $LASTEXITCODE. Stage-1 DOCX retained: $Stage1Path"
    }
    if (-not (Test-Path -LiteralPath $mathTypeOnlyPath -PathType Leaf)) {
        throw "Stage 2 finished but did not create: $mathTypeOnlyPath"
    }
}

Write-Host ""
Write-Host "Pipeline success" -ForegroundColor Green
[pscustomobject]@{
    MarkdownPath = $MarkdownPath
    TemplatePath = if ($TemplatePath) { $TemplatePath } else { "bundled default template" }
    Stage1Path = $Stage1Path
    MathTypeOnlyPath = $mathTypeOnlyPath
    OutputPath = $OutputPath
    VisioSources = if ($WriteVisioSources) { $VisioDirectory } else { $null }
}
