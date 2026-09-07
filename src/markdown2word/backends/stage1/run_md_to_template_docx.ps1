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

    [string]$BackendDirectory = $PSScriptRoot,
    [string]$PythonPath = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe",
    [string]$PandocPath = "$env:LOCALAPPDATA\Pandoc\pandoc.exe",
    [string]$VisioDirectory
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
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label file not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

$MarkdownPath = Resolve-InputFile -Path $MarkdownPath -Label "Markdown"
$PythonPath = Resolve-InputFile -Path $PythonPath -Label "Python"
$PandocPath = Resolve-InputFile -Path $PandocPath -Label "Pandoc"
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Resolve-InputFile -Path $TemplatePath -Label "Word template"
}

if ([System.IO.Path]::GetExtension($OutputPath) -ne ".docx") {
    $OutputPath += ".docx"
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if ($outputDirectory) {
    [void][System.IO.Directory]::CreateDirectory($outputDirectory)
}

# The stage-1 builder is packaged beside this internal backend wrapper.
$builder = Join-Path $PSScriptRoot "build_stage1_docx.py"
if (-not (Test-Path -LiteralPath $builder -PathType Leaf)) {
    $builder = Join-Path $BackendDirectory "build_stage1_docx.py"
}
$builder = Resolve-InputFile -Path $builder -Label "Stage-1 builder"

$arguments = @(
    $builder,
    $MarkdownPath
)
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    $arguments += $TemplatePath
}
$arguments += @(
    $OutputPath,
    "--pandoc",
    $PandocPath
)

if (-not [string]::IsNullOrWhiteSpace($VisioDirectory)) {
    $VisioDirectory = [System.IO.Path]::GetFullPath($VisioDirectory)
    $arguments += @("--visio-dir", $VisioDirectory)
}

Write-Host "Markdown to template DOCX" -ForegroundColor Cyan
Write-Host "Markdown : $MarkdownPath"
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    Write-Host "Template : $TemplatePath"
} else {
    Write-Host "Template : bundled default template"
}
Write-Host "Output   : $OutputPath"
Write-Host "Python   : $PythonPath"
Write-Host "Pandoc   : $PandocPath"
Write-Host "Converting..." -ForegroundColor Cyan

& $PythonPath @arguments
$exitCode = $LASTEXITCODE
if ($exitCode -ne 0) {
    throw "Stage-1 conversion failed with exit code $exitCode."
}
if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw "Converter returned successfully but output was not created: $OutputPath"
}

Write-Host "Success: $OutputPath" -ForegroundColor Green
[pscustomobject]@{
    MarkdownPath = $MarkdownPath
    TemplatePath = if ($TemplatePath) { $TemplatePath } else { "bundled default template" }
    OutputPath = $OutputPath
    VisioSources = if ($VisioDirectory) { $VisioDirectory } else { [System.IO.Path]::Combine($outputDirectory, [System.IO.Path]::GetFileNameWithoutExtension($OutputPath) + "_visio_sources") }
}
