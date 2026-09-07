[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$VisioDirectory,
    [switch]$StartWordIfNeeded
)

$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$env:PYTHONIOENCODING = "utf-8"
Set-StrictMode -Version Latest
. (Join-Path (Split-Path -Parent $PSScriptRoot) "office_session.ps1")
$wdDoNotSaveChanges = 0
$wdFindStop = 0

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-WordApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    } catch {
        if (-not $StartWordIfNeeded) {
            throw "Microsoft Word is not open. Enable automatic Word startup or open Word manually."
        }
        $word = Connect-DesktopWord -StartIfNeeded
        $word.Visible = $true
        return $word
    }
}

function Open-BasicFlowchartStencil {
    param([Parameter(Mandatory = $true)]$Visio)
    foreach ($name in @("BASFLO_U.VSSX", "BASFLO_M.VSSX", "SSFLOW_M.VSSX")) {
        try { return $Visio.Documents.OpenEx($name, 66) } catch {}
    }
    throw "Visio Basic Flowchart stencil was not found (BASFLO_U/M or SSFLOW_M)."
}

function Parse-MermaidFlowchart {
    param([Parameter(Mandatory = $true)][string]$Source)
    $lines = @($Source -split "`r?`n")
    if ($lines.Count -eq 0 -or $lines[0] -notmatch '^\s*(?:flowchart|graph)\s+(?<direction>TD|TB|BT|LR|RL)\s*$') {
        return $null
    }
    $direction = $Matches.direction
    $nodes = [ordered]@{}
    $edges = [System.Collections.Generic.List[object]]::new()
    $nodePattern = '(?<id>[\p{L}\p{N}_-]+)\s*(?<shape>\[[^\]]+\]|\{[^}]+\}|\([^\)]+\))?'
    $edgePattern = '^\s*(?<from>[\p{L}\p{N}_-]+)\s*(?<fromShape>\[[^\]]+\]|\{[^}]+\}|\([^\)]+\))?\s*-->(?:\|(?<label>[^|]+)\|)?\s*(?<to>[\p{L}\p{N}_-]+)\s*(?<toShape>\[[^\]]+\]|\{[^}]+\}|\([^\)]+\))?\s*$'

    foreach ($line in $lines | Select-Object -Skip 1) {
        $clean = ($line -replace '%%.*$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($clean)) { continue }
        $edgeMatch = [regex]::Match($clean, $edgePattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($edgeMatch.Success) {
            $ids = @($edgeMatch.Groups['from'].Value, $edgeMatch.Groups['to'].Value)
            $shapes = @($edgeMatch.Groups['fromShape'].Value, $edgeMatch.Groups['toShape'].Value)
            for ($i = 0; $i -lt 2; $i++) {
                $shapeText = $shapes[$i]
                if (-not $nodes.Contains($ids[$i])) {
                    $nodes[$ids[$i]] = [pscustomobject]@{ Id=$ids[$i]; Shape=$shapeText; Text=$ids[$i] }
                } elseif (-not [string]::IsNullOrWhiteSpace($shapeText)) {
                    $nodes[$ids[$i]].Shape = $shapeText
                    $nodes[$ids[$i]].Text = $shapeText.Trim('[',']','{','}','(',')')
                }
            }
            $edges.Add([pscustomobject]@{ From=$ids[0]; To=$ids[1]; Label=$edgeMatch.Groups['label'].Value })
            continue
        }
        $nodeMatch = [regex]::Match($clean, "^\s*$nodePattern\s*`$")
        if ($nodeMatch.Success) {
            $id = $nodeMatch.Groups['id'].Value
            $shapeText = $nodeMatch.Groups['shape'].Value
            $text = if ($shapeText) { $shapeText.Trim('[',']','{','}','(',')') } else { $id }
            $nodes[$id] = [pscustomobject]@{ Id=$id; Shape=$shapeText; Text=$text }
            continue
        }
        throw "Unsupported Mermaid flowchart statement: $clean"
    }
    if ($nodes.Count -eq 0) { throw "The Mermaid flowchart contains no nodes." }
    return [pscustomobject]@{ Direction=$direction; Nodes=@($nodes.Values); Edges=@($edges) }
}

function New-VisioDiagram {
    param(
        [Parameter(Mandatory = $true)]$Visio,
        [Parameter(Mandatory = $true)]$Stencil,
        [Parameter(Mandatory = $true)]$Model,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )
    $document = $Visio.Documents.Add("")
    $page = $document.Pages.Item(1)
    $processMaster = $Stencil.Masters.ItemU("Process")
    $decisionMaster = $Stencil.Masters.ItemU("Decision")
    $startMaster = $Stencil.Masters.ItemU("Start/End")
    $connectorMaster = $Stencil.Masters.ItemU("Dynamic connector")
    $shapes = @{}
    try {
        for ($i = 0; $i -lt $Model.Nodes.Count; $i++) {
            $node = $Model.Nodes[$i]
            $master = if ($node.Shape.StartsWith('{')) { $decisionMaster } elseif ($node.Shape.StartsWith('(')) { $startMaster } else { $processMaster }
            if ($Model.Direction -in @('LR', 'RL')) {
                $x = 1.5 + ($i * 2.4)
                $y = 5.0
                if ($Model.Direction -eq 'RL') { $x = 1.5 + (($Model.Nodes.Count - 1 - $i) * 2.4) }
            } else {
                $x = 4.25
                $y = 10.0 - ($i * 1.35)
                if ($Model.Direction -eq 'BT') { $y = 2.0 + ($i * 1.35) }
            }
            $shape = $page.Drop($master, $x, $y)
            $shape.Text = [string]$node.Text
            $shapes[$node.Id] = $shape
        }
        foreach ($edge in $Model.Edges) {
            $connector = $page.Drop($connectorMaster, 0, 0)
            $connector.CellsU("BeginX").GlueToPos($shapes[$edge.From], 0.5, 0.5)
            $connector.CellsU("EndX").GlueToPos($shapes[$edge.To], 0.5, 0.5)
            if ($connector.CellExistsU("ConLineRouteExt", 0)) { $connector.CellsU("ConLineRouteExt").FormulaU = "0" }
            if ($edge.Label) { $connector.Text = [string]$edge.Label }
            Release-ComObject $connector
        }
        $page.ResizeToFitContents()
        $document.SaveAs($OutputPath)
    } finally {
        foreach ($shape in $shapes.Values) { Release-ComObject $shape }
        Release-ComObject $connectorMaster
        Release-ComObject $startMaster
        Release-ComObject $decisionMaster
        Release-ComObject $processMaster
        Release-ComObject $page
        try { $document.Close() } catch {}
        Release-ComObject $document
    }
}

$DocxPath = [IO.Path]::GetFullPath($DocxPath)
$VisioDirectory = [IO.Path]::GetFullPath($VisioDirectory)
$manifestPath = Join-Path $VisioDirectory "manifest.json"
if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) { throw "DOCX not found: $DocxPath" }
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Mermaid manifest not found: $manifestPath" }
$manifest = @(Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json)
if ($manifest.Count -eq 0) {
    Write-Host "No Mermaid diagrams were found; Visio stage skipped." -ForegroundColor Yellow
    exit 0
}

$visio = $null
$stencil = $null
$word = $null
$wordDocument = $null
$created = [System.Collections.Generic.List[object]]::new()
try {
    $visio = New-Object -ComObject Visio.InvisibleApp
    $stencil = Open-BasicFlowchartStencil -Visio $visio
    foreach ($item in $manifest) {
        $sourcePath = Join-Path $VisioDirectory ([string]$item.source)
        $source = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
        $model = Parse-MermaidFlowchart -Source $source
        if ($null -eq $model) {
            Write-Warning "Unsupported Mermaid type '$($item.diagram_type)' for $($item.id); placeholder retained. Currently supported: flowchart/graph."
            continue
        }
        $vsdxPath = Join-Path $VisioDirectory ("{0}.vsdx" -f $item.id)
        New-VisioDiagram -Visio $visio -Stencil $stencil -Model $model -OutputPath $vsdxPath
        $created.Add([pscustomobject]@{ Placeholder=[string]$item.placeholder; Path=$vsdxPath; Id=[string]$item.id })
        Write-Host "Created editable Visio diagram: $vsdxPath" -ForegroundColor Green
    }

    if ($created.Count -eq 0) {
        Write-Warning "No supported Mermaid flowcharts were converted. The Word placeholders were retained."
        exit 0
    }
    $word = Get-WordApplication
    $word.Visible = $true
    $wordDocument = $word.Documents.Open($DocxPath, $false, $false, $false)
    foreach ($item in $created) {
        $range = $wordDocument.Content.Duplicate
        $find = $range.Find
        try {
            $find.ClearFormatting()
            $find.Text = $item.Placeholder
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            if (-not $find.Execute()) { throw "Word placeholder not found: $($item.Placeholder)" }
            $range.Text = ""
            $ole = $wordDocument.InlineShapes.AddOLEObject("Visio.Drawing", $item.Path, $false, $false, "", 0, "", $range)
            try {
                if ([double]$ole.Width -gt 430.0) {
                    $ole.LockAspectRatio = -1
                    $ole.Width = 430.0
                }
            } finally { Release-ComObject $ole }
            Write-Host "Embedded editable Visio object: $($item.Id)" -ForegroundColor Green
        } finally {
            Release-ComObject $find
            Release-ComObject $range
        }
    }
    $wordDocument.Save()
    Write-Host "Visio embedding succeeded: $DocxPath" -ForegroundColor Green
} finally {
    if ($null -ne $wordDocument) { try { $wordDocument.Close($wdDoNotSaveChanges) } catch {}; Release-ComObject $wordDocument }
    if ($null -ne $word) { Release-ComObject $word }
    if ($null -ne $stencil) { try { $stencil.Close() } catch {}; Release-ComObject $stencil }
    if ($null -ne $visio) { try { $visio.Quit() } catch {}; Release-ComObject $visio }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
