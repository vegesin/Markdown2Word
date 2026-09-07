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
$wdCollapseEnd = 0
$rpcRejected = @(-2147418111, -2147417846)

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Invoke-ComWithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [int]$Retries = 12,
        [int]$DelayMilliseconds = 350
    )
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return & $Operation
        } catch [Runtime.InteropServices.COMException] {
            if ($rpcRejected -notcontains $_.Exception.HResult -or $attempt -eq $Retries) { throw }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Format-ComFailure {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)]$Exception
    )
    while ($null -ne $Exception.InnerException) { $Exception = $Exception.InnerException }
    $hresult = "0x{0:X8}" -f (([int64]$Exception.HResult) -band 0xffffffffL)
    return "Visio automation failed during '$Stage'. HRESULT=$hresult; Error=$($Exception.Message)"
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
    # Accept both common Mermaid edge-label forms:
    #   A -->|yes| B
    #   A -- yes --> B
    $edgePattern = '^\s*(?<from>[\p{L}\p{N}_-]+)\s*(?<fromShape>\[[^\]]+\]|\{[^}]+\}|\([^\)]+\))?\s*(?:-->\s*(?:\|(?<label>[^|]+)\|\s*)?|--\s+(?<label>.+?)\s+-->\s*)(?<to>[\p{L}\p{N}_-]+)\s*(?<toShape>\[[^\]]+\]|\{[^}]+\}|\([^\)]+\))?\s*$'

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

function Get-MermaidTopologyPositions {
    param([Parameter(Mandatory = $true)]$Model)
    $incoming = @{}
    $outgoing = @{}
    foreach ($node in $Model.Nodes) {
        $incoming[$node.Id] = 0
        $outgoing[$node.Id] = [System.Collections.Generic.List[string]]::new()
    }
    foreach ($edge in $Model.Edges) {
        $incoming[$edge.To] = [int]$incoming[$edge.To] + 1
        $outgoing[$edge.From].Add([string]$edge.To)
    }

    $depth = @{}
    $queue = [System.Collections.Generic.Queue[string]]::new()
    foreach ($node in $Model.Nodes) {
        if ([int]$incoming[$node.Id] -eq 0) {
            $depth[$node.Id] = 0
            $queue.Enqueue([string]$node.Id)
        }
    }
    if ($queue.Count -eq 0 -and $Model.Nodes.Count -gt 0) {
        $firstId = [string]$Model.Nodes[0].Id
        $depth[$firstId] = 0
        $queue.Enqueue($firstId)
    }
    while ($queue.Count -gt 0) {
        $from = $queue.Dequeue()
        foreach ($to in $outgoing[$from]) {
            # First visit gives the shortest forward level. Edges to an
            # existing level are feedback/merge edges and must not create an
            # infinite depth increase in cyclic Mermaid graphs.
            if (-not $depth.ContainsKey($to)) {
                $depth[$to] = [int]$depth[$from] + 1
                $queue.Enqueue($to)
            }
        }
    }
    $lastDepth = if ($depth.Count) { ($depth.Values | Measure-Object -Maximum).Maximum } else { 0 }
    foreach ($node in $Model.Nodes) {
        if (-not $depth.ContainsKey($node.Id)) {
            $lastDepth++
            $depth[$node.Id] = $lastDepth
        }
    }

    $levels = @{}
    foreach ($node in $Model.Nodes) {
        $level = [int]$depth[$node.Id]
        if (-not $levels.ContainsKey($level)) { $levels[$level] = [System.Collections.Generic.List[string]]::new() }
        $levels[$level].Add([string]$node.Id)
    }
    $maxLevel = ($levels.Keys | Measure-Object -Maximum).Maximum
    $maxWidth = ($levels.Values | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $levelGap = 1.65
    $crossGap = 2.65
    $center = 1.5 + (($maxWidth - 1) * $crossGap / 2.0)
    $positions = @{}
    foreach ($level in ($levels.Keys | Sort-Object)) {
        $ids = $levels[$level]
        for ($index = 0; $index -lt $ids.Count; $index++) {
            $cross = $center + (($index - (($ids.Count - 1) / 2.0)) * $crossGap)
            $main = 1.25 + (($maxLevel - $level) * $levelGap)
            switch ($Model.Direction) {
                'BT' { $x = $cross; $y = 1.25 + ($level * $levelGap) }
                'LR' { $x = 1.25 + ($level * $levelGap); $y = $cross }
                'RL' { $x = 1.25 + (($maxLevel - $level) * $levelGap); $y = $cross }
                default { $x = $cross; $y = $main }
            }
            $positions[$ids[$index]] = [pscustomobject]@{ X=[double]$x; Y=[double]$y }
        }
    }
    return ,$positions
}

function Set-VisioCellFormulaIfPresent {
    param(
        [Parameter(Mandatory = $true)]$Shape,
        [Parameter(Mandatory = $true)][string]$CellName,
        [Parameter(Mandatory = $true)][string]$Formula
    )
    if ($Shape.CellExistsU($CellName, 0)) { $Shape.CellsU($CellName).FormulaU = $Formula }
}

function Set-MermaidPageLayout {
    param(
        [Parameter(Mandatory = $true)]$Page,
        [Parameter(Mandatory = $true)][string]$Direction
    )
    $placeStyle = switch ($Direction) {
        'BT' { 20 } # hierarchy bottom-to-top, centered
        'LR' { 23 } # hierarchy left-to-right, middle
        'RL' { 26 } # hierarchy right-to-left, middle
        default { 17 } # hierarchy top-to-bottom, centered
    }
    $routeStyle = switch ($Direction) {
        'BT' { 12 } # flowchart bottom-to-top
        'LR' { 6 }  # flowchart left-to-right
        'RL' { 13 } # flowchart right-to-left
        default { 5 } # flowchart top-to-bottom
    }
    $pageSheet = $Page.PageSheet
    try {
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'PlaceStyle' -Formula ([string]$placeStyle)
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'RouteStyle' -Formula ([string]$routeStyle)
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'PageShapeSplit' -Formula '0'
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'LineToNodeX' -Formula '0.25 in'
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'LineToNodeY' -Formula '0.25 in'
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'LineToLineX' -Formula '0.18 in'
        Set-VisioCellFormulaIfPresent -Shape $pageSheet -CellName 'LineToLineY' -Formula '0.18 in'
    } finally {
        Release-ComObject $pageSheet
    }
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
        $positions = Get-MermaidTopologyPositions -Model $Model
        Set-MermaidPageLayout -Page $page -Direction $Model.Direction
        for ($i = 0; $i -lt $Model.Nodes.Count; $i++) {
            $node = $Model.Nodes[$i]
            $master = if ($node.Shape.StartsWith('{')) { $decisionMaster } elseif ($node.Shape.StartsWith('(')) { $startMaster } else { $processMaster }
            $position = $positions[$node.Id]
            $shape = $page.Drop($master, $position.X, $position.Y)
            if ($node.Shape.StartsWith('{')) {
                $shape.CellsU('Width').FormulaU = '2.15 in'
                $shape.CellsU('Height').FormulaU = '1.15 in'
            } elseif ($node.Shape.StartsWith('(')) {
                $shape.CellsU('Width').FormulaU = '2.35 in'
                $shape.CellsU('Height').FormulaU = '0.75 in'
            } else {
                $shape.CellsU('Width').FormulaU = '2.35 in'
                $shape.CellsU('Height').FormulaU = '0.85 in'
            }
            $shape.Text = [string]$node.Text
            Set-VisioCellFormulaIfPresent -Shape $shape -CellName 'Char.Size' -Formula '10 pt'
            Set-VisioCellFormulaIfPresent -Shape $shape -CellName 'Para.HorzAlign' -Formula '1'
            Set-VisioCellFormulaIfPresent -Shape $shape -CellName 'VerticalAlign' -Formula '1'
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
        # Use Visio's native hierarchical layout after the topology-aware
        # initial placement. This is the closest editable equivalent to
        # Mermaid's layered graph rendering and reroutes feedback edges around
        # nodes instead of through them.
        $page.Layout()
        $null = $page.ResizeToFitContents()
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
$stage = "start background Visio"
try {
    $visio = New-Object -ComObject Visio.InvisibleApp
    $stage = "open Basic Flowchart stencil"
    $stencil = Open-BasicFlowchartStencil -Visio $visio
    foreach ($item in $manifest) {
        $stage = "parse Mermaid source $($item.id)"
        $sourcePath = Join-Path $VisioDirectory ([string]$item.source)
        $source = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
        $model = Parse-MermaidFlowchart -Source $source
        if ($null -eq $model) {
            Write-Warning "Unsupported Mermaid type '$($item.diagram_type)' for $($item.id); placeholder retained. Currently supported: flowchart/graph."
            continue
        }
        $stage = "create editable VSDX $($item.id)"
        $vsdxPath = Join-Path $VisioDirectory ("{0}.vsdx" -f $item.id)
        New-VisioDiagram -Visio $visio -Stencil $stencil -Model $model -OutputPath $vsdxPath
        $created.Add([pscustomobject]@{ Placeholder=[string]$item.placeholder; Path=$vsdxPath; Id=[string]$item.id })
        Write-Host "Created editable Visio diagram: $vsdxPath" -ForegroundColor Green
    }

    if ($created.Count -eq 0) {
        Write-Warning "No supported Mermaid flowcharts were converted. The Word placeholders were retained."
        exit 0
    }
    $stage = "connect to Microsoft Word"
    $word = Get-WordApplication
    $word.Visible = $true
    $stage = "open target Word document"
    $wordDocument = Invoke-ComWithRetry -Operation { $word.Documents.Open($DocxPath, $false, $false, $false) }
    foreach ($item in $created) {
        $stage = "find Word placeholder $($item.Placeholder)"
        $range = $wordDocument.Content.Duplicate
        $find = $range.Find
        try {
            $find.ClearFormatting()
            $find.Text = $item.Placeholder
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            if (-not $find.Execute()) { throw "Word placeholder not found: $($item.Placeholder)" }
            $range.Text = ""
            $range.Collapse($wdCollapseEnd)
            $stage = "embed editable Visio object $($item.Id) into Word"
            # Word requires either ClassType or FileName, never both.  Because
            # this embeds an existing VSDX, omit ClassType and pass FileName.
            $ole = Invoke-ComWithRetry -Operation {
                $wordDocument.InlineShapes.AddOLEObject(
                    [Type]::Missing,
                    $item.Path,
                    $false,
                    $false,
                    [Type]::Missing,
                    [Type]::Missing,
                    [Type]::Missing,
                    $range
                )
            }
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
    $stage = "save Word document with embedded Visio objects"
    $wordDocument.Save()
    Write-Host "Visio embedding succeeded: $DocxPath" -ForegroundColor Green
} catch {
    throw (Format-ComFailure -Stage $stage -Exception $_.Exception)
} finally {
    if ($null -ne $wordDocument) { try { $wordDocument.Close($wdDoNotSaveChanges) } catch {}; Release-ComObject $wordDocument }
    if ($null -ne $word) { Release-ComObject $word }
    if ($null -ne $stencil) { try { $stencil.Close() } catch {}; Release-ComObject $stencil }
    if ($null -ne $visio) { try { $visio.Quit() } catch {}; Release-ComObject $visio }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
