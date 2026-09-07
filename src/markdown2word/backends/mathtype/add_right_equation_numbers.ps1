[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [string]$OutputPath,

    [ValidateRange(1, 1000000)]
    [int]$StartNumber = 1,

    [ValidateRange(-1, 1000000)]
    [int]$ExpectedDisplay = -1,

    [ValidateSet("Simple", "ChapterDash", "ChapterDot")]
    [string]$EquationNumberFormat = "Simple",

    [switch]$StartWordIfNeeded,

    [string]$InsertNumberMacroName = "MTCommand_InsertEqnNum",

    [string]$UpdateNumbersMacroName = "MTCommand_UpdateEqns",

    [string]$InitializeNumbersMacroName = "",

    [switch]$Force
)

$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$env:PYTHONIOENCODING = "utf-8"
Set-StrictMode -Version Latest
. (Join-Path (Split-Path -Parent $PSScriptRoot) "office_session.ps1")

$wdDoNotSaveChanges = 0
$wdCollapseEnd = 0
$wdAlignParagraphLeft = 0
$wdAlignTabCenter = 1
$wdAlignTabRight = 2
$wdTabLeaderSpaces = 0
$wdLineSpaceAtLeast = 3
$rpcRejected = @(-2147418111, -2147417846)
$equationProgId = "Equation.DSMT4"
$displayStyleName = "MTDisplayEquation"

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [int]$Retries = 12,
        [int]$DelayMilliseconds = 350
    )
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return & $Operation
        } catch [System.Runtime.InteropServices.COMException] {
            if ($rpcRejected -notcontains $_.Exception.HResult -or $attempt -eq $Retries) {
                throw
            }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Write-UserActionError {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Reason,
        [string[]]$Actions = @()
    )
    Write-Host ""
    Write-Host "================ Equation numbering stopped ================" -ForegroundColor Red
    Write-Host $Title -ForegroundColor Red
    Write-Host "原因 / Reason: $Reason" -ForegroundColor Yellow
    if ($Actions.Count -gt 0) {
        Write-Host "处理方法 / Actions:" -ForegroundColor Cyan
        foreach ($action in $Actions) {
            Write-Host "  - $action" -ForegroundColor Cyan
        }
    }
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host ""
}

function Test-FileClosed {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label,
        [System.IO.FileAccess]$Access = [System.IO.FileAccess]::Read
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    $stream = $null
    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            $Access,
            [System.IO.FileShare]::None
        )
    } catch {
        throw "$Label 文件可能已在 Word/WPS 中打开、被其他程序占用，或当前账户无访问权限。Close the file in Word/WPS and other programs, then retry. Path=$Path; Error=$($_.Exception.Message)"
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-RunningMicrosoftWord {
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    } catch {
        if ($StartWordIfNeeded) {
            try {
                $word = Connect-DesktopWord -StartIfNeeded
                $word.Visible = $true
                $blankDocument = $word.Documents.Add()
                Release-ComObject $blankDocument
                Write-Host "Microsoft Word was started automatically." -ForegroundColor Cyan
                return $word
            } catch {
                throw "Microsoft Word could not be started automatically. Confirm desktop Word is installed and activated, or start it manually at the same privilege level. Error=$($_.Exception.Message)"
            }
        }
        $wpsProcesses = @(Get-Process -Name "wps", "wpscloudsvr" -ErrorAction SilentlyContinue)
        if ($wpsProcesses.Count -gt 0) {
            throw "检测到 WPS，但没有可连接的 Microsoft Word。Close WPS, manually open Microsoft Word, keep one blank window open, and rerun PowerShell at the same privilege level."
        }
        throw "未找到已打开的 Microsoft Word，或 Word 与 PowerShell 权限级别不一致。Manually open Microsoft Word, keep one blank window open, and run Word and PowerShell both normally or both elevated. The script will not start Word automatically."
    }

    try {
        $name = [string]$word.Name
        $path = [string]$word.Path
        if ($name -notlike "*Microsoft Word*" -or $path -match "(?i)WPS|Kingsoft") {
            throw "检测到的 Word.Application 不是 Microsoft Word: Name=$name; Path=$path. Close WPS/Kingsoft and open Microsoft Word."
        }
        return $word
    } catch {
        Release-ComObject $word
        throw
    }
}

function Assert-TargetFilesNotOpenInWord {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)][string[]]$Paths
    )
    for ($i = 1; $i -le $Word.Documents.Count; $i++) {
        $openDocument = $Word.Documents.Item($i)
        try {
            $openPath = ""
            try { $openPath = [System.IO.Path]::GetFullPath([string]$openDocument.FullName) } catch {}
            foreach ($path in $Paths) {
                if (-not [string]::IsNullOrWhiteSpace($openPath) -and
                    [System.StringComparer]::OrdinalIgnoreCase.Equals($openPath, $path)) {
                    throw "输入或输出 DOCX 已在 Microsoft Word 中打开。Close this document but keep the blank Word window open: $path"
                }
            }
        } finally {
            Release-ComObject $openDocument
        }
    }
}

function Test-MathTypeInlineShape {
    param($Shape)
    try {
        return ([string]$Shape.OLEFormat.ProgID -eq $equationProgId)
    } catch {
        return $false
    }
}

function Get-ParagraphStyleName {
    param([Parameter(Mandatory = $true)]$Paragraph)
    try {
        return [string]$Paragraph.Style.NameLocal
    } catch {
        try { return [string]$Paragraph.Style } catch { return "" }
    }
}

function Invoke-MathTypeMacro {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)][string]$MacroName,
        [Parameter(Mandatory = $true)][string]$Purpose
    )
    try {
        # These MathType automation entry points are zero-argument commands.
        # Word.Run nevertheless exposes 30 optional Variant slots. Reflection
        # must explicitly marshal them as Missing: passing only MacroName can
        # raise DISP_E_PARAMNOTOPTIONAL, while passing $null sends a real macro
        # argument and can raise DISP_E_BADPARAMCOUNT.
        $bindingFlags = [System.Reflection.BindingFlags]::InvokeMethod -bor
            [System.Reflection.BindingFlags]::OptionalParamBinding
        [object[]]$arguments = [object[]]::new(31)
        $arguments[0] = $MacroName
        for ($argumentIndex = 1; $argumentIndex -lt $arguments.Length; $argumentIndex++) {
            $arguments[$argumentIndex] = [System.Type]::Missing
        }
        $null = $Word.GetType().InvokeMember(
            "Run",
            $bindingFlags,
            $null,
            $Word,
            $arguments
        )
    } catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException) { $message = $_.Exception.InnerException.Message }
        throw "MathType $Purpose macro failed: $MacroName. Confirm the MathType Commands for Word add-in is loaded and the command works manually. Error=$message"
    }
}

function Get-MathTypeNumberFields {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $result = [System.Collections.Generic.List[object]]::new()
    $contentStart = [int]$Paragraph.Range.Start
    $contentEnd = Get-ParagraphContentEnd -Paragraph $Paragraph
    for ($i = 1; $i -le $Paragraph.Range.Fields.Count; $i++) {
        $field = $Paragraph.Range.Fields.Item($i)
        try {
            $codeStart = [int]$field.Code.Start
            $codeEnd = [int]$field.Code.End
            $belongsToParagraphContent = (
                $codeStart -lt $contentEnd -and
                $codeEnd -gt $contentStart
            )
            if ($belongsToParagraphContent -and
                [string]$field.Code.Text -match '(?i)^\s*MACROBUTTON\s+MTPlaceRef(?:\s|$)') {
                $result.Add($field)
                $field = $null
            }
        } finally {
            Release-ComObject $field
        }
    }
    return @($result)
}

function Get-DocumentMathTypeNumberCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.Fields.Count; $i++) {
        $field = $Document.Fields.Item($i)
        try {
            if ([string]$field.Code.Text -match '(?i)^\s*MACROBUTTON\s+MTPlaceRef(?:\s|$)') { $count++ }
        } finally {
            Release-ComObject $field
        }
    }
    return $count
}

function Get-DocumentLegacySequenceEquationCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.Fields.Count; $i++) {
        $field = $Document.Fields.Item($i)
        try {
            if ([string]$field.Code.Text -match '(?i)^\s*SEQ\s+Equation(?:\s|$)') { $count++ }
        } finally {
            Release-ComObject $field
        }
    }
    return $count
}

function Get-MathTypeNumberingSetupCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.Fields.Count; $i++) {
        $field = $Document.Fields.Item($i)
        try {
            if ([string]$field.Code.Text -match '(?i)^\s*MACROBUTTON\s+MTEditEquationSection') { $count++ }
        } finally {
            Release-ComObject $field
        }
    }
    return $count
}

function Get-MathTypeObjectCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        try {
            if (Test-MathTypeInlineShape -Shape $shape) { $count++ }
        } finally {
            Release-ComObject $shape
        }
    }
    return $count
}

function Get-ParagraphContentEnd {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $range = $Paragraph.Range.Duplicate
    try {
        $end = [int]$range.End
        $text = [string]$range.Text
        for ($i = $text.Length - 1; $i -ge 0; $i--) {
            $code = [int][char]$text[$i]
            if ($code -eq 13 -or $code -eq 7) {
                $end--
            } else {
                break
            }
        }
        return $end
    } finally {
        Release-ComObject $range
    }
}

function Get-TargetDisplayParagraphs {
    param([Parameter(Mandatory = $true)]$Document)
    $targets = [System.Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $keepParagraph = $false
        try {
            if ((Get-ParagraphStyleName -Paragraph $paragraph) -cne $displayStyleName) { continue }
            if ($paragraph.Range.InlineShapes.Count -ne 1) { continue }
            $shape = $paragraph.Range.InlineShapes.Item(1)
            try {
                if (-not (Test-MathTypeInlineShape -Shape $shape)) { continue }
            } finally {
                Release-ComObject $shape
            }
            $targets.Add([pscustomobject]@{
                Index = $i
                Start = [int]$paragraph.Range.Start
            })
            $keepParagraph = $true
        } finally {
            Release-ComObject $paragraph
        }
    }
    return @($targets)
}

function Get-ParagraphAtStart {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][int]$Start
    )
    $range = $Document.Range($Start, $Start)
    try {
        return $range.Paragraphs.Item(1)
    } finally {
        Release-ComObject $range
    }
}

function Get-ParagraphResidualText {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$Paragraph,
        [Parameter(Mandatory = $true)]$Shape,
        [object[]]$IgnoredFields = @()
    )
    $contentEnd = Get-ParagraphContentEnd -Paragraph $Paragraph
    $excluded = [System.Collections.Generic.List[object]]::new()
    $excluded.Add([pscustomobject]@{ Start = [int]$Shape.Range.Start; End = [int]$Shape.Range.End })
    foreach ($field in $IgnoredFields) {
        $excluded.Add([pscustomobject]@{ Start = [int]$field.Result.Start; End = [int]$field.Result.End })
    }
    $ordered = @($excluded | Sort-Object Start, End)
    $cursor = [int]$Paragraph.Range.Start
    $builder = [System.Text.StringBuilder]::new()
    foreach ($item in $ordered) {
        $start = [Math]::Max([int]$item.Start, [int]$Paragraph.Range.Start)
        $end = [Math]::Min([int]$item.End, $contentEnd)
        if ($start -gt $cursor) {
            $gap = $Document.Range($cursor, $start)
            try { [void]$builder.Append([string]$gap.Text) } finally { Release-ComObject $gap }
        }
        if ($end -gt $cursor) { $cursor = $end }
    }
    if ($cursor -lt $contentEnd) {
        $tail = $Document.Range($cursor, $contentEnd)
        try { [void]$builder.Append([string]$tail.Text) } finally { Release-ComObject $tail }
    }
    return $builder.ToString()
}

function Assert-ParagraphContentSafe {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$Paragraph,
        [Parameter(Mandatory = $true)]$Shape,
        [object[]]$NumberFields = @()
    )

    # An inline OLE object can itself be exposed by Word as an EMBED field.
    # That field is part of the MathType object, not additional paragraph text.
    # Word does not report stable Result coordinates for these legacy OLE
    # fields, so match the exact ProgID instead. The caller has already proved
    # that this paragraph contains exactly one Equation.DSMT4 InlineShape.
    $mathTypeBackingFields = 0
    $unexpectedFieldCodes = [System.Collections.Generic.List[string]]::new()
    $contentStart = [int]$Paragraph.Range.Start
    $contentEnd = Get-ParagraphContentEnd -Paragraph $Paragraph
    for ($i = 1; $i -le $Paragraph.Range.Fields.Count; $i++) {
        $field = $Paragraph.Range.Fields.Item($i)
        try {
            $code = ([string]$field.Code.Text).Trim()
            $resultStart = [int]$field.Result.Start
            $resultEnd = [int]$field.Result.End
            $codeStart = [int]$field.Code.Start
            $codeEnd = [int]$field.Code.End
            $fieldStart = [Math]::Min($resultStart, $codeStart)
            $fieldEnd = [Math]::Max($resultEnd, $codeEnd)

            # Word can include a field from the following paragraph when a
            # Paragraph.Range ends on the shared paragraph boundary. It is not
            # content of this equation paragraph and must not trigger deletion
            # protection or idempotency checks.
            if ($fieldEnd -le $contentStart -or $fieldStart -ge $contentEnd) { continue }
            if ($code -match '(?i)^MACROBUTTON\s+MTPlaceRef(?:\s|$)') { continue }
            if ($NumberFields.Count -eq 1 -and $code -match '(?i)^SEQ\s+(?:MTEqn|MTChap|MTSec)(?:\s|$)') { continue }

            $isMathTypeBackingField = $code -match '(?i)^EMBED\s+Equation\.DSMT4(?:\s|$)'
            if ($isMathTypeBackingField) {
                $mathTypeBackingFields++
            } else {
                if ($code.Length -gt 120) { $code = $code.Substring(0, 117) + "..." }
                $unexpectedFieldCodes.Add($code)
            }
        } finally {
            Release-ComObject $field
        }
    }
    if ($mathTypeBackingFields -gt 1 -or $unexpectedFieldCodes.Count -gt 0) {
        $details = if ($unexpectedFieldCodes.Count -gt 0) {
            $unexpectedFieldCodes -join " | "
        } else {
            "Multiple EMBED Equation.DSMT4 fields: $mathTypeBackingFields"
        }
        throw "段落除 MathType 公式及其 OLE/EMBED 支撑域外还包含其他 Word 域，已停止以避免删除内容。ParagraphStart=$($Paragraph.Range.Start); Fields=$details"
    }
    $residual = Get-ParagraphResidualText -Document $Document -Paragraph $Paragraph -Shape $Shape -IgnoredFields $NumberFields
    if ($residual -notmatch '^[\s\u00A0]*$') {
        throw "行间公式段落除公式、MathType 原生编号和空白外还包含正文，已停止。Paragraph=$($Paragraph.Range.Start); ResidualText=$residual"
    }
}

function Get-DocxPartText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$PartName
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    $stream = $null
    $reader = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $archive.GetEntry($PartName)
        if ($null -eq $entry) { return $null }
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        return $reader.ReadToEnd()
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Test-DocxHasMathTypeNativeNumbers {
    param([Parameter(Mandatory = $true)][string]$Path)
    $documentXml = Get-DocxPartText -Path $Path -PartName "word/document.xml"
    return ($null -ne $documentXml -and $documentXml -match 'MACROBUTTON\s+MTPlaceRef')
}

function Set-DocxMathTypeNumberFormat {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Value
    )
    # Word's late-bound CustomDocumentProperties collection can be null even
    # though docProps/custom.xml exists. Update only the output DOCX package
    # before Word opens it, avoiding that unstable COM interface entirely.
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $customNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    $valueNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    $archive = $null
    $entryStream = $null
    $reader = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::Open(
            $Path,
            [System.IO.Compression.ZipArchiveMode]::Update
        )
        $entry = $archive.GetEntry("docProps/custom.xml")
        if ($null -eq $entry) {
            throw "DOCX 中缺少 docProps/custom.xml，无法安全设置 MathType 编号格式。请先在该文档副本中用 MathType 插入一次公式编号以初始化编号元数据，然后以未编号原件重试。"
        }

        $entryStream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($entryStream, [System.Text.Encoding]::UTF8, $true)
        $xmlText = $reader.ReadToEnd()
        $reader.Dispose()
        $reader = $null
        $entryStream = $null

        $xml = New-Object System.Xml.XmlDocument
        $xml.PreserveWhitespace = $true
        $xml.LoadXml($xmlText)
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $namespaceManager.AddNamespace("cp", $customNamespace)
        $property = $xml.SelectSingleNode("/cp:Properties/cp:property[@name='MTEquationNumber2']", $namespaceManager)
        if ($null -eq $property) {
            $property = $xml.CreateElement("property", $customNamespace)
            $property.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            $maxPid = 1
            foreach ($existing in $xml.SelectNodes("/cp:Properties/cp:property", $namespaceManager)) {
                $candidatePid = 0
                if ([int]::TryParse($existing.GetAttribute("pid"), [ref]$candidatePid) -and $candidatePid -gt $maxPid) {
                    $maxPid = $candidatePid
                }
            }
            $property.SetAttribute("pid", [string]($maxPid + 1))
            $property.SetAttribute("name", "MTEquationNumber2")
            [void]$xml.DocumentElement.AppendChild($property)
        } else {
            while ($property.HasChildNodes) {
                [void]$property.RemoveChild($property.FirstChild)
            }
        }
        $valueNode = $xml.CreateElement("vt", "lpwstr", $valueNamespace)
        $valueNode.InnerText = $Value
        [void]$property.AppendChild($valueNode)

        $entry.Delete()
        $replacement = $archive.CreateEntry("docProps/custom.xml", [System.IO.Compression.CompressionLevel]::Optimal)
        $replacementStream = $null
        $writer = $null
        try {
            $replacementStream = $replacement.Open()
            $settings = New-Object System.Xml.XmlWriterSettings
            $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
            $settings.Indent = $false
            $writer = [System.Xml.XmlWriter]::Create($replacementStream, $settings)
            $xml.Save($writer)
        } finally {
            if ($null -ne $writer) { $writer.Dispose() }
            elseif ($null -ne $replacementStream) { $replacementStream.Dispose() }
        }
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $entryStream) { $entryStream.Dispose() }
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Get-DocxMathTypeDisplayNumberTexts {
    param([Parameter(Mandatory = $true)][string]$Path)
    $documentXml = Get-DocxPartText -Path $Path -PartName "word/document.xml"
    if ([string]::IsNullOrWhiteSpace($documentXml)) {
        throw "DOCX is missing word/document.xml: $Path"
    }
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($documentXml)
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $namespaceManager.AddNamespace("o", "urn:schemas-microsoft-com:office:office")
    $paragraphs = $xml.SelectNodes(
        "//w:p[w:pPr/w:pStyle[@w:val='MTDisplayEquation'] and .//o:OLEObject[@ProgID='Equation.DSMT4']]",
        $namespaceManager
    )
    $results = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($paragraph in $paragraphs) {
        $index++
        $instructionTexts = [System.Collections.Generic.List[string]]::new()
        foreach ($instruction in $paragraph.SelectNodes(".//w:instrText", $namespaceManager)) {
            $instructionTexts.Add([string]$instruction.InnerText)
        }
        # Word may split a single instruction over several w:instrText nodes.
        $combinedInstructions = $instructionTexts -join " "
        $nativeNumberFields = [regex]::Matches(
            $combinedInstructions,
            '(?i)MACROBUTTON\s+MTPlaceRef(?:\s|$)'
        ).Count
        $visibleBuilder = New-Object System.Text.StringBuilder
        foreach ($textNode in $paragraph.SelectNodes(".//w:t", $namespaceManager)) {
            [void]$visibleBuilder.Append([string]$textNode.InnerText)
        }
        $results.Add([pscustomobject]@{
            Index = $index
            NativeNumberFields = $nativeNumberFields
            VisibleText = $visibleBuilder.ToString()
        })
    }
    return @($results)
}

function Set-EquationParagraphLayout {
    param(
        [Parameter(Mandatory = $true)]$Paragraph,
        [Parameter(Mandatory = $true)][double]$UsableWidth
    )
    $Paragraph.Format.Alignment = $wdAlignParagraphLeft
    $Paragraph.Format.LeftIndent = 0
    $Paragraph.Format.RightIndent = 0
    $Paragraph.Format.FirstLineIndent = 0
    $Paragraph.Format.TabStops.ClearAll()
    $centerTab = $Paragraph.Format.TabStops.Add($UsableWidth / 2.0, $wdAlignTabCenter, $wdTabLeaderSpaces)
    $rightTab = $Paragraph.Format.TabStops.Add($UsableWidth, $wdAlignTabRight, $wdTabLeaderSpaces)
    Release-ComObject $centerTab
    Release-ComObject $rightTab
    $Paragraph.Format.LineSpacingRule = $wdLineSpaceAtLeast
    $Paragraph.Format.LineSpacing = 30.0
    $Paragraph.Format.KeepTogether = $true
}

function Get-UsableWidthForParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $section = $Paragraph.Range.Sections.Item(1)
    $pageSetup = $null
    try {
        $pageSetup = $section.PageSetup
        $width = [double]$pageSetup.PageWidth - [double]$pageSetup.LeftMargin - [double]$pageSetup.RightMargin
        if ($width -le 0) {
            throw "Invalid PageSetup text width: PageWidth=$($pageSetup.PageWidth), LeftMargin=$($pageSetup.LeftMargin), RightMargin=$($pageSetup.RightMargin)"
        }
        return $width
    } finally {
        Release-ComObject $pageSetup
        Release-ComObject $section
    }
}

function Add-NumberToParagraph {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$Paragraph,
        [Parameter(Mandatory = $true)]$Word
    )
    $shape = $Paragraph.Range.InlineShapes.Item(1)
    try {
        $contentEnd = Get-ParagraphContentEnd -Paragraph $Paragraph
        if ([int]$shape.Range.End -lt $contentEnd) {
            $suffix = $Document.Range([int]$shape.Range.End, $contentEnd)
            try { $suffix.Text = "" } finally { Release-ComObject $suffix }
        }
        $shapeStart = [int]$shape.Range.Start
        if ([int]$Paragraph.Range.Start -lt $shapeStart) {
            $prefix = $Document.Range([int]$Paragraph.Range.Start, $shapeStart)
            try { $prefix.Text = "`t" } finally { Release-ComObject $prefix }
        } else {
            $prefix = $Document.Range($shapeStart, $shapeStart)
            try { $prefix.Text = "`t" } finally { Release-ComObject $prefix }
        }
    } finally {
        Release-ComObject $shape
    }

    $shape = $Paragraph.Range.InlineShapes.Item(1)
    try {
        $beforeNumbers = Get-DocumentMathTypeNumberCount -Document $Document
        $insert = $Document.Range([int]$shape.Range.End, [int]$shape.Range.End)
        try {
            $insert.Text = "`t"
            $insert.Collapse($wdCollapseEnd)
            $insert.Select()
            Invoke-WithRetry -Operation {
                Invoke-MathTypeMacro -Word $Word -MacroName $InsertNumberMacroName -Purpose "Insert Equation Number"
            } | Out-Null
        } finally {
            Release-ComObject $insert
        }
        $afterNumbers = Get-DocumentMathTypeNumberCount -Document $Document
        if ($afterNumbers -ne ($beforeNumbers + 1)) {
            throw "MathType Insert Equation Number did not create exactly one native MTPlaceRef number. Before=$beforeNumbers; After=$afterNumbers"
        }
    } finally {
        Release-ComObject $shape
    }
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
if ($StartNumber -ne 1) {
    throw "MathType native continuous numbering currently starts at 1. Use -StartNumber 1, then use MathType chapter/section breaks if a different scope is required."
}
$numberFormatValues = @{
    Simple = "(#E1)"
    ChapterDash = "(#C1-#E1)"
    ChapterDot = "(#C1.#E1)"
}
$requestedNumberFormat = [string]$numberFormatValues[$EquationNumberFormat]
if ([string]::IsNullOrWhiteSpace($InitializeNumbersMacroName)) {
    $InitializeNumbersMacroName = if ($EquationNumberFormat -eq "Simple") {
        "MTCommand_InsertNextSection"
    } else {
        "MTCommand_InsertNextChapter"
    }
}
if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    throw "Input DOCX not found: $InputPath"
}
if ([System.IO.Path]::GetExtension($InputPath) -ine ".docx") {
    throw "InputPath must be a .docx file: $InputPath"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $directory = [System.IO.Path]::GetDirectoryName($InputPath)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $OutputPath = Join-Path $directory ($stem + "_numbered.docx")
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if ([System.IO.Path]::GetExtension($OutputPath) -ine ".docx") {
    throw "OutputPath must be a .docx file: $OutputPath"
}
if ([System.StringComparer]::OrdinalIgnoreCase.Equals($InputPath, $OutputPath)) {
    throw "InputPath 与 OutputPath 不能相同。The input document is never modified in place."
}
if (Test-Path -LiteralPath $OutputPath) {
    if (-not $Force) {
        throw "Output already exists: $OutputPath. Close it and use -Force to replace it."
    }
    Test-FileClosed -Path $OutputPath -Label "旧输出 / Existing output" -Access ([System.IO.FileAccess]::ReadWrite)
}
Test-FileClosed -Path $InputPath -Label "输入 / Input" -Access ([System.IO.FileAccess]::Read)

$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
    [void][System.IO.Directory]::CreateDirectory($outputDirectory)
}

$word = $null
$document = $null
$alreadyNumbered = $false
$mathTypeNumberFormat = "Preserved existing MathType format"
$numberingStarted = $false
$completedInsertions = 0
$outputSavedByWord = $false
$lastSavedInsertionCount = 0
try {
    $word = Get-RunningMicrosoftWord
    Write-Host "Attached to Microsoft Word $($word.Version). 已连接用户打开的 Word。" -ForegroundColor Cyan
    Assert-TargetFilesNotOpenInWord -Word $word -Paths @($InputPath, $OutputPath)

    Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
    $packageHasNativeNumbers = Test-DocxHasMathTypeNativeNumbers -Path $OutputPath
    if ($packageHasNativeNumbers) {
        $existingPackageNumbers = @(Get-DocxMathTypeDisplayNumberTexts -Path $OutputPath)
        $incompletePackageNumbers = @($existingPackageNumbers | Where-Object {
            $_.NativeNumberFields -gt 0 -and [string]::IsNullOrWhiteSpace([string]$_.VisibleText)
        })
        if ($incompletePackageNumbers.Count -gt 0) {
            throw "输入文档包含没有可见结果的残缺 MathType MTPlaceRef 编号。Count=$($incompletePackageNumbers.Count); FirstParagraph=$($incompletePackageNumbers[0].Index). 请从未编号的 MathType 原件重新运行；不要把异常恢复输出作为新的 InputPath。"
        }
    } else {
        Set-DocxMathTypeNumberFormat -Path $OutputPath -Value $requestedNumberFormat
        $mathTypeNumberFormat = $requestedNumberFormat
    }
    $document = Invoke-WithRetry -Operation {
        $word.Documents.Open($OutputPath, $false, $false, $false)
    }
    if ([bool]$document.ReadOnly) {
        throw "输出副本被 Word 以只读方式打开。Close the output document everywhere and retry: $OutputPath"
    }

    $initialMathTypeObjects = Get-MathTypeObjectCount -Document $document
    $initialTables = [int]$document.Tables.Count
    $initialOMathObjects = [int]$document.OMaths.Count
    $legacySequenceFields = Get-DocumentLegacySequenceEquationCount -Document $document
    if ($legacySequenceFields -gt 0) {
        throw "检测到旧版脚本创建的 SEQ Equation 编号（Count=$legacySequenceFields）。它们不是 MathType 原生编号。请使用尚未添加旧编号的 MathType DOCX 作为 InputPath；脚本不会自动删除旧编号，以免破坏正文或引用。"
    }
    $preexistingNativeNumbers = Get-DocumentMathTypeNumberCount -Document $document
    if ($preexistingNativeNumbers -eq 0 -and (Get-MathTypeNumberingSetupCount -Document $document) -eq 0) {
        Write-Host "Initializing MathType equation numbering metadata..." -ForegroundColor Cyan
        $documentStart = $document.Range(0, 0)
        try { $documentStart.Select() } finally { Release-ComObject $documentStart }
        Invoke-WithRetry -Operation {
            Invoke-MathTypeMacro -Word $word -MacroName $InitializeNumbersMacroName -Purpose "Initialize Equation Numbering"
        } | Out-Null
    }
    $targets = @(Get-TargetDisplayParagraphs -Document $document)
    $initialDisplayParagraphs = $targets.Count

    if ($ExpectedDisplay -ge 0 -and $initialDisplayParagraphs -ne $ExpectedDisplay) {
        throw "行间公式数量与 ExpectedDisplay 不一致。Detected=$initialDisplayParagraphs; Expected=$ExpectedDisplay"
    }
    if ($initialDisplayParagraphs -eq 0) {
        throw "未找到符合条件的行间公式段落。Expected style=$displayStyleName, exactly one inline OLE object, ProgID=$equationProgId."
    }

    $numberedTargets = 0
    foreach ($target in $targets) {
        $paragraph = Get-ParagraphAtStart -Document $document -Start $target.Start
        $shape = $null
        $fields = @()
        try {
            $shape = $paragraph.Range.InlineShapes.Item(1)
            $fields = @(Get-MathTypeNumberFields -Paragraph $paragraph)
            if ($fields.Count -gt 1) {
                throw "行间公式段落包含多个 MathType MTPlaceRef 编号域。Paragraph=$($target.Index); Count=$($fields.Count)"
            }
            Assert-ParagraphContentSafe -Document $document -Paragraph $paragraph -Shape $shape -NumberFields $fields
            if ($fields.Count -eq 1) { $numberedTargets++ }
        } finally {
            foreach ($field in $fields) { Release-ComObject $field }
            Release-ComObject $shape
            Release-ComObject $paragraph
        }
    }

    $initialNumberFields = Get-DocumentMathTypeNumberCount -Document $document
    if ($numberedTargets -gt 0 -and $numberedTargets -lt $initialDisplayParagraphs) {
        throw "检测到部分编号状态，已停止以避免形成混合编号。Numbered display paragraphs=$numberedTargets; Display paragraphs=$initialDisplayParagraphs."
    }
    if ($numberedTargets -eq 0 -and $initialNumberFields -ne 0) {
        throw "文档中存在不属于目标行间公式段落的 MathType MTPlaceRef 编号域。Fields=$initialNumberFields. 请先清理或核查这些域。"
    }
    if ($numberedTargets -eq $initialDisplayParagraphs -and $initialNumberFields -ne $initialDisplayParagraphs) {
        throw "目标公式均有编号，但文档还有额外的 MathType MTPlaceRef 编号域。Target=$initialDisplayParagraphs; Fields=$initialNumberFields."
    }

    if ($numberedTargets -eq $initialDisplayParagraphs) {
        $alreadyNumbered = $true
        Write-Host "All display equations already have MathType-native numbers; updating through MathType. 已存在完整 MathType 编号，仅更新。" -ForegroundColor Yellow
    } else {
        $numberingStarted = $true
        Write-Host "Adding MathType-native right-side numbers to $initialDisplayParagraphs display equations..." -ForegroundColor Cyan
        # Work from the end so edits cannot invalidate the stored starts of
        # paragraphs that have not been processed yet.
        for ($i = $targets.Count - 1; $i -ge 0; $i--) {
            $paragraph = Get-ParagraphAtStart -Document $document -Start $targets[$i].Start
            try {
                $usableWidth = Get-UsableWidthForParagraph -Paragraph $paragraph
                Set-EquationParagraphLayout -Paragraph $paragraph -UsableWidth $usableWidth
                Add-NumberToParagraph -Document $document -Paragraph $paragraph -Word $word
                $completedInsertions = $targets.Count - $i
                if ($completedInsertions % 10 -eq 0 -or $completedInsertions -eq $targets.Count) {
                    $document.Save()
                    $outputSavedByWord = $true
                    $lastSavedInsertionCount = $completedInsertions
                    Write-Host ("Checkpoint saved: {0}/{1}. 已保存恢复检查点。" -f $completedInsertions, $targets.Count) -ForegroundColor DarkCyan
                }
                if ($completedInsertions -eq 1 -or $completedInsertions % 10 -eq 0 -or $completedInsertions -eq $targets.Count) {
                    Write-Host ("Inserted MathType numbers: {0}/{1}" -f $completedInsertions, $targets.Count)
                }
            } finally {
                Release-ComObject $paragraph
            }
        }
    }

    $collapsedSelection = $document.Range(0, 0)
    try { $collapsedSelection.Select() } finally { Release-ComObject $collapsedSelection }
    Invoke-WithRetry -Operation {
        Invoke-MathTypeMacro -Word $word -MacroName $UpdateNumbersMacroName -Purpose "Update Equation Numbers"
    } | Out-Null
    $document.Fields.Update() | Out-Null

    $finalMathTypeObjects = Get-MathTypeObjectCount -Document $document
    $finalTables = [int]$document.Tables.Count
    $finalOMathObjects = [int]$document.OMaths.Count
    $finalTargets = @(Get-TargetDisplayParagraphs -Document $document)
    $finalNumberFields = Get-DocumentMathTypeNumberCount -Document $document

    if ($finalMathTypeObjects -ne $initialMathTypeObjects) {
        throw "MathType object count changed: $initialMathTypeObjects -> $finalMathTypeObjects"
    }
    if ($finalTables -ne $initialTables) {
        throw "Table count changed: $initialTables -> $finalTables"
    }
    if ($finalTargets.Count -ne $initialDisplayParagraphs) {
        throw "Display-equation paragraph count changed: $initialDisplayParagraphs -> $($finalTargets.Count)"
    }
    if ($finalNumberFields -ne $finalTargets.Count) {
        throw "MathType MTPlaceRef number count mismatch: Fields=$finalNumberFields; Display=$($finalTargets.Count)"
    }
    if ($finalOMathObjects -ne 0 -or $initialOMathObjects -ne 0) {
        throw "Word OMML/OMath objects are not allowed. Before=$initialOMathObjects; After=$finalOMathObjects"
    }

    for ($targetIndex = 0; $targetIndex -lt $finalTargets.Count; $targetIndex++) {
        $target = $finalTargets[$targetIndex]
        $paragraph = Get-ParagraphAtStart -Document $document -Start $target.Start
        $fields = @()
        try {
            $fields = @(Get-MathTypeNumberFields -Paragraph $paragraph)
            if ($fields.Count -ne 1) {
                throw "Target paragraph does not contain exactly one MathType MTPlaceRef number. Paragraph=$($target.Index); Fields=$($fields.Count)"
            }
        } finally {
            foreach ($field in $fields) { Release-ComObject $field }
            Release-ComObject $paragraph
        }
    }

    $document.Save()
    $outputSavedByWord = $true
    $lastSavedInsertionCount = $completedInsertions
    $document.Close($wdDoNotSaveChanges)
    Release-ComObject $document
    $document = $null

    $packageNumberResults = @(Get-DocxMathTypeDisplayNumberTexts -Path $OutputPath)
    if ($packageNumberResults.Count -ne $finalTargets.Count) {
        throw "Saved DOCX display-equation count mismatch: Package=$($packageNumberResults.Count); Word=$($finalTargets.Count)"
    }
    for ($targetIndex = 0; $targetIndex -lt $packageNumberResults.Count; $targetIndex++) {
        $packageNumber = $packageNumberResults[$targetIndex]
        if ($packageNumber.NativeNumberFields -ne 1) {
            throw "Saved DOCX target does not contain exactly one MathType MTPlaceRef number. Paragraph=$($packageNumber.Index); Fields=$($packageNumber.NativeNumberFields)"
        }
        if (-not $alreadyNumbered) {
            $numberPattern = switch ($EquationNumberFormat) {
                "Simple" { '^\s*\((?<equation>\d+)\)\s*$' }
                "ChapterDash" { '^\s*\((?<chapter>\d+)-(?<equation>\d+)\)\s*$' }
                "ChapterDot" { '^\s*\((?<chapter>\d+)\.(?<equation>\d+)\)\s*$' }
            }
            $match = [regex]::Match([string]$packageNumber.VisibleText, $numberPattern)
            if (-not $match.Success) {
                throw "Saved MathType native number does not match format $EquationNumberFormat. Paragraph=$($packageNumber.Index); VisibleText=$($packageNumber.VisibleText)"
            }
            $expectedNumber = $StartNumber + $targetIndex
            if ([int]$match.Groups['equation'].Value -ne $expectedNumber) {
                throw "Saved MathType native number sequence mismatch. Paragraph=$($packageNumber.Index); Actual=$($match.Groups['equation'].Value); Expected=$expectedNumber"
            }
            if ($EquationNumberFormat -ne "Simple" -and [int]$match.Groups['chapter'].Value -ne 1) {
                throw "Saved MathType chapter number mismatch. Paragraph=$($packageNumber.Index); Actual=$($match.Groups['chapter'].Value); Expected=1"
            }
        }
    }

    Write-Host "Equation numbering succeeded. 公式编号完成。" -ForegroundColor Green
    $summary = [pscustomobject]@{
        Status = "Success"
        InputPath = $InputPath
        OutputPath = $OutputPath
        AlreadyNumbered = $alreadyNumbered
        StartNumber = $StartNumber
        MathTypeNumberFormat = $mathTypeNumberFormat
        EquationNumberFormat = $EquationNumberFormat
        DisplayEquationParagraphs = $finalTargets.Count
        MathTypeNumberFields = $finalNumberFields
        MathTypeObjectsBefore = $initialMathTypeObjects
        MathTypeObjectsAfter = $finalMathTypeObjects
        TablesBefore = $initialTables
        TablesAfter = $finalTables
        WordOMathObjectsBefore = $initialOMathObjects
        WordOMathObjectsAfter = $finalOMathObjects
    }
    Write-Host ("RESULT_JSON " + ($summary | ConvertTo-Json -Compress))
    $summary
} catch {
    $failure = $_
    $recoverySaveError = ""
    if ($numberingStarted -and $null -ne $document) {
        try {
            $document.Save()
            $outputSavedByWord = $true
            $lastSavedInsertionCount = $completedInsertions
            Write-Host ("Recovery output saved after failure: {0}" -f $OutputPath) -ForegroundColor Yellow
        } catch {
            $recoverySaveError = $_.Exception.Message
            Write-Host ("Could not save the latest recovery state: {0}" -f $recoverySaveError) -ForegroundColor Red
        }
    }
    $failureActions = [System.Collections.Generic.List[string]]::new()
    $failureActions.Add("关闭输入文档和旧输出文档，只保留一个空白 Microsoft Word 窗口。")
    $failureActions.Add("Close the input and old output DOCX; keep one blank Microsoft Word window open.")
    $failureActions.Add("确保 Word 与 PowerShell 权限级别相同，并关闭 WPS/Kingsoft。")
    if ($outputSavedByWord) {
        $failureActions.Add("已保存恢复输出：$OutputPath；已确认插入 $completedInsertions 个编号，最近检查点为 $lastSavedInsertionCount。请勿把部分编号文档作为新的 InputPath 自动续跑。")
        $failureActions.Add("Recovery output was saved. Inspect it manually or rerun from the original unnumbered input with -Force.")
        Write-Host ("RECOVERY_JSON " + ([pscustomobject]@{
            Status = "PartialRecovery"
            OutputPath = $OutputPath
            CompletedInsertions = $completedInsertions
            LastSavedInsertionCount = $lastSavedInsertionCount
        } | ConvertTo-Json -Compress))
    } else {
        $failureActions.Add("输入文件从未被修改；本次没有可确认的 Word 恢复检查点。")
        if (-not [string]::IsNullOrWhiteSpace($recoverySaveError)) {
            $failureActions.Add("Recovery save error: $recoverySaveError")
        }
    }
    Write-UserActionError `
        -Title "DOCX 行间公式批量编号失败 / Batch equation numbering failed." `
        -Reason $failure.Exception.Message `
        -Actions @($failureActions)
    throw $failure
} finally {
    if ($null -ne $document) {
        try { $document.Close($wdDoNotSaveChanges) } catch {}
        Release-ComObject $document
    }
    if ($null -ne $word) {
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
