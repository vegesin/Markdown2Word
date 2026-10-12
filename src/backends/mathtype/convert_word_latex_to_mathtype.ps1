[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,
    [string]$OutputPath,
    [ValidateRange(-1, 100000)]
    [int]$ExpectedInline = -1,
    [ValidateRange(-1, 100000)]
    [int]$ExpectedDisplay = -1,
    [string]$MacroName = "MTCommand_OnTexToggle",
    [ValidateRange(0, 3000)]
    [int]$DelayMilliseconds = 100,
    [switch]$StartWordIfNeeded,
    [switch]$AddEquationNumbers,
    [ValidateSet("Simple", "ChapterDash", "ChapterDot")]
    [string]$EquationNumberFormat = "Simple",
    [string]$InsertNumberMacroName = "MTCommand_InsertEqnNum",
    [string]$UpdateNumbersMacroName = "MTCommand_UpdateEqns",
    [string]$InitializeNumbersMacroName = "",
    [ValidateRange(12, 200)]
    [double]$DisplayMinPt = 30,
    [ValidateRange(0, 200)]
    [double]$DisplaySpaceBeforePt = 3,
    [ValidateRange(0, 200)]
    [double]$DisplaySpaceAfterPt = 3,
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
$wdCollapseStart = 1
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdAlignTabCenter = 1
$wdAlignTabRight = 2
$wdTabLeaderSpaces = 0
$wdLineSpaceAtLeast = 3
$wdWithInTable = 12
$rpcRejected = @(-2147418111, -2147417846)
$script:NumericFallbackConversions = 0
$script:PrimeFallbackConversions = 0
$script:CommaListFallbackConversions = 0
$script:TableInlineDialogWatchers = 0
$script:TableInlineDialogsHandled = 0
$script:TransientFormulaRetries = 0
$script:DisplayParagraphsReformatted = 0
$script:DisplayEmptyParagraphsRemoved = 0
$script:MathTypeNativeNumbersInserted = 0

function Write-UserActionError {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Reason,
        [string[]]$Actions = @()
    )
    Write-Host ""
    Write-Host "================ MathType conversion stopped ================" -ForegroundColor Red
    Write-Host $Title -ForegroundColor Red
    Write-Host "Reason : $Reason" -ForegroundColor Yellow
    if ($Actions.Count -gt 0) {
        Write-Host "Action :" -ForegroundColor Cyan
        foreach ($action in $Actions) {
            Write-Host "  - $action" -ForegroundColor Cyan
        }
    }
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host ""
}


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
        [int]$Delay = 350
    )
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return & $Operation
        } catch [System.Runtime.InteropServices.COMException] {
            if ($rpcRejected -notcontains $_.Exception.HResult -or $attempt -eq $Retries) {
                throw
            }
            Start-Sleep -Milliseconds $Delay
        }
    }
}

function Get-RunningWord {
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    } catch {
        if ($StartWordIfNeeded) {
            try {
                $word = Connect-DesktopWord -StartIfNeeded
                $word.Visible = $true
                # Keep one harmless document open so the Word instance remains
                # available for the optional numbering stage after this script.
                $blankDocument = $word.Documents.Add()
                Release-ComObject $blankDocument
                Write-Host "Microsoft Word was started automatically." -ForegroundColor Cyan
                return $word
            } catch {
                Write-UserActionError `
                    -Title "Microsoft Word could not be started automatically." `
                    -Reason $_.Exception.Message `
                    -Actions @(
                        "Confirm desktop Microsoft Word is installed and activated.",
                        "Start Word manually at the same privilege level, then retry."
                    )
                throw
            }
        }
        Write-UserActionError `
            -Title "Microsoft Word is not open or is not visible to this PowerShell session." `
            -Reason "No running Word instance was found and automatic startup is disabled." `
            -Actions @(
                "Open Microsoft Word manually.",
                "Create or keep one blank Word document open.",
                "Run PowerShell and Word at the same privilege level; do not run one as Administrator and the other normally.",
                "Verify MathType Toggle TeX works manually, then rerun this script."
            )
        throw "No running Word instance was found. Open Word manually, verify MathType Toggle TeX works, leave Word open, and rerun."
    }
    if ($word.Name -notlike "*Microsoft Word*" -or $word.Path -match "WPS|Kingsoft") {
        Write-UserActionError `
            -Title "The active Word.Application is not Microsoft Word." `
            -Reason "Detected $($word.Name) at $($word.Path)." `
            -Actions @("Close WPS/Kingsoft Word-compatible editors.", "Open Microsoft Word and rerun.")
        throw "The running Word.Application is not Microsoft Word: $($word.Name) at $($word.Path)"
    }
    $word.Visible = $true
    return $word
}

function Invoke-MathTypeToggle {
    param([Parameter(Mandatory = $true)]$Word)
    try {
        $null = $Word.GetType().InvokeMember(
            "Run",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $Word,
            @($MacroName, $null)
        )
    } catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $message = $_.Exception.InnerException.Message
        }
        Write-UserActionError `
            -Title "MathType Toggle TeX could not run." `
            -Reason $message `
            -Actions @(
                "Confirm MathType 6 is licensed and activated in this Windows account.",
                "In Word, manually select a test formula such as `$x^2`$ and run MathType Toggle TeX.",
                "If Word shows a MathType license/demo dialog, close it and fix activation before rerunning.",
                "Keep Word open and do not interact with it while conversion is running."
            )
        throw "MathType Toggle TeX failed: $message"
    }
}

function Test-MathTypeInlineShape {
    param($Shape)
    try {
        return $Shape.OLEFormat.ProgID -eq "Equation.DSMT4"
    } catch {
        return $false
    }
}

function Get-MathTypeInlineShapes {
    param([Parameter(Mandatory = $true)]$Document)
    $result = [System.Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        if (Test-MathTypeInlineShape -Shape $shape) {
            $result.Add($shape)
        }
    }
    return @($result)
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

function Wait-MathTypeObjectCount {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][int]$ExpectedCount,
        [ValidateRange(100, 10000)][int]$TimeoutMilliseconds = 3000
    )
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    $count = Get-MathTypeObjectCount -Document $Document
    while ($count -ne $ExpectedCount -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 100
        $count = Get-MathTypeObjectCount -Document $Document
    }
    return $count
}

function Clear-TemporaryDocument {
    param([Parameter(Mandatory = $true)]$Document)
    $range = $Document.Content
    $range.Delete() | Out-Null
    Release-ComObject $range
}

function Put-FormulaInTemporaryDocument {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$TempDocument,
        [Parameter(Mandatory = $true)][string]$Literal
    )
    Clear-TemporaryDocument -Document $TempDocument
    $TempDocument.Activate()
    $range = $TempDocument.Range(0, 0)
    $range.Text = $Literal
    $range.SetRange(0, $Literal.Length)
    $range.Select()
    Release-ComObject $range

    Invoke-WithRetry -Operation { Invoke-MathTypeToggle -Word $Word } | Out-Null
    if ($DelayMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $DelayMilliseconds
    }

    $shapes = @(Get-MathTypeInlineShapes -Document $TempDocument)
    if ($shapes.Count -ne 1) {
        throw "MathType produced $($shapes.Count) Equation.DSMT4 objects for formula: $Literal"
    }
    return $shapes[0]
}

function ConvertTo-AsciiPrimeText {
    param([Parameter(Mandatory = $true)][string]$Text)
    # Avoid non-ASCII literals in regex source. Windows PowerShell 5 may read a
    # UTF-8 script through the active ANSI code page when the file has no BOM.
    return $Text.Replace([char]0x2018, [char]0x27).Replace([char]0x2019, [char]0x27).Replace([char]0x2032, [char]0x27)
}

function Test-FormulaRangeInTable {
    param([Parameter(Mandatory = $true)]$FormulaRange)
    try {
        if ([bool]$FormulaRange.Information($wdWithInTable)) { return $true }
    } catch {}
    # Word can report wdWithInTable incorrectly for a range whose endpoints
    # sit on run/cell boundaries. Use the range collections as fallbacks.
    try { if ([int]$FormulaRange.Cells.Count -gt 0) { return $true } } catch {}
    try { if ([int]$FormulaRange.Tables.Count -gt 0) { return $true } } catch {}
    return $false
}

function Get-TaskWordWindowHandle {
    param($Word)
    $activeWindow = $null
    try {
        # Hwnd belongs to Word.Window, not Word.Application. The caller has
        # activated the target document before starting this scoped watcher.
        $activeWindow = $Word.ActiveWindow
        $handle = [int64]$activeWindow.Hwnd
        if ($handle -eq 0) { throw 'The active Word window has no valid handle.' }
        return $handle
    } finally {
        if ($null -ne $activeWindow -and [Runtime.InteropServices.Marshal]::IsComObject($activeWindow)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($activeWindow)
        }
    }
}

function Start-InlineStyleDialogWatcher {
    param([Parameter(Mandatory = $true)]$Word)
    $watcherCode = @'
$ErrorActionPreference = "SilentlyContinue"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

public static class MathTypeDialogNative {
    private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowProc callback, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowProc callback, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maxCount);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
    private const uint BM_CLICK = 0x00F5;

    public static bool BelongsToTask(IntPtr window) {
        long handle;
        if (!long.TryParse(Environment.GetEnvironmentVariable("M2W_WORD_HWND"), out handle)) return false;
        uint wordPid;
        GetWindowThreadProcessId(new IntPtr(handle), out wordPid);
        for (int i = 0; i < 20 && window != IntPtr.Zero; i++) {
            uint pid;
            GetWindowThreadProcessId(window, out pid);
            if (wordPid != 0 && pid == wordPid) return true;
            window = GetWindow(window, 4); // GW_OWNER
        }
        return false;
    }

    private static string WindowText(IntPtr hwnd) {
        int length = Math.Max(GetWindowTextLength(hwnd) + 1, 512);
        StringBuilder text = new StringBuilder(length);
        GetWindowText(hwnd, text, text.Capacity);
        return text.ToString();
    }

    private static string NormalizedCaption(IntPtr hwnd) {
        return Regex.Replace(WindowText(hwnd).Replace("&", ""), "\\s+", " ").Trim();
    }

    public static bool ClickInlineStyleButton() {
        bool clicked = false;
        EnumWindows(delegate(IntPtr window, IntPtr unused) {
            if (!BelongsToTask(window)) return true;
            string title = NormalizedCaption(window);
            if (!Regex.IsMatch(title, "^Convert\\s+To\\s+Display\\s+Style\\??$", RegexOptions.IgnoreCase)) return true;
            EnumChildWindows(window, delegate(IntPtr child, IntPtr childUnused) {
                string caption = NormalizedCaption(child);
                // MathType 6 may use a custom/owner-drawn control class.  Match
                // the exact dialog title and exact control caption instead of
                // assuming that its Win32 class contains the word "Button".
                if (Regex.IsMatch(caption, "^Create\\s+Inline\\s+Style\\s+Equation$", RegexOptions.IgnoreCase)) {
                    clicked = PostMessage(child, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return !clicked;
        }, IntPtr.Zero);
        return clicked;
    }
}
"@
$deadline = [DateTime]::UtcNow.AddSeconds(20)
$root = [System.Windows.Automation.AutomationElement]::RootElement
$buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::Button
)
while ([DateTime]::UtcNow -lt $deadline) {
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($window in $windows) {
        if (-not [MathTypeDialogNative]::BelongsToTask([IntPtr]$window.Current.NativeWindowHandle)) { continue }
        if ($window.Current.Name -notmatch '^Convert\s+To\s+Display\s+Style\??$') { continue }
        $buttons = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
        foreach ($button in $buttons) {
            $name = $button.Current.Name
            $normalizedName = (($name -replace '&', '') -replace '\s+', ' ').Trim()
            if ($normalizedName -match '^Create\s+Inline\s+Style\s+Equation$|^Inline\s+Style\s+Equation$|^\u521b\u5efa.*\u884c\u5185.*\u516c\u5f0f$|^\u884c\u5185.*\u516c\u5f0f$') {
                try {
                    $pattern = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                    ([System.Windows.Automation.InvokePattern]$pattern).Invoke()
                    exit 0
                } catch {
                    try {
                        $legacy = $button.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern)
                        ([System.Windows.Automation.LegacyIAccessiblePattern]$legacy).DoDefaultAction()
                        exit 0
                    } catch {}
                }
            }
        }
    }
    # MathType 6 uses a legacy modal dialog on some systems. It can expose the
    # button caption without a usable UI Automation InvokePattern, so use a
    # narrowly targeted Win32 button click as a fallback.
    if ([MathTypeDialogNative]::ClickInlineStyleButton()) { exit 0 }
    Start-Sleep -Milliseconds 100
}
exit 2
'@
    $wordHandle = Get-TaskWordWindowHandle -Word $Word
    $watcherCode = "`$env:M2W_WORD_HWND = '$wordHandle'`n" + $watcherCode
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($watcherCode))
    $powerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $process = Start-Process -FilePath $powerShellExe -ArgumentList @(
        "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", $encoded
    ) -PassThru -WindowStyle Hidden
    Start-Sleep -Milliseconds 150
    return $process
}

function Stop-InlineStyleDialogWatcher {
    param($Process)
    if ($null -eq $Process) { return $false }
    $handled = $false
    try {
        if (-not $Process.HasExited) { [void]$Process.WaitForExit(300) }
        if ($Process.HasExited -and $Process.ExitCode -eq 0) { $handled = $true }
        if (-not $Process.HasExited) {
            $Process.Kill()
            [void]$Process.WaitForExit(1000)
        }
    } catch {
    } finally {
        $Process.Dispose()
    }
    return $handled
}

function Invoke-MathTypeToggleForFormula {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [bool]$WatchInlineStyleDialog,
        [bool]$FormulaIsInTable
    )
    $watcher = $null
    try {
        if ($WatchInlineStyleDialog) {
            $watcher = Start-InlineStyleDialogWatcher -Word $Word
            $script:TableInlineDialogWatchers++
            Write-Host ("Inline-style dialog guard started. InTable={0}" -f $FormulaIsInTable) -ForegroundColor DarkCyan
        }
        Invoke-WithRetry -Operation { Invoke-MathTypeToggle -Word $Word } | Out-Null
    } finally {
        $handled = Stop-InlineStyleDialogWatcher -Process $watcher
        if ($handled) {
            $script:TableInlineDialogsHandled++
            Write-Host "MathType table dialog handled automatically: Create Inline Style Equation." -ForegroundColor DarkCyan
        }
    }
}

function Get-FormulaRetryLiterals {
    param([Parameter(Mandatory = $true)]$Formula)

    if ([string]$Formula.Kind -ne "inline") { return @() }
    $literal = [string]$Formula.Literal
    if ($literal.Length -lt 2 -or -not ($literal.StartsWith('$') -and $literal.EndsWith('$'))) {
        return @()
    }
    $content = $literal.Substring(1, $literal.Length - 2).Trim()
    $number = '[+-]?(?:\d+(?:\.\d*)?|\.\d+)'
    $interval = [regex]::Match($content, "^\s*([\[\(])\s*($number)\s*,\s*($number)\s*([\]\)])\s*$")
    if ($interval.Success) {
        $left = $interval.Groups[1].Value
        $first = $interval.Groups[2].Value
        $second = $interval.Groups[3].Value
        $right = $interval.Groups[4].Value
        return @(
            ('$' + '\left' + $left + $first + ',\,' + $second + '\right' + $right + '$'),
            ('$' + '\mathrm{' + $left + $first + ',\,' + $second + $right + '}' + '$')
        )
    }
    if ([regex]::IsMatch($content, "^\s*$number\s*$")) {
        return @(('$' + '\mathrm{' + $content + '}' + '$'))
    }
    $primeSafeContent = ConvertTo-AsciiPrimeText -Text $content
    $primeMatch = [regex]::Match($primeSafeContent, "^(?<base>.+?)(?<primes>'+)$")
    if ($primeMatch.Success) {
        $base = $primeMatch.Groups['base'].Value.TrimEnd()
        $primeCount = $primeMatch.Groups['primes'].Value.Length
        if ($base.Length -gt 0 -and $primeCount -gt 0) {
            $primeCommands = ((1..$primeCount) | ForEach-Object { '\prime' }) -join ''
            return @(('$' + $base + '^{' + $primeCommands + '}' + '$'))
        }
    }
    $commaListPattern = '^[A-Za-z][A-Za-z0-9_{}]*(?:\s*,\s*[A-Za-z][A-Za-z0-9_{}]*)+$'
    if ([regex]::IsMatch($content, $commaListPattern)) {
        $commaSafe = [regex]::Replace($content, '\s*,\s*', ',\,')
        return @(('$' + $commaSafe + '$'))
    }
    return @()
}

function Get-ParagraphStyleName {
    param([Parameter(Mandatory = $true)]$Paragraph)
    try { return [string]$Paragraph.Style.NameLocal } catch { try { return [string]$Paragraph.Style } catch { return "" } }
}

function Test-DisplayEquationParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $styleName = Get-ParagraphStyleName -Paragraph $Paragraph
    if ($styleName -like "MTDisplay*") { return $true }
    return $false
}

function Test-EmptyParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    try {
        if ($Paragraph.Range.InlineShapes.Count -ne 0) { return $false }
        if ($Paragraph.Range.Fields.Count -ne 0) { return $false }
        $text = [string]$Paragraph.Range.Text
        return [regex]::IsMatch($text, '^[\s\r\n\a]*$')
    } catch {
        return $false
    }
}

function Get-MathTypeShapeInParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $found = @()
    for ($i = 1; $i -le $Paragraph.Range.InlineShapes.Count; $i++) {
        $shape = $Paragraph.Range.InlineShapes.Item($i)
        if (Test-MathTypeInlineShape -Shape $shape) {
            $found += $shape
        } else {
            Release-ComObject $shape
        }
    }
    if ($found.Count -eq 1) { return $found[0] }
    foreach ($shape in $found) { Release-ComObject $shape }
    return $null
}

function Repair-DisplayEquationLayout {
    param([Parameter(Mandatory = $true)]$Document)
    $reformatted = 0
    $removed = 0
    for ($i = $Document.Paragraphs.Count; $i -ge 1; $i--) {
        $paragraph = $Document.Paragraphs.Item($i)
        $shape = $null
        try {
            if (-not (Test-DisplayEquationParagraph -Paragraph $paragraph)) { continue }
            $shape = Get-MathTypeShapeInParagraph -Paragraph $paragraph
            if ($null -ne $shape) {
                $prefix = $Document.Range($paragraph.Range.Start, $shape.Range.Start)
                try {
                    if (-not [string]::IsNullOrWhiteSpace([string]$prefix.Text)) {
                        $prefix.Text = ""
                    }
                } finally {
                    Release-ComObject $prefix
                }
                $suffixEnd = [int]$paragraph.Range.End - 1
                if ($shape.Range.End -le $suffixEnd) {
                    $suffix = $Document.Range($shape.Range.End, $suffixEnd)
                    try {
                        if (-not [string]::IsNullOrWhiteSpace([string]$suffix.Text)) {
                            $suffix.Text = ""
                        }
                    } finally {
                        Release-ComObject $suffix
                    }
                }
            }

            $paragraph.Format.Alignment = $wdAlignParagraphCenter
            $paragraph.Format.LeftIndent = 0
            $paragraph.Format.RightIndent = 0
            $paragraph.Format.FirstLineIndent = 0
            $paragraph.Format.TabStops.ClearAll()
            $paragraph.Format.LineSpacingRule = $wdLineSpaceAtLeast
            $paragraph.Format.LineSpacing = $DisplayMinPt
            $paragraph.Format.SpaceBefore = $DisplaySpaceBeforePt
            $paragraph.Format.SpaceAfter = $DisplaySpaceAfterPt
            $paragraph.Format.KeepTogether = $true
            $reformatted++

            if ($i -lt $Document.Paragraphs.Count) {
                $next = $Document.Paragraphs.Item($i + 1)
                try {
                    if (Test-EmptyParagraph -Paragraph $next) {
                        $next.Range.Delete() | Out-Null
                        $removed++
                    }
                } finally {
                    Release-ComObject $next
                }
            }
        } finally {
            Release-ComObject $shape
            Release-ComObject $paragraph
        }
    }
    $script:DisplayParagraphsReformatted += $reformatted
    $script:DisplayEmptyParagraphsRemoved += $removed
}

function Test-InlineConversionExtraPeriod {
    param(
        [AllowEmptyString()][string]$ExpectedPrefix,
        [AllowEmptyString()][string]$ActualPrefix,
        [AllowEmptyString()][string]$ExpectedSuffix,
        [AllowEmptyString()][string]$ActualSuffix
    )
    # Do not normalize punctuation globally. Only an exact one-character
    # insertion at the converted object's boundary may be repaired.
    return ($ExpectedPrefix -ceq $ActualPrefix -and $ActualSuffix -ceq ('.' + $ExpectedSuffix))
}

function Repair-InlineConversionExtraPeriod {
    param($Document, [int]$FormulaStart, [string]$ExpectedPrefix, [string]$ExpectedSuffix)
    $probe = $null; $shape = $null; $paragraph = $null
    $prefix = $null; $suffix = $null; $extra = $null
    try {
        $probe = $Document.Range($FormulaStart, $FormulaStart + 1)
        if ($probe.InlineShapes.Count -ne 1) { return }
        $shape = $probe.InlineShapes.Item(1)
        if (-not (Test-MathTypeInlineShape -Shape $shape)) { return }
        if ([int]$shape.Range.Start -ne $FormulaStart) { return }
        $paragraph = $shape.Range.Paragraphs.Item(1)
        $prefix = $Document.Range([int]$paragraph.Range.Start, $FormulaStart)
        $suffix = $Document.Range([int]$shape.Range.End, [int]$paragraph.Range.End)
        if (Test-InlineConversionExtraPeriod -ExpectedPrefix $ExpectedPrefix -ActualPrefix ([string]$prefix.Text) `
                -ExpectedSuffix $ExpectedSuffix -ActualSuffix ([string]$suffix.Text)) {
            $extra = $Document.Range([int]$shape.Range.End, [int]$shape.Range.End + 1)
            $extra.Delete() | Out-Null
            Write-Host "Removed one verified extra period after inline MathType object at Start=$FormulaStart."
        }
    } finally {
        foreach ($item in @($extra, $suffix, $prefix, $paragraph, $shape, $probe)) { Release-ComObject $item }
    }
}

function Convert-OneFormula {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$TargetDocument,
        [Parameter(Mandatory = $true)]$Formula
    )
    $stage = "create target range"
    $targetRange = $null
    $watchInlineStyleDialog = $false
    $checkExtraPeriod = $false
    $expectedPrefix = ''; $expectedSuffix = ''
    try {
        $contentStart = [int]$TargetDocument.Content.Start
        $contentEnd = [int]$TargetDocument.Content.End
        $targetStart = [int]$Formula.Start
        $targetEnd = [int]($Formula.Start + $Formula.Length)
        if ($targetStart -lt $contentStart -or $targetEnd -gt $contentEnd -or $targetStart -ge $targetEnd) {
            throw "Target range [$targetStart,$targetEnd) is outside document content [$contentStart,$contentEnd)."
        }
        $targetRange = $TargetDocument.Range($targetStart, $targetEnd)

        $actualText = [string]$targetRange.Text
        if ($actualText -cne [string]$Formula.Literal) {
            throw "Target text does not match the expected formula. Expected=$($Formula.Literal); Actual=$actualText"
        }

        $formulaIsInTable = Test-FormulaRangeInTable -FormulaRange $targetRange
        # Do not make dialog protection depend on Word's table-range metadata.
        # MathType can request display conversion for inline formulas in narrow
        # cells and other layout contexts even when Word reports InTable=false.
        $watchInlineStyleDialog = ([string]$Formula.Kind -eq "inline")
        # Guard against the observed extra period after an inline ellipsis
        # conversion. Snapshot the complete paragraph sides before conversion
        # so a real source period can never be discarded.
        $checkExtraPeriod = $watchInlineStyleDialog -and ([string]$Formula.Literal -cmatch '\\ldots\b')
        if ($checkExtraPeriod) {
            $paragraph = $null; $prefix = $null; $suffix = $null
            try {
                $paragraph = $targetRange.Paragraphs.Item(1)
                $prefix = $TargetDocument.Range([int]$paragraph.Range.Start, $targetStart)
                $suffix = $TargetDocument.Range($targetEnd, [int]$paragraph.Range.End)
                $expectedPrefix = [string]$prefix.Text
                $expectedSuffix = [string]$suffix.Text
            } finally {
                Release-ComObject $suffix; Release-ComObject $prefix; Release-ComObject $paragraph
            }
        }

        # MathType 6 officially accepts $...$ as inline Texvc and \[...\] as
        # display Texvc. Convert the selected range in the target document itself;
        # do not create or copy an OLE object across documents.
        $beforeObjects = Get-MathTypeObjectCount -Document $TargetDocument
        $stage = "select target formula"
        $TargetDocument.Activate()
        $targetRange.Select()

        $stage = "run MathType Toggle TeX on target selection"
        Invoke-MathTypeToggleForFormula -Word $Word -WatchInlineStyleDialog $watchInlineStyleDialog -FormulaIsInTable $formulaIsInTable
        if ($DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }

        $stage = "validate MathType object count increment"
        $afterObjects = Wait-MathTypeObjectCount -Document $TargetDocument -ExpectedCount ($beforeObjects + 1)
        if ($afterObjects -ne ($beforeObjects + 1)) {
            # MathType 6 can occasionally return from its inline/display modal
            # without converting the selected range. Retry only when the exact
            # original text is still present, so a partially changed formula is
            # never processed twice.
            $stage = "retry unchanged formula after MathType modal"
            $sameRange = $TargetDocument.Range($targetStart, $targetStart + ([string]$Formula.Literal).Length)
            try {
                if ([string]$sameRange.Text -ceq [string]$Formula.Literal) {
                    Write-Host "MathType returned without inserting an object; retrying the unchanged formula once." -ForegroundColor Yellow
                    $TargetDocument.Activate()
                    $sameRange.Select()
                    Invoke-MathTypeToggleForFormula -Word $Word -WatchInlineStyleDialog $watchInlineStyleDialog -FormulaIsInTable $formulaIsInTable
                    $afterObjects = Wait-MathTypeObjectCount -Document $TargetDocument -ExpectedCount ($beforeObjects + 1)
                    if ($afterObjects -eq ($beforeObjects + 1)) {
                        $script:TransientFormulaRetries++
                        if ($checkExtraPeriod) {
                            Repair-InlineConversionExtraPeriod -Document $TargetDocument -FormulaStart $targetStart `
                                -ExpectedPrefix $expectedPrefix -ExpectedSuffix $expectedSuffix
                        }
                        return
                    }
                }
            } finally {
                Release-ComObject $sameRange
            }

            $stage = "validate MathType object count after transient retry"
            $retryLiterals = @(Get-FormulaRetryLiterals -Formula $Formula)
            if ($retryLiterals.Count -eq 0) {
                throw "Expected MathType object count to increase from $beforeObjects to $($beforeObjects + 1), but found $afterObjects after one safe retry."
            }

            Release-ComObject $targetRange
            $targetRange = $null
            $currentLiteral = [string]$Formula.Literal
            $convertedByFallback = $false
            foreach ($retryLiteral in $retryLiterals) {
                $stage = "prepare MathType compatibility fallback"
                $retryRange = $TargetDocument.Range($targetStart, $targetStart + $currentLiteral.Length)
                try {
                    if ([string]$retryRange.Text -cne $currentLiteral) {
                        throw "Formula text changed after failed Toggle TeX; refusing automatic rewrite. Current=$($retryRange.Text)"
                    }
                    $retryRange.Text = $retryLiteral
                    $retryRange.SetRange($targetStart, $targetStart + $retryLiteral.Length)
                    Write-Host "Retrying formula as: $retryLiteral" -ForegroundColor Yellow

                    $stage = "run MathType Toggle TeX on compatibility fallback"
                    $TargetDocument.Activate()
                    $retryRange.Select()
                    Invoke-MathTypeToggleForFormula -Word $Word -WatchInlineStyleDialog $watchInlineStyleDialog -FormulaIsInTable $formulaIsInTable
                    if ($DelayMilliseconds -gt 0) {
                        Start-Sleep -Milliseconds $DelayMilliseconds
                    }
                    $afterObjects = Wait-MathTypeObjectCount -Document $TargetDocument -ExpectedCount ($beforeObjects + 1)
                    if ($afterObjects -eq ($beforeObjects + 1)) {
                        $convertedByFallback = $true
                        $originalContent = ([string]$Formula.Literal).Substring(1, ([string]$Formula.Literal).Length - 2).Trim()
                        $primeSafeOriginal = ConvertTo-AsciiPrimeText -Text $originalContent
                        if ([regex]::IsMatch($primeSafeOriginal, "'+$")) {
                            $script:PrimeFallbackConversions++
                        } elseif ([regex]::IsMatch($originalContent, '^[A-Za-z][A-Za-z0-9_{}]*(?:\s*,\s*[A-Za-z][A-Za-z0-9_{}]*)+$')) {
                            $script:CommaListFallbackConversions++
                        } else {
                            $script:NumericFallbackConversions++
                        }
                        break
                    }
                    $currentLiteral = $retryLiteral
                } finally {
                    Release-ComObject $retryRange
                }
            }
            if (-not $convertedByFallback) {
                $stage = "restore formula after failed compatibility fallbacks"
                $restoreRange = $TargetDocument.Range($targetStart, $targetStart + $currentLiteral.Length)
                try {
                    if ([string]$restoreRange.Text -ceq $currentLiteral) {
                        $restoreRange.Text = [string]$Formula.Literal
                    }
                } finally {
                    Release-ComObject $restoreRange
                }
                throw "MathType ignored the original formula and all safe compatibility fallback forms."
            }
        }
        if ($checkExtraPeriod) {
            $stage = "check inline ellipsis conversion boundary"
            Repair-InlineConversionExtraPeriod -Document $TargetDocument -FormulaStart $targetStart `
                -ExpectedPrefix $expectedPrefix -ExpectedSuffix $expectedSuffix
        }
    } catch {
        $hresult = ""
        if ($_.Exception -is [System.Runtime.InteropServices.COMException]) {
            $hresult = " HResult=0x{0:X8}." -f ($_.Exception.HResult -band 0xffffffffL)
        }
        $sample = [string]$Formula.Literal
        if ($sample.Length -gt 160) { $sample = $sample.Substring(0, 157) + "..." }
        throw "Formula conversion failed during '$stage'.$hresult Kind=$($Formula.Kind); Start=$($Formula.Start); Length=$($Formula.Length); Formula=$sample; Error=$($_.Exception.Message)"
    } finally {
        Release-ComObject $targetRange
    }
}

function Normalize-ConvertedText {
    param([Parameter(Mandatory = $true)][string]$Text)
    $normalized = $Text.Replace([string][char]1, "[MATH]")
    return [regex]::Replace($normalized, '[\s\u00A0]+', '')
}

function Get-FormulaWordRanges {
    param([Parameter(Mandatory = $true)]$Document)
    $result = [System.Collections.Generic.List[object]]::new()
    for ($bookmarkIndex = 1; $bookmarkIndex -le $Document.Bookmarks.Count; $bookmarkIndex++) {
        $bookmark = $Document.Bookmarks.Item($bookmarkIndex)
        $range = $null
        try {
            if ([string]$bookmark.Name -notmatch '^M2WMath\d+$') { continue }
            $range = $bookmark.Range.Duplicate
            $text = [string]$range.Text
            $kind = if ($text.StartsWith('\[') -and $text.EndsWith('\]')) { 'display' }
                    elseif ($text.StartsWith('$') -and $text.EndsWith('$')) { 'inline' }
                    else { throw "Invalid formula bookmark: $($bookmark.Name)" }
            $result.Add([pscustomobject]@{
                Start = [int]$range.Start; Length = [int]($range.End - $range.Start)
                Kind = $kind; Literal = $text; Bookmark = [string]$bookmark.Name
            })
        } finally {
            Release-ComObject $range
            Release-ComObject $bookmark
        }
    }
    return @($result)
}

function Read-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $entry = $Archive.GetEntry($Name)
    if ($null -eq $entry) { return $null }
    $stream = $null
    $reader = $null
    try {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        return $reader.ReadToEnd()
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
    }
}

function Write-ZipEntryXml {
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)]$Xml
    )
    $oldEntry = $Archive.GetEntry($Name)
    if ($null -ne $oldEntry) { $oldEntry.Delete() }
    $entry = $Archive.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $null
    $writer = $null
    try {
        $stream = $entry.Open()
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $false
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        $Xml.Save($writer)
    } finally {
        if ($null -ne $writer) { $writer.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
    }
}

function Set-DocxMathTypeNumberFormat {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Value
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $customNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    $valueNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    $contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    $relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    $customRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)

        $customText = Read-ZipEntryText -Archive $archive -Name "docProps/custom.xml"
        $customXml = New-Object System.Xml.XmlDocument
        $customXml.PreserveWhitespace = $true
        if ([string]::IsNullOrWhiteSpace($customText)) {
            $declaration = $customXml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
            [void]$customXml.AppendChild($declaration)
            $root = $customXml.CreateElement("Properties", $customNamespace)
            $root.SetAttribute("xmlns:vt", $valueNamespace)
            [void]$customXml.AppendChild($root)
        } else {
            $customXml.LoadXml($customText)
        }
        $customNs = New-Object System.Xml.XmlNamespaceManager($customXml.NameTable)
        $customNs.AddNamespace("cp", $customNamespace)
        $property = $customXml.SelectSingleNode("/cp:Properties/cp:property[@name='MTEquationNumber2']", $customNs)
        if ($null -eq $property) {
            $property = $customXml.CreateElement("property", $customNamespace)
            $property.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            $maxPid = 1
            foreach ($existing in $customXml.SelectNodes("/cp:Properties/cp:property", $customNs)) {
                $candidatePid = 0
                if ([int]::TryParse($existing.GetAttribute("pid"), [ref]$candidatePid) -and $candidatePid -gt $maxPid) {
                    $maxPid = $candidatePid
                }
            }
            $property.SetAttribute("pid", [string]($maxPid + 1))
            $property.SetAttribute("name", "MTEquationNumber2")
            [void]$customXml.DocumentElement.AppendChild($property)
        } else {
            while ($property.HasChildNodes) { [void]$property.RemoveChild($property.FirstChild) }
        }
        $valueNode = $customXml.CreateElement("vt", "lpwstr", $valueNamespace)
        $valueNode.InnerText = $Value
        [void]$property.AppendChild($valueNode)
        Write-ZipEntryXml -Archive $archive -Name "docProps/custom.xml" -Xml $customXml

        $contentTypesText = Read-ZipEntryText -Archive $archive -Name "[Content_Types].xml"
        if ([string]::IsNullOrWhiteSpace($contentTypesText)) { throw "DOCX is missing [Content_Types].xml." }
        $contentTypesXml = New-Object System.Xml.XmlDocument
        $contentTypesXml.PreserveWhitespace = $true
        $contentTypesXml.LoadXml($contentTypesText)
        $contentNs = New-Object System.Xml.XmlNamespaceManager($contentTypesXml.NameTable)
        $contentNs.AddNamespace("ct", $contentTypesNamespace)
        $override = $contentTypesXml.SelectSingleNode("/ct:Types/ct:Override[@PartName='/docProps/custom.xml']", $contentNs)
        if ($null -eq $override) {
            $override = $contentTypesXml.CreateElement("Override", $contentTypesNamespace)
            $override.SetAttribute("PartName", "/docProps/custom.xml")
            $override.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
            [void]$contentTypesXml.DocumentElement.AppendChild($override)
            Write-ZipEntryXml -Archive $archive -Name "[Content_Types].xml" -Xml $contentTypesXml
        }

        $relsText = Read-ZipEntryText -Archive $archive -Name "_rels/.rels"
        if ([string]::IsNullOrWhiteSpace($relsText)) { throw "DOCX is missing _rels/.rels." }
        $relsXml = New-Object System.Xml.XmlDocument
        $relsXml.PreserveWhitespace = $true
        $relsXml.LoadXml($relsText)
        $relsNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $relsNs.AddNamespace("r", $relationshipsNamespace)
        $relationship = $relsXml.SelectSingleNode("/r:Relationships/r:Relationship[@Type='$customRelationshipType']", $relsNs)
        if ($null -eq $relationship) {
            $usedIds = @{}
            foreach ($existing in $relsXml.SelectNodes("/r:Relationships/r:Relationship", $relsNs)) {
                $usedIds[[string]$existing.GetAttribute("Id")] = $true
            }
            $nextId = 1
            while ($usedIds.ContainsKey("rId$nextId")) { $nextId++ }
            $relationship = $relsXml.CreateElement("Relationship", $relationshipsNamespace)
            $relationship.SetAttribute("Id", "rId$nextId")
            $relationship.SetAttribute("Type", $customRelationshipType)
            $relationship.SetAttribute("Target", "docProps/custom.xml")
            [void]$relsXml.DocumentElement.AppendChild($relationship)
            Write-ZipEntryXml -Archive $archive -Name "_rels/.rels" -Xml $relsXml
        }
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Invoke-MathTypeNumberMacro {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Purpose
    )
    try {
        # Word.Application.Run exposes MacroName plus 30 optional Variant
        # arguments. With a late-bound COM object, OptionalParamBinding alone
        # does not reliably marshal the omitted Variants and can raise
        # DISP_E_PARAMNOTOPTIONAL. Pass Missing for every unused slot; unlike
        # $null, Missing is not forwarded to the zero-argument MathType macro.
        $flags = [System.Reflection.BindingFlags]::InvokeMethod -bor [System.Reflection.BindingFlags]::OptionalParamBinding
        [object[]]$arguments = [object[]]::new(31)
        $arguments[0] = $Name
        for ($argumentIndex = 1; $argumentIndex -lt $arguments.Length; $argumentIndex++) {
            $arguments[$argumentIndex] = [System.Type]::Missing
        }
        $null = $Word.GetType().InvokeMember("Run", $flags, $null, $Word, $arguments)
    } catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException) { $message = $_.Exception.InnerException.Message }
        throw "MathType $Purpose macro failed: $Name. Error=$message"
    }
}

function Get-DocumentMathTypeNumberCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.Fields.Count; $i++) {
        $field = $Document.Fields.Item($i)
        try {
            if ([string]$field.Code.Text -match '(?i)^\s*MACROBUTTON\s+MTPlaceRef(?:\s|$)') { $count++ }
        } finally { Release-ComObject $field }
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
        } finally { Release-ComObject $field }
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
        } finally { Release-ComObject $field }
    }
    return $count
}

function Get-ParagraphContentEndForNumbering {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $range = $Paragraph.Range.Duplicate
    try {
        $end = [int]$range.End
        $text = [string]$range.Text
        for ($i = $text.Length - 1; $i -ge 0; $i--) {
            $code = [int][char]$text[$i]
            if ($code -eq 13 -or $code -eq 7) { $end-- } else { break }
        }
        return $end
    } finally { Release-ComObject $range }
}

function Get-UsableWidthForNumberedParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $section = $Paragraph.Range.Sections.Item(1)
    $pageSetup = $null
    try {
        $pageSetup = $section.PageSetup
        $width = [double]$pageSetup.PageWidth - [double]$pageSetup.LeftMargin - [double]$pageSetup.RightMargin
        if ($width -le 0) { throw "Invalid PageSetup text width: $width" }
        return $width
    } finally {
        Release-ComObject $pageSetup
        Release-ComObject $section
    }
}

function Get-DisplayMathTypeParagraphNearPosition {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][int]$Position
    )
    $contentStart = [int]$Document.Content.Start
    $contentEnd = [int]$Document.Content.End
    $seenStarts = @{}
    foreach ($offset in @(0, 1, -1, 2, -2, 3, -3)) {
        $probePosition = [Math]::Max($contentStart, [Math]::Min($contentEnd - 1, $Position + $offset))
        $probe = $Document.Range($probePosition, $probePosition)
        $paragraph = $null
        $shape = $null
        try {
            $paragraph = $probe.Paragraphs.Item(1)
            $paragraphStart = [int]$paragraph.Range.Start
            if ($seenStarts.ContainsKey($paragraphStart)) { continue }
            $seenStarts[$paragraphStart] = $true
            $shape = Get-MathTypeShapeInParagraph -Paragraph $paragraph
            if ($null -ne $shape -and (Test-DisplayEquationParagraph -Paragraph $paragraph)) {
                $result = $paragraph
                $paragraph = $null
                return $result
            }
        } finally {
            Release-ComObject $shape
            Release-ComObject $paragraph
            Release-ComObject $probe
        }
    }
    return $null
}

function Add-MathTypeNativeNumberAtFormula {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)][int]$FormulaStart
    )
    $paragraph = $null
    $shape = $null
    try {
        $paragraph = Get-DisplayMathTypeParagraphNearPosition -Document $Document -Position $FormulaStart
        if ($null -eq $paragraph) {
            throw "Could not locate the converted MTDisplayEquation paragraph near Start=$FormulaStart."
        }
        $shape = Get-MathTypeShapeInParagraph -Paragraph $paragraph
        if ($null -eq $shape) {
            throw "Converted display formula paragraph does not contain exactly one Equation.DSMT4 object. Start=$FormulaStart"
        }

        if ($paragraph.Range.End -lt $Document.Content.End) {
            $nextAnchor = $Document.Range([int]$paragraph.Range.End, [int]$paragraph.Range.End)
            $nextParagraph = $null
            try {
                $nextParagraph = $nextAnchor.Paragraphs.Item(1)
                if (Test-EmptyParagraph -Paragraph $nextParagraph) {
                    $nextParagraph.Range.Delete() | Out-Null
                    $script:DisplayEmptyParagraphsRemoved++
                }
            } finally {
                Release-ComObject $nextParagraph
                Release-ComObject $nextAnchor
            }
        }

        $contentEnd = Get-ParagraphContentEndForNumbering -Paragraph $paragraph
        if ([int]$shape.Range.End -lt $contentEnd) {
            $suffix = $Document.Range([int]$shape.Range.End, $contentEnd)
            try {
                if (-not [string]::IsNullOrWhiteSpace([string]$suffix.Text)) {
                    throw "Display equation paragraph contains non-whitespace text after the formula. Start=$FormulaStart; Text=$($suffix.Text)"
                }
                $suffix.Text = ""
            } finally { Release-ComObject $suffix }
        }
        $shapeStart = [int]$shape.Range.Start
        if ([int]$paragraph.Range.Start -lt $shapeStart) {
            $prefix = $Document.Range([int]$paragraph.Range.Start, $shapeStart)
            try {
                if (-not [string]::IsNullOrWhiteSpace([string]$prefix.Text)) {
                    throw "Display equation paragraph contains non-whitespace text before the formula. Start=$FormulaStart; Text=$($prefix.Text)"
                }
                $prefix.Text = "`t"
            } finally { Release-ComObject $prefix }
        } else {
            $prefix = $Document.Range($shapeStart, $shapeStart)
            try { $prefix.Text = "`t" } finally { Release-ComObject $prefix }
        }

        $width = Get-UsableWidthForNumberedParagraph -Paragraph $paragraph
        $paragraph.Format.Alignment = $wdAlignParagraphLeft
        $paragraph.Format.LeftIndent = 0
        $paragraph.Format.RightIndent = 0
        $paragraph.Format.FirstLineIndent = 0
        $paragraph.Format.TabStops.ClearAll()
        $centerTab = $paragraph.Format.TabStops.Add($width / 2.0, $wdAlignTabCenter, $wdTabLeaderSpaces)
        $rightTab = $paragraph.Format.TabStops.Add($width, $wdAlignTabRight, $wdTabLeaderSpaces)
        Release-ComObject $centerTab
        Release-ComObject $rightTab
        $paragraph.Format.LineSpacingRule = $wdLineSpaceAtLeast
        $paragraph.Format.LineSpacing = $DisplayMinPt
        $paragraph.Format.SpaceBefore = $DisplaySpaceBeforePt
        $paragraph.Format.SpaceAfter = $DisplaySpaceAfterPt
        $paragraph.Format.KeepTogether = $true
        $script:DisplayParagraphsReformatted++

        Release-ComObject $shape
        $shape = $paragraph.Range.InlineShapes.Item(1)
        $before = Get-DocumentMathTypeNumberCount -Document $Document
        $insert = $Document.Range([int]$shape.Range.End, [int]$shape.Range.End)
        try {
            $insert.Text = "`t"
            $insert.Collapse($wdCollapseEnd)
            $insert.Select()
            Invoke-WithRetry -Operation {
                Invoke-MathTypeNumberMacro -Word $Word -Name $InsertNumberMacroName -Purpose "Insert Equation Number"
            } | Out-Null
        } finally { Release-ComObject $insert }
        $after = Get-DocumentMathTypeNumberCount -Document $Document
        if ($after -ne ($before + 1)) {
            throw "MathType did not create exactly one native MTPlaceRef number. Before=$before; After=$after; Start=$FormulaStart"
        }
        $script:MathTypeNativeNumbersInserted++
    } finally {
        Release-ComObject $shape
        Release-ComObject $paragraph
    }
}

function Get-DocumentTextWithoutMathTypeManagedFieldResults {
    param([Parameter(Mandatory = $true)]$Document, [switch]$ReplaceFormulas)
    $spans = [System.Collections.Generic.List[object]]::new()
    if ($ReplaceFormulas) {
        foreach ($formula in @(Get-FormulaWordRanges -Document $Document)) {
            $spans.Add([pscustomobject]@{ Start=$formula.Start; End=($formula.Start + $formula.Length); Replacement='[MATH]' })
        }
    }
    for ($i = 1; $i -le $Document.Fields.Count; $i++) {
        $field = $Document.Fields.Item($i)
        try {
            $code = [string]$field.Code.Text
            # MathType owns and may update both kinds of field results:
            # - MTPlaceRef: the visible equation number, e.g. (1)
            # - MTEditEquationSection: the hidden chapter/section break state
            # They are metadata, not source-document prose, and must be removed
            # from both sides of the non-formula text integrity comparison.
            if ($code -match '(?i)^\s*MACROBUTTON\s+(?:MTPlaceRef|MTEditEquationSection)(?:\s|$)') {
                $spans.Add([pscustomobject]@{
                    # Include the opening and closing field characters.
                    Start = [int]$field.Code.Start - 1
                    End = [int]$field.Result.End + 1
                    Replacement = ''
                })
            }
        } finally { Release-ComObject $field }
    }
    $builder = [System.Text.StringBuilder]::new()
    $cursor = [int]$Document.Content.Start
    foreach ($span in @($spans | Sort-Object Start, End)) {
        if ($span.Start -gt $cursor) {
            $gap = $Document.Range($cursor, $span.Start)
            try { [void]$builder.Append([string]$gap.Text) } finally { Release-ComObject $gap }
        }
        if ($span.End -gt $cursor) {
            [void]$builder.Append([string]$span.Replacement)
            $cursor = $span.End
        }
    }
    $documentEnd = [int]$Document.Content.End
    if ($cursor -lt $documentEnd) {
        $tail = $Document.Range($cursor, $documentEnd)
        try { [void]$builder.Append([string]$tail.Text) } finally { Release-ComObject $tail }
    }
    return $builder.ToString()
}

function Get-TextDifferenceSummary {
    param(
        [Parameter(Mandatory = $true)][string]$Expected,
        [Parameter(Mandatory = $true)][string]$Actual
    )
    $commonLength = [Math]::Min($Expected.Length, $Actual.Length)
    $differenceIndex = 0
    while ($differenceIndex -lt $commonLength -and $Expected[$differenceIndex] -ceq $Actual[$differenceIndex]) {
        $differenceIndex++
    }
    $contextStart = [Math]::Max(0, $differenceIndex - 30)
    $expectedContextLength = [Math]::Min(70, $Expected.Length - $contextStart)
    $actualContextLength = [Math]::Min(70, $Actual.Length - $contextStart)
    $expectedContext = if ($expectedContextLength -gt 0) { $Expected.Substring($contextStart, $expectedContextLength) } else { "<end>" }
    $actualContext = if ($actualContextLength -gt 0) { $Actual.Substring($contextStart, $actualContextLength) } else { "<end>" }
    return "Index=$differenceIndex; ExpectedLength=$($Expected.Length); ActualLength=$($Actual.Length); ExpectedContext=[$expectedContext]; ActualContext=[$actualContext]"
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
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
if (-not (Test-Path -LiteralPath $InputPath)) {
    throw "Input file not found: $InputPath"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inputDirectory = [System.IO.Path]::GetDirectoryName($InputPath)
    $inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $inputExtension = [System.IO.Path]::GetExtension($InputPath)
    $OutputPath = Join-Path $inputDirectory ($inputBaseName + '_mathtype' + $inputExtension)
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if ($InputPath -eq $OutputPath) {
    throw "InputPath and OutputPath must be different."
}
if (Test-Path -LiteralPath $OutputPath) {
    if (-not $Force) {
        throw "Output already exists: $OutputPath. Use -Force to replace it."
    }
}
if ($env:USERNAME -eq "codexsandboxonline") {
    throw "Run this script in your normal Windows PowerShell session, not the Codex terminal."
}

$outputDir = Split-Path -Parent $OutputPath
if ($outputDir) { [void][System.IO.Directory]::CreateDirectory($outputDir) }
Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
if ($AddEquationNumbers) {
    # This non-empty custom property suppresses MathType's first-use format
    # dialog. Create the complete OPC registration when the template has never
    # used native MathType numbering before.
    Set-DocxMathTypeNumberFormat -Path $OutputPath -Value $requestedNumberFormat
}

$word = $null
$document = $null
$tempDocument = $null

try {
    $word = Get-RunningWord
    Write-Host "Attached to: $($word.Name) $($word.Version)" -ForegroundColor Cyan

    $tempDocument = $word.Documents.Add()
    Write-Host "Preflight: converting inline and display formula smoke tests" -ForegroundColor Cyan
    $smokeShape = Put-FormulaInTemporaryDocument -Word $word -TempDocument $tempDocument -Literal '$x^2$'
    Release-ComObject $smokeShape
    Clear-TemporaryDocument -Document $tempDocument
    $smokeDisplayShape = Put-FormulaInTemporaryDocument -Word $word -TempDocument $tempDocument -Literal '\[x^2\]'
    Release-ComObject $smokeDisplayShape
    Clear-TemporaryDocument -Document $tempDocument

    $document = Invoke-WithRetry -Operation {
        $word.Documents.Open($OutputPath, $false, $false, $false)
    }
    if ($document.ReadOnly) { throw 'Output document was opened read-only; close the document and retry.' }
    $document.Activate()

    if ($AddEquationNumbers) {
        $legacyNumbers = Get-DocumentLegacySequenceEquationCount -Document $document
        if ($legacyNumbers -gt 0) {
            throw "Legacy SEQ Equation fields were found (Count=$legacyNumbers). They are not MathType-native; use an unnumbered source document."
        }
        if ((Get-DocumentMathTypeNumberCount -Document $document) -gt 0) {
            throw "The input already contains MathType MTPlaceRef equation numbers. Direct numbering requires an unnumbered source document."
        }
        if ((Get-MathTypeNumberingSetupCount -Document $document) -eq 0) {
            Write-Host "Initializing MathType-native equation numbering metadata" -ForegroundColor Cyan
            $documentStart = $document.Range(0, 0)
            try { $documentStart.Select() } finally { Release-ComObject $documentStart }
            Invoke-WithRetry -Operation {
                Invoke-MathTypeNumberMacro -Word $word -Name $InitializeNumbersMacroName -Purpose "Initialize Equation Numbering"
            } | Out-Null
        }
    }

    $originalParagraphs = $document.Paragraphs.Count
    $originalTables = $document.Tables.Count
    $initialMathTypeObjects = Get-MathTypeObjectCount -Document $document
    $originalText = [string]$document.Content.Text
    $sourceIntegrityText = Get-DocumentTextWithoutMathTypeManagedFieldResults -Document $document -ReplaceFormulas
    $sourceFingerprint = Normalize-ConvertedText -Text $sourceIntegrityText

    $formulas = @(Get-FormulaWordRanges -Document $document)
    $displayMatches = @($formulas | Where-Object Kind -eq 'display')
    $inlineMatches = @($formulas | Where-Object Kind -eq 'inline')
    if ($ExpectedInline -ge 0 -and $inlineMatches.Count -ne $ExpectedInline) {
        throw "Formula count mismatch before conversion. Inline=$($inlineMatches.Count)/$ExpectedInline, Display=$($displayMatches.Count)/$ExpectedDisplay"
    }
    if ($ExpectedDisplay -ge 0 -and $displayMatches.Count -ne $ExpectedDisplay) {
        throw "Formula count mismatch before conversion. Inline=$($inlineMatches.Count)/$ExpectedInline, Display=$($displayMatches.Count)/$ExpectedDisplay"
    }
    if ($ExpectedInline -lt 0) { $ExpectedInline = $inlineMatches.Count }
    if ($ExpectedDisplay -lt 0) { $ExpectedDisplay = $displayMatches.Count }
    if (($ExpectedInline + $ExpectedDisplay) -le 0) {
        throw 'No formulas were found. Required delimiters are $...$ for inline formulas and \[...\] for display formulas.'
    }
    Write-Host ("Detected formulas: {0} inline, {1} display" -f $ExpectedInline, $ExpectedDisplay) -ForegroundColor Green

    Write-Host "Resolving LaTeX ranges in Word-native coordinates" -ForegroundColor Cyan
    $resolvedInline = @($formulas | Where-Object Kind -eq 'inline').Count
    $resolvedDisplay = @($formulas | Where-Object Kind -eq 'display').Count
    if ($resolvedInline -ne $ExpectedInline -or $resolvedDisplay -ne $ExpectedDisplay) {
        throw "Formula count mismatch after resolving Word ranges. Inline=$resolvedInline/$ExpectedInline, Display=$resolvedDisplay/$ExpectedDisplay"
    }
    $ordered = @($formulas | Sort-Object Start -Descending)

    Write-Host "Converting $($ordered.Count) formulas individually. Do not interact with Word." -ForegroundColor Cyan
    for ($index = 0; $index -lt $ordered.Count; $index++) {
        if ($env:M2W_CANCEL_FILE -and (Test-Path -LiteralPath $env:M2W_CANCEL_FILE)) { throw 'Conversion cancelled by user.' }
        $formula = $ordered[$index]
        $formulaTimer = [Diagnostics.Stopwatch]::StartNew()
        $done = $index + 1
        if ($done -eq 1 -or $done % 10 -eq 0) {
            Write-Host ("Starting {0}/{1}: {2}, Start={3}, Length={4}" -f $done, $ordered.Count, $formula.Kind, $formula.Start, $formula.Length)
        }
        Convert-OneFormula -Word $word -TargetDocument $document -Formula $formula
        if ($AddEquationNumbers -and $formula.Kind -eq 'display') {
            Add-MathTypeNativeNumberAtFormula -Document $document -Word $word -FormulaStart $formula.Start
        }
        if ($done -eq 1 -or $done % 10 -eq 0 -or $done -eq $ordered.Count) {
            Write-Host ("Converted {0}/{1}" -f $done, $ordered.Count)
        }
        $formulaTimer.Stop()
        Write-Host ('M2W_EVENT ' + (@{type='progress'; stage='mathtype'; current=$done; total=$ordered.Count; elapsed_ms=$formulaTimer.ElapsedMilliseconds} | ConvertTo-Json -Compress))
    }

    if ($AddEquationNumbers) {
        Write-Host "Updating MathType-native equation numbers" -ForegroundColor Cyan
        $collapsedSelection = $document.Range(0, 0)
        try { $collapsedSelection.Select() } finally { Release-ComObject $collapsedSelection }
        Invoke-WithRetry -Operation {
            Invoke-MathTypeNumberMacro -Word $word -Name $UpdateNumbersMacroName -Purpose "Update Equation Numbers"
        } | Out-Null
        $document.Fields.Update() | Out-Null
    } else {
        Write-Host "Repairing display-equation paragraph layout" -ForegroundColor Cyan
        Repair-DisplayEquationLayout -Document $document
    }

    $mathTypeObjects = @(Get-MathTypeInlineShapes -Document $document).Count
    $convertedText = [string]$document.Content.Text
    # Literal TeX in source code is intentional. Per-target object checks plus
    # the fingerprint below audit all AST-approved formulas without scanning code.
    $remainingInline = 0
    $remainingDisplay = 0
    $expectedFinalMathTypeObjects = $initialMathTypeObjects + $ExpectedInline + $ExpectedDisplay
    if ($mathTypeObjects -ne $expectedFinalMathTypeObjects) {
        throw "Expected $expectedFinalMathTypeObjects MathType objects, found $mathTypeObjects."
    }
    if ($remainingInline -ne 0 -or $remainingDisplay -ne 0) {
        throw "Residual LaTeX remains. Inline=$remainingInline, Display=$remainingDisplay"
    }
    if ($document.OMaths.Count -ne 0) {
        throw "Word-native OMath objects were found: $($document.OMaths.Count)"
    }
    $mathTypeNativeNumbers = Get-DocumentMathTypeNumberCount -Document $document
    if ($AddEquationNumbers -and $mathTypeNativeNumbers -ne $ExpectedDisplay) {
        throw "MathType native equation number count mismatch. Fields=$mathTypeNativeNumbers; Display=$ExpectedDisplay"
    }
    if (-not $AddEquationNumbers -and $mathTypeNativeNumbers -ne 0) {
        throw "Unexpected pre-existing MathType equation numbers were found: $mathTypeNativeNumbers"
    }

    $finalParagraphs = $document.Paragraphs.Count
    $paragraphDelta = $finalParagraphs - $originalParagraphs
    if ($document.Tables.Count -ne $originalTables) {
        throw "Table structure changed during conversion. Tables $originalTables->$($document.Tables.Count)."
    }
    if ($paragraphDelta -lt 0 -or $paragraphDelta -gt $ExpectedDisplay) {
        throw "Unexpected paragraph change during conversion. Paragraphs $originalParagraphs->$finalParagraphs (delta=$paragraphDelta; allowed=0..$ExpectedDisplay)."
    }
    if ($paragraphDelta -gt 0) {
        Write-Host ("MathType added {0} display-equation paragraph boundaries (allowed up to {1})." -f $paragraphDelta, $ExpectedDisplay) -ForegroundColor Yellow
    }
    $integrityText = if ($AddEquationNumbers) {
        Get-DocumentTextWithoutMathTypeManagedFieldResults -Document $document
    } else {
        $convertedText
    }
    $convertedFingerprint = Normalize-ConvertedText -Text $integrityText
    if ($sourceFingerprint -cne $convertedFingerprint) {
        $difference = Get-TextDifferenceSummary -Expected $sourceFingerprint -Actual $convertedFingerprint
        throw "Normalized non-formula text changed during conversion. $difference."
    }

    $document.Save()
    Write-Host "Success" -ForegroundColor Green
    [pscustomobject]@{
        InputPath = $InputPath
        OutputPath = $OutputPath
        MathTypeObjects = $mathTypeObjects
        InitialMathTypeObjects = $initialMathTypeObjects
        DetectedInlineLatex = $ExpectedInline
        DetectedDisplayLatex = $ExpectedDisplay
        EquationNumberFormat = if ($AddEquationNumbers) { $EquationNumberFormat } else { "Disabled" }
        MathTypeNativeNumbers = $mathTypeNativeNumbers
        MathTypeNativeNumbersInserted = $script:MathTypeNativeNumbersInserted
        NumericFallbackConversions = $script:NumericFallbackConversions
        PrimeFallbackConversions = $script:PrimeFallbackConversions
        CommaListFallbackConversions = $script:CommaListFallbackConversions
        TransientFormulaRetries = $script:TransientFormulaRetries
        TableInlineDialogWatchers = $script:TableInlineDialogWatchers
        TableInlineDialogsHandled = $script:TableInlineDialogsHandled
        DisplayParagraphsReformatted = $script:DisplayParagraphsReformatted
        DisplayEmptyParagraphsRemoved = $script:DisplayEmptyParagraphsRemoved
        RemainingInlineLatex = $remainingInline
        RemainingDisplayLatex = $remainingDisplay
        WordOMathObjects = $document.OMaths.Count
        Paragraphs = $document.Paragraphs.Count
        Tables = $document.Tables.Count
        TextIntegrityChecked = $true
    }
} catch {
    $conversionError = $_
    Write-UserActionError `
        -Title "DOCX LaTeX to MathType conversion failed." `
        -Reason $conversionError.Exception.Message `
        -Actions @(
            "InputPath : $InputPath",
            "OutputPath: $OutputPath",
            "Restart conversion from the original Markdown after fixing the reported error.",
            "For delimiter issues, ensure inline formulas are `$...`$ and display formulas are \[...\]; do not use `$\$...`$\$.",
            "For Word/MathType issues, leave Word open, verify Toggle TeX manually, and rerun."
        )
    throw $conversionError
} finally {
    if ($null -ne $document) {
        try { $document.Close($wdDoNotSaveChanges) } catch {}
        Release-ComObject $document
    }
    if ($null -ne $tempDocument) {
        try { $tempDocument.Close($wdDoNotSaveChanges) } catch {}
        Release-ComObject $tempDocument
    }
    if ($null -ne $word) {
        try { $word.Visible = $true } catch {}
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
