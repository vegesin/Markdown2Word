param([switch]$CheckOnly)

function Get-DesktopWordExecutable {
    # App Paths can be absent on otherwise healthy Click-to-Run installations.
    $candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
        foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
            $base = $null; $key = $null; $clickToRun = $null
            try {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $key = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE')
                if ($key) {
                    $value = [string]$key.GetValue('')
                    if ($value) { $candidates.Add($value.Trim('"')) }
                }
                $clickToRun = $base.OpenSubKey('Software\Microsoft\Office\ClickToRun\Configuration')
                if ($clickToRun) {
                    $installation = [string]$clickToRun.GetValue('InstallationPath')
                    if ($installation) {
                        $candidates.Add((Join-Path $installation 'root\Office16\WINWORD.EXE'))
                        $candidates.Add((Join-Path $installation 'Office16\WINWORD.EXE'))
                    }
                }
            } catch {
                # A registry view may not exist or be readable on this machine.
            } finally {
                if ($key) { $key.Dispose() }
                if ($clickToRun) { $clickToRun.Dispose() }
                if ($base) { $base.Dispose() }
            }
        }
    }
    $command = Get-Command WINWORD.EXE -ErrorAction SilentlyContinue
    if ($command) { $candidates.Add($command.Source) }
    foreach ($root in @($env:ProgramW6432, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $root) { continue }
        foreach ($relative in @('Microsoft Office\root\Office16', 'Microsoft Office\Office16', 'Microsoft Office\Office15', 'Microsoft Office\Office14')) {
            $candidates.Add((Join-Path $root "$relative\WINWORD.EXE"))
        }
    }
    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return [IO.Path]::GetFullPath($candidate) }
    }
    return $null
}

function Connect-DesktopWord {
    param([switch]$StartIfNeeded)
    try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
    catch { $application = $null }
    if ($null -eq $application -and $StartIfNeeded) {
        $existing = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
        if ($existing.Count -eq 0) {
            $executable = Get-DesktopWordExecutable
            if (-not $executable) { throw 'Microsoft Word desktop executable was not found. Install Microsoft Word or choose native Word equations.' }
            Start-Process -FilePath $executable -ArgumentList '/q' -WindowStyle Hidden
            Write-Host "Started desktop Microsoft Word: $executable"
        }
        $deadline = [DateTime]::UtcNow.AddSeconds(30)
        do {
            try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch {}
            if ($null -ne $application) { break }
            Start-Sleep -Milliseconds 250
        } while ([DateTime]::UtcNow -lt $deadline)
    }
    if ($null -eq $application) {
        throw 'Desktop Word is not available. Check startup dialogs and use the same Windows privilege level as the GUI.'
    }
    if ($application.Name -notlike '*Microsoft Word*' -or $application.Path -match 'WPS|Kingsoft') {
        throw 'The registered Word application is not Microsoft Word.'
    }
    Write-Host "Connected to Microsoft Word $($application.Version): $($application.Path)"
    return $application
}

if ($CheckOnly) {
    [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
    $wordPath = Get-DesktopWordExecutable
    $mathTypeRegistered = $null -ne [Type]::GetTypeFromProgID('Equation.DSMT4')
    [pscustomobject]@{ word_path=$wordPath; mathtype_registered=$mathTypeRegistered } | ConvertTo-Json -Compress
}
