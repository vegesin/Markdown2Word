function Connect-DesktopWord {
    param([switch]$StartIfNeeded)
    try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
    catch { $application = $null }
    if ($null -eq $application -and $StartIfNeeded) {
        $existing = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
        if ($existing.Count -eq 0) {
            $executable = $null
            foreach ($key in @(
                'HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
                'HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
            )) {
                if (Test-Path $key) {
                    $candidate = (Get-Item $key).GetValue('')
                    if ($candidate -and (Test-Path -LiteralPath $candidate)) { $executable = $candidate; break }
                }
            }
            if (-not $executable) { throw 'Microsoft Word desktop executable was not found in App Paths.' }
            Start-Process -FilePath $executable -ArgumentList '/q' -WindowStyle Minimized
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
