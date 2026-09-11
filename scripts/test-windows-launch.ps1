param([Parameter(Mandatory = $true)][string]$ApplicationDirectory)
$ErrorActionPreference = 'Stop'
$testDirectory = Join-Path $env:RUNNER_TEMP ('pictos-launch-' + [Guid]::NewGuid().ToString('N'))
$env:PICTOS_DATA_DIR = $testDirectory
$program = Join-Path $ApplicationDirectory 'Neurow.Pictos.exe'
$nativeProcess = $null
$initialPairingHash = $null
try {
    for ($iteration = 0; $iteration -lt 2; $iteration++) {
        $nativeProcess = Start-Process -FilePath $program -ArgumentList '--startup' -PassThru
        $deadline = [DateTime]::UtcNow.AddSeconds(30)
        $health = $null
        while ([DateTime]::UtcNow -lt $deadline) {
            try { $health = Invoke-RestMethod 'http://127.0.0.1:43129/health' -TimeoutSec 1; break } catch { Start-Sleep -Milliseconds 250 }
        }
        if ($health.application -ne 'atelier-pictos') { throw 'The packaged application did not start its local server.' }
        $nativeProcess.Refresh()
        if ($nativeProcess.HasExited -or $nativeProcess.MainWindowHandle -ne [IntPtr]::Zero) { throw 'The launcher exited or created a visible main window.' }
        $pairingFile = Join-Path $testDirectory 'pairing.dat'
        $pairingHash = (Get-FileHash $pairingFile -Algorithm SHA256).Hash
        if ($iteration -eq 0) { $initialPairingHash = $pairingHash }
        elseif ($pairingHash -ne $initialPairingHash) { throw 'The protected pairing record changed after restarting.' }
        Add-Type -AssemblyName System.Security
        $secret = [Security.Cryptography.ProtectedData]::Unprotect([IO.File]::ReadAllBytes($pairingFile), [Text.Encoding]::UTF8.GetBytes('Neurow.Pictos.Pairing.v1'), [Security.Cryptography.DataProtectionScope]::CurrentUser)
        $token = ([BitConverter]::ToString($secret)).Replace('-', '').ToLowerInvariant()
        $state = Invoke-RestMethod 'http://127.0.0.1:43129/v1/connection' -Headers @{ Authorization = ('Bearer ' + $token) } -TimeoutSec 35
        [Array]::Clear($secret, 0, $secret.Length); $token = $null
        if ($state.version -ne '1.1.0' -or $state.connected -ne $false -or $state.pairingPersistent -ne $true) { throw 'The packaged Codex connection did not match the isolated test profile.' }
        Stop-Process -Id $nativeProcess.Id -Force
        $nativeProcess.WaitForExit(10000) | Out-Null
        $nativeProcess = $null
        $stopped = $false
        $deadline = [DateTime]::UtcNow.AddSeconds(10)
        while ([DateTime]::UtcNow -lt $deadline) {
            try { Invoke-RestMethod 'http://127.0.0.1:43129/health' -TimeoutSec 1 | Out-Null; Start-Sleep -Milliseconds 250 }
            catch { $stopped = $true; break }
        }
        if (-not $stopped) { throw 'The local server survived the launcher shutdown.' }
    }
    Write-Output 'Packaged Windows application: two hidden launches, real Codex account reads without login or inference, stable DPAPI pairing, and complete shutdown passed.'
}
finally {
    if ($null -ne $nativeProcess -and -not $nativeProcess.HasExited) { Stop-Process -Id $nativeProcess.Id -Force }
    if (Test-Path $testDirectory) { Remove-Item -LiteralPath $testDirectory -Recurse -Force }
}
