$ErrorActionPreference = "Stop"
$start = Get-Date
$url = "http://127.0.0.1:8000/api/status"

while ($true) {
    try {
        $s = Invoke-RestMethod -Uri $url -TimeoutSec 5
    } catch {
        Write-Output ("[{0}] status request failed: {1}" -f (Get-Date -Format "HH:mm:ss"), $_.Exception.Message)
        Start-Sleep -Seconds 5
        continue
    }

    $w = $s.warmup
    $prog = if ($null -ne $w -and $null -ne $w.progress) { [int]$w.progress } else { -1 }
    $phase = if ($null -ne $w -and $null -ne $w.phase) { [string]$w.phase } else { "unknown" }
    $eta = if ($null -ne $w -and $null -ne $w.eta_s) { [int]$w.eta_s } else { -1 }
    $msg = if ($null -ne $w -and $null -ne $w.message) { [string]$w.message } else { "" }
    $elapsed = [int]((Get-Date) - $start).TotalSeconds

    Write-Output ("[{0}] ready={1} phase={2} progress={3}% eta={4}s elapsed={5}s msg={6}" -f (Get-Date -Format "HH:mm:ss"), $s.ready, $phase, $prog, $eta, $elapsed, $msg)

    if ($s.ready -eq $true) {
        Write-Output ("DONE total_elapsed_s={0}" -f $elapsed)
        break
    }
    Start-Sleep -Seconds 10
}
