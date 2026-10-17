$ErrorActionPreference = 'Continue'
$src = 'C:\Users\金东权\wps-skills\wps-multitool-mcp'
$dst = 'C:\Users\金东权\wps-skills\wps-multitool-mcp'

Write-Output '--- 占用进程 ---'
Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*wps-multitool-mcp*' } | ForEach-Object {
  Write-Output ("{0}  {1}" -f $_.ProcessId, $_.Name)
}

Start-Sleep -Seconds 3
Write-Output '--- 尝试改名 ---'
try {
  Rename-Item -Path $src -NewName 'wps-multitool-mcp' -ErrorAction Stop
  Write-Output 'RENAMED'
} catch {
  Write-Output ('FAILED: ' + $_.Exception.Message)
}
if (Test-Path $dst) { Write-Output 'DST-EXISTS' } else { Write-Output 'DST-MISSING' }
