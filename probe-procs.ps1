Get-CimInstance Win32_Process | Where-Object {
  ($_.CommandLine -like '*wps-multitool-mcp*') -or
  ($_.CommandLine -like '*wps-comd.ps1*') -or
  ($_.CommandLine -like '*bridge/index.js*')
} | ForEach-Object {
  $cmd = ''
  if ($_.CommandLine) { $cmd = $_.CommandLine }
  if ($cmd.Length -gt 110) { $cmd = $cmd.Substring(0, 110) }
  Write-Output ("{0}`t{1}`t{2}" -f $_.ProcessId, $_.Name, $cmd)
}
