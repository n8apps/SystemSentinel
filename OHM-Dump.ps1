# OHM-Dump.ps1 — list every hardware + sensor + current value

# Pick one DLL (leave the other line commented)
$OHM_DLL = "C:\Program Files (x86)\OpenHardwareMonitor\OpenHardwareMonitorLib.dll"
# $OHM_DLL = "C:\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll"

if (-not [Environment]::Is64BitProcess) { throw "Run in 64-bit PowerShell (System32)!" }

$admin = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $admin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Warning "Not running as Administrator — some sensors may be missing."
}

if (-not (Test-Path $OHM_DLL)) { throw "DLL not found at: $OHM_DLL" }
try { Unblock-File -Path $OHM_DLL -ErrorAction SilentlyContinue } catch {}

[void][Reflection.Assembly]::LoadFrom($OHM_DLL)

$computer = New-Object OpenHardwareMonitor.Hardware.Computer
$computer.CPUEnabled=$true; $computer.GPUEnabled=$true; $computer.MainboardEnabled=$true
$computer.RAMEnabled=$true; $computer.HDDEnabled=$true; $computer.FanControllerEnabled=$true
$computer.Open()

function Update-HW([OpenHardwareMonitor.Hardware.IHardware]$h){
  $h.Update(); foreach($s in $h.SubHardware){ Update-HW $s }
}

foreach($h in $computer.Hardware){
  Update-HW $h
  Write-Host ""
  Write-Host ("=== [{0}] {1}" -f $h.HardwareType,$h.Name) -ForegroundColor Cyan
  $sensors=@()
  if ($h.Sensors) { $sensors += $h.Sensors }
  foreach($sub in $h.SubHardware){ if ($sub.Sensors){ $sensors += $sub.Sensors } }
  if (-not $sensors) { Write-Host "  (no sensors)"; continue }
  $sensors | Sort-Object SensorType, Name | ForEach-Object {
    $val = if ($_.Value -ne $null) { $_.Value } else { "(null)" }
    "{0,-12}  {1,-40}  {2}" -f $_.SensorType, $_.Name, $val
  } | Write-Host
}
$computer.Close()