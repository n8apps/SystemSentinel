# OHM-Dump-Short.ps1 — concise one-shot sensor check (PowerShell 5.1 safe)
# Defaults to OpenHardwareMonitorLib.dll; switch to LibreHardwareMonitor by changing $OHM_DLL.

# --- CONFIG ---
$OHM_DLL = "C:\Program Files (x86)\OpenHardwareMonitor\OpenHardwareMonitorLib.dll"
# $OHM_DLL = "C:\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll"  # alternative

# --- Checks (non-fatal except missing DLL) ---
if (-not [Environment]::Is64BitProcess) { Write-Warning "Not 64-bit PowerShell; sensors may fail." }

# --- LOAD OHM ---
if (-not (Test-Path $OHM_DLL)) { throw "OpenHardwareMonitorLib.dll not found at $OHM_DLL" }
[void][Reflection.Assembly]::LoadFrom($OHM_DLL)

# --- INIT OHM ---
$computer = New-Object OpenHardwareMonitor.Hardware.Computer
$computer.CPUEnabled=$true; $computer.GPUEnabled=$true; $computer.MainboardEnabled=$true
$computer.RAMEnabled=$true; $computer.HDDEnabled=$true; $computer.FanControllerEnabled=$true
$computer.Open()

function Update-HW([OpenHardwareMonitor.Hardware.IHardware]$h){
  $h.Update(); foreach($s in $h.SubHardware){ $s.Update() }
}
function Get-Sensors([OpenHardwareMonitor.Hardware.IHardware]$hw,[string[]]$Types){
  $s=@(); if($hw.Sensors){$s+=$hw.Sensors}; foreach($sub in $hw.SubHardware){ if($sub.Sensors){ $s+=$sub.Sensors } }
  if($Types){ $s = $s | Where-Object { $Types -contains $_.SensorType.ToString() } }
  return $s
}

# --- Refresh once ---
foreach($h in $computer.Hardware){ Update-HW $h }
$roots = $computer.Hardware

# ---- CPU ----
$cpu = $roots | Where-Object { $_.HardwareType.ToString() -eq "CPU" } | Select-Object -First 1
if ($cpu) {
  $cpuLoad = (Get-Sensors $cpu @("Load") | Where-Object Name -match 'CPU Total' | Select-Object -ExpandProperty Value -First 1)
  $cpuTemp = (Get-Sensors $cpu @("Temperature") | Select-Object -ExpandProperty Value -First 1)

  $cpuLoadText = if ($null -ne $cpuLoad) { "{0:N1}%" -f $cpuLoad } else { "n/a" }
  $cpuTempText = if ($null -ne $cpuTemp) { "{0:N1}°C" -f $cpuTemp } else { "n/a" }

  "CPU:   Load=$cpuLoadText  Temp=$cpuTempText" | Write-Output
} else {
  "CPU:   n/a" | Write-Output
}

# ---- GPU ----
$gpu = $roots | Where-Object { $_.HardwareType.ToString() -like "Gpu*" } | Select-Object -First 1
if ($gpu) {
  $gTemp = (Get-Sensors $gpu @("Temperature") | Select-Object -ExpandProperty Value -First 1)
  $gFan  = (Get-Sensors $gpu @("Fan")         | Select-Object -ExpandProperty Value -First 1)
  $gPow  = (Get-Sensors $gpu @("Power")       | Select-Object -ExpandProperty Value -First 1)
  $gUsed = (Get-Sensors $gpu @("SmallData")   | Where-Object Name -match 'GPU Memory Used'  | Select-Object -ExpandProperty Value -First 1)
  $gTot  = (Get-Sensors $gpu @("SmallData")   | Where-Object Name -match 'GPU Memory Total' | Select-Object -ExpandProperty Value -First 1)

  $gTempText = if ($null -ne $gTemp) { "{0:N0}°C" -f $gTemp } else { "n/a" }
  $gFanText  = if ($null -ne $gFan) { "{0:N0} RPM" -f $gFan } else { "n/a" }
  $gPowText  = if ($null -ne $gPow) { "{0:N1} W" -f $gPow } else { "n/a" }
  $vramText  = if ($null -ne $gUsed -and $null -ne $gTot) { "{0:N0}/{1:N0} MB" -f $gUsed,$gTot } else { "n/a" }

  "GPU:   Temp=$gTempText  Fan=$gFanText  Power=$gPowText  VRAM=$vramText" | Write-Output
} else {
  "GPU:   n/a" | Write-Output
}

# ---- RAM ----
$ram = $roots | Where-Object { $_.HardwareType.ToString() -eq "RAM" } | Select-Object -First 1
if ($ram) {
  $rLoad = (Get-Sensors $ram @("Load") | Select-Object -ExpandProperty Value -First 1)
  $rUsed = (Get-Sensors $ram @("Data") | Where-Object Name -match 'Used Memory'      | Select-Object -ExpandProperty Value -First 1)
  $rAvail= (Get-Sensors $ram @("Data") | Where-Object Name -match 'Available Memory' | Select-Object -ExpandProperty Value -First 1)
  $rTot  = if ($null -ne $rUsed -and $null -ne $rAvail) { $rUsed + $rAvail } else { $null }

  $rLoadText = if ($null -ne $rLoad) { "{0:N1}%" -f $rLoad } else { "n/a" }
  $ramText   = if ($null -ne $rTot) { "{0:N1}/{1:N1} GB" -f $rUsed,$rTot } else { "n/a" }

  "RAM:   Load=$rLoadText  Used/Total=$ramText" | Write-Output
} else {
  "RAM:   n/a" | Write-Output
}

# ---- Storage ----
$drives = $roots | Where-Object { $_.HardwareType.ToString() -in @("HDD","Storage") }
if ($drives) {
  foreach($d in $drives) {
    $temp = (Get-Sensors $d @("Temperature") | Select-Object -ExpandProperty Value -First 1)
    if ($null -ne $temp) {
      $tempText = "{0:N0}°C" -f $temp
      "Disk:  $($d.Name)  Temp=$tempText" | Write-Output
    } else {
      $used = (Get-Sensors $d @("Load") | Where-Object Name -match 'Used Space' | Select-Object -ExpandProperty Value -First 1)
      $usedText = if ($null -ne $used) { "{0:N0}%" -f $used } else { "n/a" }
      "Disk:  $($d.Name)  Used=$usedText" | Write-Output
    }
  }
} else {
  "Storage: n/a" | Write-Output
}

# ---- Fans (from mainboard/SuperIO if present) ----
$mbRoots = $roots | Where-Object { $_.HardwareType.ToString() -in @("Mainboard","SuperIO","Cooler") }
$fans = @(); foreach($m in $mbRoots){ $fans += Get-Sensors $m @("Fan") }
if ($fans -and $fans.Count -gt 0) {
  $list = ($fans | Select-Object -First 4 | ForEach-Object { "{0}:{1:N0}RPM" -f $_.Name, ($_.Value -as [double]) }) -join "  "
  "Fans:  $list" | Write-Output
} else {
  "Fans:  n/a" | Write-Output
}

# ---- Network (safe) ----
try {
  $net = (Get-Counter "\Network Interface(*)\Bytes Total/sec").CounterSamples
  $sum = if($net){ ($net | Measure-Object -Property CookedValue -Sum).Sum } else { 0 }
  $sumText = "{0:N0} B/s total" -f [double]$sum
  "Net:   $sumText" | Write-Output
} catch {
  "Net:   n/a" | Write-Output
}

$computer.Close()
