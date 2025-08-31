# monitor.ps1 � one-shot snapshot with printed sections + JSON output
# REQUIREMENTS:
# - OpenHardwareMonitorLib.dll installed and unblocked
# - Run 64-bit PowerShell (System32), preferably as Administrator

# --- CONFIG ---
$OHM_DLL   = "C:\Program Files (x86)\OpenHardwareMonitor\OpenHardwareMonitorLib.dll"
$JsonOut   = "$PSScriptRoot\hw-snapshot.json"
$ShowFans  = $true   # set $false to hide fan list
$LoopEverySeconds = 0  # set to >0 (e.g., 5) for a live loop that refreshes every N seconds

# --- LOAD OHM ---
if (-not (Test-Path $OHM_DLL)) { throw "OpenHardwareMonitorLib.dll not found at $OHM_DLL" }
[void][Reflection.Assembly]::LoadFrom($OHM_DLL)

# --- INIT OHM ---
$computer = New-Object OpenHardwareMonitor.Hardware.Computer
$computer.CPUEnabled           = $true
$computer.GPUEnabled           = $true
$computer.MainboardEnabled     = $true
$computer.RAMEnabled           = $true
$computer.HDDEnabled           = $true
$computer.FanControllerEnabled = $true
$computer.Open()

function Update-HW([OpenHardwareMonitor.Hardware.IHardware]$h){ $h.Update(); foreach($s in $h.SubHardware){ Update-HW $s } }
function Get-Sensors {
  param(
    [OpenHardwareMonitor.Hardware.IHardware]$hw,
    [string[]]$Types = @(),
    [string[]]$NameContains = @()
  )
  $s = @()
  if ($hw.Sensors) { $s += $hw.Sensors }
  foreach($sub in $hw.SubHardware){ if ($sub.Sensors) { $s += $sub.Sensors } }
  if ($Types.Count) { $s = $s | Where-Object { $Types -contains $_.SensorType.ToString() } }
  if ($NameContains.Count) {
    $s = $s | Where-Object {
      $n = $_.Name.ToLower()
      ($NameContains | Where-Object { $n -like "*$($_.ToLower())*" }).Count -gt 0
    }
  }
  $s
}

function Get-Snapshot {
  foreach($h in $computer.Hardware){ Update-HW $h }

  $roots = $computer.Hardware

  # CPU
  $cpuOut = foreach($c in $roots | Where-Object { $_.HardwareType.ToString() -eq "CPU" }){
    $temps = Get-Sensors $c -Types Temperature
    $clks  = Get-Sensors $c -Types Clock
    [pscustomobject]@{
      Name        = $c.Name
      PackageTemp = ($temps | Where-Object Name -match 'package|cpu temp' | Select-Object -ExpandProperty Value -First 1)
      CoreTemps   = @($temps | Where-Object Name -match 'core' | ForEach-Object { @{Name=$_.Name; C=[math]::Round($_.Value,1)} })
      CoreClocks  = @($clks  | Where-Object Name -match 'core' | ForEach-Object { @{Name=$_.Name; MHz=[math]::Round($_.Value,0)} })
    }
  }

  # Fans (mainboard + fan controller)
  $fanOut = @()
  foreach($mb in $roots | Where-Object { $_.HardwareType.ToString() -in @("Mainboard","SuperIO","Cooler") }){
    $fanOut += Get-Sensors $mb -Types Fan | ForEach-Object {
      [pscustomobject]@{ Sensor=$_.Name; RPM=[math]::Round($_.Value,0) }
    }
  }

  # GPU
  $gpuOut = foreach($g in $roots | Where-Object { $_.HardwareType.ToString() -match "Gpu" }){
    $t = Get-Sensors $g -Types Temperature
    $k = Get-Sensors $g -Types Clock
    $f = Get-Sensors $g -Types Fan
    [pscustomobject]@{
      Name   = $g.Name
      TempC  = ($t | Select-Object -ExpandProperty Value -First 1)
      Clocks = @($k | ForEach-Object { @{Name=$_.Name; MHz=[math]::Round($_.Value,0)} })
      Fans   = @($f | ForEach-Object { @{Name=$_.Name; RPM=[math]::Round($_.Value,0)} })
    }
  }

  # RAM (from OHM)
  $ram = ($roots | Where-Object { $_.HardwareType.ToString() -eq "RAM" } | Select-Object -First 1)
  $ramOut = $null
  if ($ram) {
    $ramLoad = (Get-Sensors $ram -Types Load | Where-Object Name -match 'memory|ram' | Select-Object -ExpandProperty Value -First 1)
    $ramData = (Get-Sensors $ram -Types Data)
    $ramOut = [pscustomobject]@{
      LoadPercent = [math]::Round($ramLoad,1)
      UsedGB      = [math]::Round(($ramData | Where-Object Name -match 'used'      | Select-Object -ExpandProperty Value -First 1),2)
      AvailableGB = [math]::Round(($ramData | Where-Object Name -match 'available' | Select-Object -ExpandProperty Value -First 1),2)
      TotalGB     = [math]::Round(($ramData | Where-Object Name -match 'total'     | Select-Object -ExpandProperty Value -First 1),2)
    }
  }

  # Storage devices (temps + any load/data exposed)
  $diskOut = foreach($d in $roots | Where-Object { $_.HardwareType.ToString() -in @("HDD","Storage") }){
    $temps = Get-Sensors $d -Types Temperature
    $loads = ($d.Sensors + ($d.SubHardware|ForEach-Object{$_.Sensors})) | Where-Object { $_.SensorType.ToString() -in @("Load","Data") }
    [pscustomobject]@{
      Device  = $d.Name
      TempC   = [math]::Round(($temps | Select-Object -ExpandProperty Value -First 1),1)
      Metrics = @($loads | ForEach-Object { @{ Name=$_.Name; Type=$_.SensorType.ToString(); Value=[math]::Round($_.Value,2) } })
    }
  }

  # Network throughput (instantaneous)
  $net = (Get-Counter "\Network Interface(*)\Bytes Total/sec").CounterSamples | ForEach-Object {
    [pscustomobject]@{ Interface=$_.InstanceName; BytesPerSec=[math]::Round($_.CookedValue,0) }
  }

  # Top processes (CPU time and Working Set)
  $procs  = Get-Process | Where-Object { $null -ne $_.CPU }
  $topCPU = $procs | Sort-Object CPU -Descending | Select-Object -First 10 Name,Id,@{n='CPUSeconds';e={[math]::Round($_.CPU,1)}},@{n='WorkingSetMB';e={[math]::Round($_.WorkingSet64/1MB,1)}}
  $topMem = $procs | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 Name,Id,@{n='WorkingSetMB';e={[math]::Round($_.WorkingSet64/1MB,1)}},@{n='CPUSeconds';e={[math]::Round($_.CPU,1)}}

  # Assemble
  [pscustomobject]@{
    Timestamp = (Get-Date).ToString("s")
    CPU       = $cpuOut
    Fans      = $fanOut
    GPU       = $gpuOut
    RAM       = $ramOut
    Storage   = $diskOut
    Network   = $net
    TopCPU    = $topCPU
    TopMemory = $topMem
  }
}

function Write-Report($result){
  "=== CPU ==="
  $result.CPU | ForEach-Object {
    "CPU: $($_.Name)"
    if ($null -ne $_.PackageTemp) { "  Package Temp: $([math]::Round($_.PackageTemp,1)) �C" }
    if ($_.CoreTemps) { $_.CoreTemps | ForEach-Object { "  $($_.Name): $($_.C) �C" } }
    if ($_.CoreClocks){ $_.CoreClocks| ForEach-Object { "  $($_.Name): $($_.MHz) MHz" } }
  }
  if ($ShowFans -and $result.Fans) {
    "=== FANS ==="
    $result.Fans | Sort-Object Sensor | Format-Table -AutoSize
  }
  "=== GPU ==="
  $result.GPU | ForEach-Object {
    "GPU: $($_.Name)"
    "  Temp: $([math]::Round($_.TempC,1)) �C"
    if ($_.Clocks){ $_.Clocks | ForEach-Object { "  $($_.Name): $($_.MHz) MHz" } }
    if ($_.Fans)  { $_.Fans   | ForEach-Object { "  Fan $($_.Name): $($_.RPM) RPM" } }
  }
  "=== RAM ==="
  $result.RAM | Format-List
  "=== STORAGE ==="
  $result.Storage | ForEach-Object {
    "[$($_.Device)] Temp: $($_.TempC) �C"
    if ($_.Metrics) { $_.Metrics | Format-Table -AutoSize }
  }
  "=== NETWORK (Bytes/sec) ==="
  $result.Network | Sort-Object -Property BytesPerSec -Descending | Format-Table -AutoSize
  "=== TOP 10 by CPU time ==="
  $result.TopCPU | Format-Table -AutoSize
  "=== TOP 10 by Working Set ==="
  $result.TopMemory | Format-Table -AutoSize
}

# --- RUN ---
do {
  $snapshot = Get-Snapshot
  Clear-Host
  "Snapshot: $($snapshot.Timestamp)"
  Write-Report $snapshot
  $snapshot | ConvertTo-Json -Depth 8 | Out-File -Encoding UTF8 $JsonOut
  "JSON written to: $JsonOut"
  if ($LoopEverySeconds -gt 0) { Start-Sleep -Seconds $LoopEverySeconds }
} while ($LoopEverySeconds -gt 0)
