# SystemSentinel.ps1 � windowed, killable live dashboard (PowerShell 5.1 safe)
# Uses OpenHardwareMonitorLib.dll (or LibreHardwareMonitorLib.dll) for sensors.
# Hotkeys: Esc / Ctrl+Q = Exit, F11 = Fullscreen toggle.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Native kill fallback (kernel32 TerminateProcess)
try {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class NativeKill {
  [DllImport("kernel32.dll")] public static extern IntPtr GetCurrentProcess();
  [DllImport("kernel32.dll", SetLastError=true)] public static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);
}
"@
} catch {}

# ===== CONFIG =====
# Choose ONE of these:
$OHM_DLL = "C:\Program Files (x86)\OpenHardwareMonitor\OpenHardwareMonitorLib.dll"
# $OHM_DLL = "C:\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll"   # <-- recommended for new Ryzen temps

$REFRESH_MS = 1000        # update rate (ms)
$START_FULLSCREEN = $false
$ALWAYS_ON_TOP = $true    # set $false if you don't want it pinned
$TARGET_MONITOR_INDEX = 0 # 0 = primary, 1 = secondary, etc.

# External config
$cfgPath = Join-Path (Get-Location) 'config.json'
if (-not (Test-Path $cfgPath)) {
  $defaultCfg = @{ FontSize = 10; ChartColor = 'DodgerBlue'; Theme = 'light'; NetworkFontScale = 3 }
  $defaultCfg | ConvertTo-Json | Out-File -Encoding UTF8 $cfgPath
}
$cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
if (-not $cfg) { $cfg = [pscustomobject]@{ FontSize = 10; ChartColor = 'DodgerBlue'; Theme = 'light'; NetworkFontScale = 3 } }

# ===== RUNTIME CHECKS =====
if (-not [Environment]::Is64BitProcess) { Write-Warning "Not 64-bit PowerShell; sensors may fail." }
$admin = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $admin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Warning "Not running as Administrator � some sensors may be missing." }
if (-not (Test-Path $OHM_DLL)) { throw "Hardware monitor DLL not found at '$OHM_DLL'." }
try { Unblock-File -Path $OHM_DLL -ErrorAction SilentlyContinue } catch {}

# ===== LOAD & INIT OHM =====
[void][Reflection.Assembly]::LoadFrom($OHM_DLL)

$computer = New-Object OpenHardwareMonitor.Hardware.Computer
$computer.CPUEnabled           = $true
$computer.GPUEnabled           = $true
$computer.MainboardEnabled     = $true
$computer.RAMEnabled           = $true
$computer.HDDEnabled           = $true
$computer.FanControllerEnabled = $true
$computer.Open()

# ===== LOGGING =====
$logPath = Join-Path (Get-Location) 'monitor.log'
try { Set-Content -Path $logPath -Encoding UTF8 -Value "" } catch {}
function Write-Log([string]$message,[object]$ex=$null){
  try {
    $line = "{0} {1}" -f (Get-Date).ToString('s'), $message
    Add-Content -Path $logPath -Value $line
    if ($ex) { Add-Content -Path $logPath -Value ($ex | Out-String) }
  } catch {}
}
Write-Log "SystemSentinel started"

# Brutal self-termination utility
function Kill-Now {
  try { [System.Environment]::FailFast("UserExit") } catch {}
  try { [System.Diagnostics.Process]::GetCurrentProcess().Kill() } catch {}
  try { Stop-Process -Id $PID -Force } catch {}
  try { [NativeKill]::TerminateProcess([NativeKill]::GetCurrentProcess(),0) | Out-Null } catch {}
  try { Start-Process -FilePath "taskkill" -ArgumentList "/F","/PID", $PID -WindowStyle Hidden } catch {}
}

function Update-HW([OpenHardwareMonitor.Hardware.IHardware]$h){
  $h.Update()
  foreach($s in $h.SubHardware){ $s.Update() }
}
function Get-AllSensors([OpenHardwareMonitor.Hardware.IHardware]$hw){
  $s=@()
  if($hw.Sensors){ $s += $hw.Sensors }
  foreach($sub in $hw.SubHardware){ if($sub.Sensors){ $s += $sub.Sensors } }
  return $s
}
function Get-Sensors([OpenHardwareMonitor.Hardware.IHardware]$hw,[string[]]$Types){
  $s = Get-AllSensors $hw
  if($Types){ $s = $s | Where-Object { $Types -contains $_.SensorType.ToString() } }
  return $s
}
function Set-Bar($meter,[double]$val,[double]$max=100){
  $pct = 0
  if ($max -ne 0) { $pct = [math]::Min(100,[math]::Max(0,[math]::Round(($val/$max)*100))) }
  $meter.Bar.Value = [int]$pct
  $pct
}

# Parse color strings: named colors or hex (#RRGGBB or #AARRGGBB)
function Resolve-Color([object]$value,[System.Drawing.Color]$fallback){
  try {
    if ($value -is [string] -and $value.Trim().Length -gt 0) {
      $s = $value.Trim()
      if ($s.StartsWith('#')) {
        return [System.Drawing.ColorTranslator]::FromHtml($s)
      } else {
        $c = [System.Drawing.Color]::FromName($s)
        if ($c.IsKnownColor -or $c.IsNamedColor -or $c.IsSystemColor) { return $c }
      }
    }
  } catch {}
  return $fallback
}

# Human-friendly size formatter (bytes -> KB/MB/GB/TB)
function Format-Size([double]$bytes){
  if ($bytes -lt 0) { return "n/a" }
  $units = @("B","KB","MB","GB","TB","PB")
  $i = 0
  while ($bytes -ge 1024 -and $i -lt ($units.Count-1)) { $bytes /= 1024; $i++ }
  return ("{0:N0} {1}" -f $bytes, $units[$i])
}

# Get drive usage for a root like "C:\"; returns object with Used, Total, Percent
function Get-DriveUsage([string]$root){
  try {
    # Normalize the root path to match DriveInfo format (single backslash)
    $normalizedRoot = $root.TrimEnd('\')
    $drive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady -and $_.Name.TrimEnd('\') -ieq $normalizedRoot }
    if (-not $drive) { return $null }
    $total = [double]$drive.TotalSize
    $free  = [double]$drive.TotalFreeSpace
    $used  = $total - $free
    [pscustomobject]@{ Root=$root; Used=$used; Total=$total; Percent=([math]::Round(($used/$total)*100,0)) }
  } catch { return $null }
}

# ===== UI BASE =====
$screens=[System.Windows.Forms.Screen]::AllScreens
$screen=$screens[[Math]::Min([int]$TARGET_MONITOR_INDEX,$screens.Count-1)]

$form  = New-Object System.Windows.Forms.Form
$appName = if($cfg.AppName){ [string]$cfg.AppName } else { "System Sentinel" }
$form.Text          = $appName
$form.StartPosition = 'Manual'
$lightBg = [System.Drawing.Color]::FromArgb(0xF5,0xF6,0xFA)
$lightFg = [System.Drawing.Color]::FromArgb(0x22,0x22,0x22)
$darkBg  = [System.Drawing.Color]::FromArgb(0x15,0x18,0x1E)
$darkFg  = [System.Drawing.Color]::FromArgb(0xE6,0xE8,0xEE)
$useDark = ($cfg.Theme -and $cfg.Theme.ToString().ToLower() -eq 'dark')
$bgColor = $lightBg; if($useDark){ $bgColor = $darkBg }
$fgColor = $lightFg; if($useDark){ $fgColor = $darkFg }
$fontColor = Resolve-Color ($cfg.Font.Color) ($fgColor)
$form.BackColor     = $bgColor
$form.ForeColor     = $fontColor

# Background image support
if ($cfg.BgImage -and $cfg.BgImage.Display -eq $true -and $cfg.BgImage.Path) {
  try {
    $imagePath = Join-Path (Get-Location) $cfg.BgImage.Path
    if (Test-Path $imagePath) {
      $bgImage = [System.Drawing.Image]::FromFile($imagePath)
      $form.BackgroundImage = $bgImage
      $form.BackgroundImageLayout = 'Stretch'
      Write-Log "Background image loaded: $imagePath"
    } else {
      Write-Log "Background image not found: $imagePath"
    }
  } catch {
    Write-Log "Error loading background image" $_
  }
}

$defaultFontSize = if ($cfg.Font -and $cfg.Font.DefaultSize) { [float]$cfg.Font.DefaultSize } else { 10.0 }
$form.Font          = New-Object System.Drawing.Font('Segoe UI', $defaultFontSize)
$form.KeyPreview    = $true
$form.TopMost       = $ALWAYS_ON_TOP
$form.FormBorderStyle = 'Sizable'
$form.ShowInTaskbar = $true

# Override default close behavior to force kill
$form.add_FormClosing({
  try { $_.Cancel = $true } catch {}
  $script:shouldExit = $true
  try { if ($timer) { $timer.Stop(); $timer.Dispose() } } catch {}
  try { if ($tray) { $tray.Visible = $false } } catch {}
  Kill-Now
})

# default window size and center on target monitor; allow config AppWindowSize override
if ($cfg.AppWindowSize -and $cfg.AppWindowSize.Width -gt 0 -and $cfg.AppWindowSize.Height -gt 0) {
  $winW = [int][Math]::Min([int]$cfg.AppWindowSize.Width,  $screen.WorkingArea.Width  - 120)
  $winH = [int][Math]::Min([int]$cfg.AppWindowSize.Height, $screen.WorkingArea.Height - 120)
} else {
  $winW = [int][Math]::Min(1024, $screen.WorkingArea.Width  - 120)
  $winH = [int][Math]::Min(600 , $screen.WorkingArea.Height - 120)
}
$form.Bounds = [System.Drawing.Rectangle]::new(
  [int]($screen.WorkingArea.X + ($screen.WorkingArea.Width - $winW)/2),
  [int]($screen.WorkingArea.Y + ($screen.WorkingArea.Height- $winH)/2),
  $winW,$winH
)

# fullscreen helpers
$state = [ordered]@{
  IsFullscreen = $false
  PrevBounds   = $form.Bounds
  PrevBorder   = $form.FormBorderStyle
  PrevTopMost  = $form.TopMost
}
function Enter-Fullscreen {
  $state.PrevBounds=$form.Bounds; $state.PrevBorder=$form.FormBorderStyle; $state.PrevTopMost=$form.TopMost
  $form.FormBorderStyle='None'; $form.TopMost=$true; $form.Bounds=$screen.Bounds; $state.IsFullscreen=$true
}
function Exit-Fullscreen {
  $form.FormBorderStyle=$state.PrevBorder; $form.TopMost=$state.PrevTopMost; $form.Bounds=$state.PrevBounds; $state.IsFullscreen=$false
}
if($START_FULLSCREEN){ Enter-Fullscreen }

# Global exit flag for clean shutdown
$script:shouldExit = $false

# hotkeys
$form.add_KeyDown({
  if ($_.KeyCode -eq 'Escape' -or ($_.Control -and $_.KeyCode -eq 'Q')) { 
    $script:shouldExit = $true
    $timer.Stop()
    $form.Close() 
  }
  elseif ($_.KeyCode -eq 'F11') { if($state.IsFullscreen){ Exit-Fullscreen } else { Enter-Fullscreen } }
})

# tray icon
$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon = [System.Drawing.SystemIcons]::Application
$tray.Visible = $true
$tray.Text = $appName
$ctx = New-Object System.Windows.Forms.ContextMenuStrip
$miExit = $ctx.Items.Add("Exit")
$miExit.Add_Click({ 
  $script:shouldExit = $true
  $timer.Stop()
  $form.Close() 
})
$tray.ContextMenuStrip = $ctx
$tray.Add_DoubleClick({ if($state.IsFullscreen){ Exit-Fullscreen } else { Enter-Fullscreen } })


$form.add_FormClosed({ 
  # Backup kill - if we somehow get here, force exit
  $script:shouldExit = $true
  Kill-Now
})

# header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 36
if ($useDark) {
  $header.BackColor = [System.Drawing.Color]::FromArgb(35,38,46)
} else {
$header.BackColor = [System.Drawing.Color]::FromArgb(230,232,238)
}
$form.Controls.Add($header) | Out-Null

$ttl = New-Object System.Windows.Forms.Label
$ttl.AutoSize = $true
$ttl.Location = [System.Drawing.Point]::new(10,9)
if ($cfg.Header -and $cfg.Header.Display -ne $false) {
  $text = if ($cfg.Header.Text -and $cfg.Header.Text.Trim().Length -gt 0) { [string]$cfg.Header.Text } else { "System Sentinel - Esc/Ctrl+Q exit, F11 fullscreen" }
  $ttl.Text = $text
  $header.Visible = $true
} else {
  $ttl.Text = ""
  $header.Visible = $false
}
$header.Controls.Add($ttl) | Out-Null

# simple layout helpers
function New-Card([string]$title,[int]$x,[int]$y,[int]$w,[int]$h){
  $panel = New-Object System.Windows.Forms.Panel
  $panel.BackColor = [System.Drawing.Color]::White
  if ($useDark) { $panel.BackColor = [System.Drawing.Color]::FromArgb(28,31,38) }
  $panel.Location  = [System.Drawing.Point]::new($x,$y)
  $panel.Size      = [System.Drawing.Size]::new($w,$h)
  $panel.Padding   = [System.Windows.Forms.Padding]::new(12,10,12,10)

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.Text = $title
  $lbl.Font = New-Object System.Drawing.Font('Segoe UI Semibold', [float]([math]::Max(8, $cfg.FontSize + 1)))
  $lbl.AutoSize = $true
  $lbl.Location = [System.Drawing.Point]::new(10,6)
  $panel.Controls.Add($lbl) | Out-Null

  $form.Controls.Add($panel) | Out-Null
  return $panel
}
function New-Meter([System.Windows.Forms.Control]$parent,[string]$labelText,[int]$y){
  $lbl = New-Object System.Windows.Forms.Label
  $lbl.Text=$labelText; $lbl.AutoSize=$true
  $lbl.Location=[System.Drawing.Point]::new(12,$y)
  $parent.Controls.Add($lbl) | Out-Null

  $val = New-Object System.Windows.Forms.Label
  $val.AutoSize=$true
  $val.TextAlign='TopLeft'
  $val.Location=[System.Drawing.Point]::new(12,$y+18)
  $parent.Controls.Add($val) | Out-Null

  $bar = New-Object System.Windows.Forms.ProgressBar
  $bar.Style='Continuous'; $bar.Minimum=0; $bar.Maximum=100
  $bar.Size=[System.Drawing.Size]::new([int]($parent.Width-24),16)
  $bar.Location=[System.Drawing.Point]::new(12,$y+40)
  $parent.Controls.Add($bar) | Out-Null

  return @{ Label=$lbl; Bar=$bar; Value=$val }
}

# place cards
$colSpace = if($cfg.Spacing -and $cfg.Spacing.Column){ [int]$cfg.Spacing.Column } else { 16 }
$rowSpace = if($cfg.Spacing -and $cfg.Spacing.Row){ [int]$cfg.Spacing.Row } else { 16 }
$availW = $form.ClientSize.Width
$topY   = if($header.Visible){ $header.Bottom + 8 } else { 8 }
# compute widths/heights: prefer pixel overrides, otherwise percentage
$baseRow1W  = [int](($availW - ($colSpace*3)) / 2)
$baseRow2W  = [int](($availW - ($colSpace*4)) / 3)
$baseFullW  = [int]($availW - ($colSpace*2))

function Resolve-BoxSize($boxCfg,$baseW){
  # Pixels-only: honor WidthPx/HeightPx, otherwise fall back to baseW and a sane default height
  $w = if ($boxCfg -and $boxCfg.WidthPx -gt 0) { [int]$boxCfg.WidthPx } else { [int]$baseW }
  $h = if ($boxCfg -and $boxCfg.HeightPx -gt 0) { [int]$boxCfg.HeightPx } else { 160 }

  return @{ W=$w; H=$h }
}

$cpuSize = Resolve-BoxSize $cfg.CpuBox $baseRow1W
$gpuSize = Resolve-BoxSize $cfg.GpuBox $baseRow1W
$ramSize = Resolve-BoxSize $cfg.RamBox $baseRow2W
$netSize = Resolve-BoxSize $cfg.NetworkBox $baseRow2W
$stoSize = Resolve-BoxSize $cfg.StorageBox $baseRow2W
$topSize = Resolve-BoxSize $cfg.TopProcessesBox $baseFullW

$row1W = [int]$cpuSize.W
$row2W = [int]$ramSize.W
$fullW = [int]$topSize.W
$cardH = [int][Math]::Max([int]$cpuSize.H,[int]$gpuSize.H)
# centered X origins so right edges align per row
$row1LeftX = [int]([math]::Round(($availW - ($row1W*2 + $colSpace)) / 2))
$row2LeftX = [int]([math]::Round(($availW - ($row2W*3 + $colSpace*2)) / 2))
$row3LeftX = [int]([math]::Round(($availW - $fullW) / 2))

# Row 1: CPU (left) and GPU (right)
$cpuCard = New-Card "CPU" $row1LeftX $topY $cpuSize.W $cpuSize.H
$gpuCard = New-Card "GPU" ($row1LeftX + $cpuSize.W + $colSpace) $topY $gpuSize.W $gpuSize.H

# Row 2: RAM, Network, Storage in thirds
$row2Y   = $topY + $cardH + $rowSpace
$ramCard = New-Card "RAM" $row2LeftX $row2Y $ramSize.W $ramSize.H
$netCard = New-Card "Network" ($row2LeftX + $ramSize.W + $colSpace) $row2Y $netSize.W $netSize.H
$storCard= New-Card "Storage" ($row2LeftX + ($ramSize.W + $colSpace)*2) $row2Y $stoSize.W $stoSize.H

# Row 3: Top Processes full width - account for actual row 2 height
$row2Height = [math]::Max([math]::Max($ramSize.H, $netSize.H), $stoSize.H)
$row3Y   = $row2Y + $row2Height + $rowSpace
$procCard= New-Card "Top Processes" $row3LeftX $row3Y $topSize.W $topSize.H

# CPU widgets: left sparkline chart for Load%, right labels for Temp/Clock/Fan
# degree symbol to avoid encoding issues
$deg = [char]0x00B0

# CPU gauge (replace previous line chart)
$cpuGauge = New-Object System.Windows.Forms.Panel
$cpuGauge.Location = [System.Drawing.Point]::new(12,26)
$cpuGauge.Size     = [System.Drawing.Size]::new(110,110)
$cpuGauge.BackColor = [System.Drawing.Color]::Transparent
$cpuCard.Controls.Add($cpuGauge)|Out-Null

$cpuGauge.add_Paint({
  param($s,$e)
  $rect = [System.Drawing.Rectangle]::new(5,5,100,100)
  $basePen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray,8)
  $usePen  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromName([string]$cfg.ChartColor),8)
  $start = 270; $sweepTotal = 280
  $e.Graphics.SmoothingMode = 'AntiAlias'
  $e.Graphics.DrawArc($basePen,$rect,$start,$sweepTotal)
  if ($script:cpuLoadPct -ge 0) { $e.Graphics.DrawArc($usePen,$rect,$start,[float]($sweepTotal*[math]::Min(1,[math]::Max(0,$script:cpuLoadPct/100)))) }
  $txt = if ($script:cpuLoadPct -ge 0) { "{0:N0}%" -f $script:cpuLoadPct } else { "n/a" }
  $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment='Center'; $fmt.LineAlignment='Center'
  $e.Graphics.DrawString($txt,(New-Object System.Drawing.Font('Segoe UI Semibold',[float]([math]::Max(8,$defaultFontSize+2)))),(New-Object System.Drawing.SolidBrush($fontColor)),([System.Drawing.RectangleF]$rect),$fmt)
})

# Right-side labels next to gauge
$cpuTempLabel  = New-Object System.Windows.Forms.Label; $cpuTempLabel.AutoSize=$true; $cpuTempLabel.ForeColor=$fontColor; $cpuTempLabel.Location=[System.Drawing.Point]::new([int]($cpuGauge.Right + 12),40); $cpuCard.Controls.Add($cpuTempLabel)|Out-Null
$cpuTempBar    = New-Object System.Windows.Forms.Panel; $cpuTempBar.Size=[System.Drawing.Size]::new(120,8); $cpuTempBar.Location=[System.Drawing.Point]::new([int]($cpuGauge.Right + 12),28); $cpuCard.Controls.Add($cpuTempBar)|Out-Null
$cpuClockLabel = New-Object System.Windows.Forms.Label; $cpuClockLabel.AutoSize=$true; $cpuClockLabel.ForeColor=$fontColor; $cpuClockLabel.Location=[System.Drawing.Point]::new([int]($cpuGauge.Right + 12),70); $cpuCard.Controls.Add($cpuClockLabel)|Out-Null
$cpuFanL       = New-Object System.Windows.Forms.Label; $cpuFanL.AutoSize=$true;      $cpuFanL.ForeColor=$fontColor; $cpuFanL.Location=[System.Drawing.Point]::new([int]($cpuGauge.Right + 12),100); $cpuCard.Controls.Add($cpuFanL)|Out-Null

# GPU gauge (same as CPU)
$gpuGauge = New-Object System.Windows.Forms.Panel
$gpuGauge.Location = [System.Drawing.Point]::new(12,26)
$gpuGauge.Size     = [System.Drawing.Size]::new(110,110)
$gpuGauge.BackColor = [System.Drawing.Color]::Transparent
$gpuCard.Controls.Add($gpuGauge)|Out-Null

$gpuGauge.add_Paint({
  param($s,$e)
  $rect = [System.Drawing.Rectangle]::new(5,5,100,100)
  $basePen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray,8)
  $usePen  = New-Object System.Drawing.Pen($chartColor,8)
  $start = 270; $sweepTotal = 280
  $e.Graphics.SmoothingMode = 'AntiAlias'
  $e.Graphics.DrawArc($basePen,$rect,$start,$sweepTotal)
  if ($script:gpuLoadPct -ge 0) { $e.Graphics.DrawArc($usePen,$rect,$start,[float]($sweepTotal*[math]::Min(1,[math]::Max(0,$script:gpuLoadPct/100)))) }
  $txt = if ($script:gpuLoadPct -ge 0) { "{0:N0}%" -f $script:gpuLoadPct } else { "n/a" }
  $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment='Center'; $fmt.LineAlignment='Center'
  $fontColor2 = Resolve-Color ($cfg.Font.Color) ($fgColor)
  $e.Graphics.DrawString($txt,(New-Object System.Drawing.Font('Segoe UI Semibold',[float]([math]::Max(8,$defaultFontSize+2)))),(New-Object System.Drawing.SolidBrush($fontColor2)),([System.Drawing.RectangleF]$rect),$fmt)
})

$gpuTempLabel  = New-Object System.Windows.Forms.Label; $gpuTempLabel.AutoSize=$true; $gpuTempLabel.ForeColor=$fontColor; $gpuTempLabel.Location=[System.Drawing.Point]::new([int]($gpuGauge.Right + 12),40); $gpuCard.Controls.Add($gpuTempLabel)|Out-Null
$gpuTempBar    = New-Object System.Windows.Forms.Panel; $gpuTempBar.Size=[System.Drawing.Size]::new(120,8); $gpuTempBar.Location=[System.Drawing.Point]::new([int]($gpuGauge.Right + 12),28); $gpuCard.Controls.Add($gpuTempBar)|Out-Null
$gpuClockLabel = New-Object System.Windows.Forms.Label; $gpuClockLabel.AutoSize=$true; $gpuClockLabel.ForeColor=$fontColor; $gpuClockLabel.Location=[System.Drawing.Point]::new([int]($gpuGauge.Right + 12),70); $gpuCard.Controls.Add($gpuClockLabel)|Out-Null
$gpuFanL       = New-Object System.Windows.Forms.Label; $gpuFanL.AutoSize=$true;     $gpuFanL.ForeColor=$fontColor; $gpuFanL.Location=[System.Drawing.Point]::new([int]($gpuGauge.Right + 12),100); $gpuCard.Controls.Add($gpuFanL)|Out-Null
$gpuTempLabel.MinimumSize = [System.Drawing.Size]::new(220,0)
$gpuClockLabel.MinimumSize = [System.Drawing.Size]::new(220,0)
$gpuFanL.MinimumSize = [System.Drawing.Size]::new(220,0)

# RAM widgets (Load + Used/Total)
$ramStats = New-Object System.Windows.Forms.Label; $ramStats.AutoSize=$false; $ramStats.Width=[int]($ramCard.Width-24); $ramStats.TextAlign='MiddleCenter'; $ramStats.Location=[System.Drawing.Point]::new(12,146); $ramCard.Controls.Add($ramStats)|Out-Null
if ($false) { $null = 0 } # placeholder to keep structure; removed RAM top percentage label

# RAM gauge (semi-circular clock-style 6pm->4pm)
$ramGauge = New-Object System.Windows.Forms.Panel
$ramGauge.Location = [System.Drawing.Point]::new([int](($ramCard.Width-110)/2),30)
$ramGauge.Size     = [System.Drawing.Size]::new(110,110)
$ramGauge.BackColor = [System.Drawing.Color]::Transparent
$ramCard.Controls.Add($ramGauge)|Out-Null

$chartColorSource = if ($cfg -and $cfg.Chart -and $cfg.Chart.Color) { $cfg.Chart.Color } else { $cfg.ChartColor }
$chartColor = Resolve-Color $chartColorSource ([System.Drawing.Color]::FromName('DodgerBlue'))
$ramGauge.add_Paint({
  param($s,$e)
  $rect = [System.Drawing.Rectangle]::new(5,5,100,100)
  $basePen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray,8)
  $usePen  = New-Object System.Drawing.Pen($chartColor,8)
  # Rotate to start near 6 o'clock and end near 4 o'clock (~240 degree offset request)
  $start = 240
  $sweepTotal = 280
  $e.Graphics.SmoothingMode = 'AntiAlias'
  $e.Graphics.DrawArc($basePen,$rect,$start,$sweepTotal)
  # filled portion uses current RAM load percent (global updated later)
  if ($script:ramLoadPct -ge 0) {
    $e.Graphics.DrawArc($usePen,$rect,$start, [float]($sweepTotal * [math]::Min(1,[math]::Max(0,$script:ramLoadPct/100))))
  }
  # center text
  $txt = if ($script:ramLoadPct -ge 0) { "{0:N0}%" -f $script:ramLoadPct } else { "n/a" }
  $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment='Center'; $fmt.LineAlignment='Center'
  $fontColor = Resolve-Color ($cfg.Font.Color) ($fgColor)
  $brush = New-Object System.Drawing.SolidBrush($fontColor)
  $e.Graphics.DrawString($txt, (New-Object System.Drawing.Font('Segoe UI Semibold', [float]([math]::Max(8,$defaultFontSize+2)))), $brush, ([System.Drawing.RectangleF]$rect), $fmt)
})

# Network
$netTop = if($cfg.NetworkBox -and $cfg.NetworkBox.TopOffsetPx){ [int]$cfg.NetworkBox.TopOffsetPx } else { 40 }
$netGap = if($cfg.NetworkBox -and $cfg.NetworkBox.LineSpacingPx){ [int]$cfg.NetworkBox.LineSpacingPx } else { 44 }
$netTextWidth = if($cfg.NetworkBox -and $cfg.NetworkBox.TextWidthPx){ [int]$cfg.NetworkBox.TextWidthPx } else { [int]($netCard.Width-24) }
$netNow = New-Object System.Windows.Forms.Label; $netNow.Location=[System.Drawing.Point]::new(12,$netTop); $netNow.AutoSize=$false; $netNow.Width=$netTextWidth; $netNow.Height=30; $netNow.TextAlign='MiddleLeft'; $netCard.Controls.Add($netNow)|Out-Null
$net2   = New-Object System.Windows.Forms.Label; $net2.Location=[System.Drawing.Point]::new(12,($netTop+$netGap)); $net2.AutoSize=$false; $net2.Width=$netTextWidth; $net2.Height=30; $net2.TextAlign='MiddleLeft'; $netCard.Controls.Add($net2)|Out-Null
$arrowDown = [char]0x2193  # down arrow (ASCII-safe via char code)
$arrowUp   = [char]0x2191  # up arrow
# Scale network font
try {
  $nfs = if($cfg.NetworkBox -and $cfg.NetworkBox.FontSize){ [float]$cfg.NetworkBox.FontSize } else { [float]($form.Font.Size * 2.4) }
  $f = New-Object System.Drawing.Font($form.Font.FontFamily, $nfs, $form.Font.Style)
  $netNow.Font = $f; $net2.Font = $f
  $netNow.ForeColor = $fontColor; $net2.ForeColor = $fontColor
} catch {}

    # Storage meters - dynamically create for all available drives
    $script:storageMeters = @()
    $availableDrives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady -and $_.DriveType -eq 'Fixed' } | Sort-Object Name
    $yOffset = 40
    foreach($drive in $availableDrives) {
      $driveLetter = $drive.Name.TrimEnd('\')
      $meter = New-Meter $storCard $driveLetter $yOffset
      $script:storageMeters += @{ Drive=$driveLetter; Meter=$meter; YOffset=$yOffset }
      $yOffset += 70  # Increased spacing since text is now above bar
    }

# Top processes with CPU%, GPU%, RAM MB, Down/Up KB/s
$procList = New-Object System.Windows.Forms.ListView
$procList.View='Details'; $procList.FullRowSelect=$true; $procList.HeaderStyle='Nonclickable'
[void]$procList.Columns.Add("Name", 220)
[void]$procList.Columns.Add("CPU %", 70)
[void]$procList.Columns.Add("GPU %", 70)
[void]$procList.Columns.Add("RAM MB", 80)
[void]$procList.Columns.Add("Down KB/s", 90)
[void]$procList.Columns.Add("Up KB/s", 90)
$procList.Location=[System.Drawing.Point]::new(12,40)
$procList.Size=[System.Drawing.Size]::new([int]($procCard.Width-24),[int]($procCard.Height-60))
$procCard.Controls.Add($procList)|Out-Null
# Enable DoubleBuffering to reduce flicker
try { $flags=[System.Reflection.BindingFlags] "NonPublic,Instance"; $pi=$procList.GetType().GetProperty('DoubleBuffered',$flags); if($pi){ $pi.SetValue($procList,$true,$null) } } catch {}
$procList.Scrollable = $false

# Status
$status = New-Object System.Windows.Forms.Label
$status.Location=[System.Drawing.Point]::new(12, ($procCard.Bottom + 8))
$status.AutoSize=$true
$status.ForeColor=$fontColor
$form.Controls.Add($status) | Out-Null

# ===== TIMER: POLL + REFRESH =====
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int]$REFRESH_MS
$timer.Add_Tick({
  # Exit early if shutdown requested - AGGRESSIVE
  if ($script:shouldExit) { 
    try { $timer.Stop(); $timer.Dispose() } catch {}
    Kill-Now
    return
  }
  
  try{
    foreach($h in $computer.Hardware){ Update-HW $h }
    $roots = $computer.Hardware

    # ---- CPU (Chart + Temp/Clock/Fan) ----
    $cpu = $roots | Where-Object { $_.HardwareType.ToString() -eq "CPU" } | Select-Object -First 1
    if ($cpu) {
      $cpuLoadVal = (Get-Sensors $cpu @("Load") | Where-Object Name -match 'CPU Total' | Select-Object -ExpandProperty Value -First 1)
      if ($cpuLoadVal -ne $null) { 
        # Fix CPU percentage - sensor value is already a percentage, just cap at 100%
        $rawValue = [double]$cpuLoadVal
        $script:cpuLoadPct = [math]::Min(100, [math]::Max(0, $rawValue))
        # Debug: log the CPU value to verify
        Write-Log "CPU Load: raw=$rawValue, capped=$script:cpuLoadPct"
      } else { 
        $script:cpuLoadPct = -1 
      }
      $cpuGauge.Invalidate()

      $cpuTempVal = (Get-Sensors $cpu @("Temperature") | Select-Object -ExpandProperty Value -First 1)
      if ($cpuTempVal -ne $null) { $cpuTempLabel.Text = ("Temp: {0:N0}{1}C" -f $cpuTempVal, $deg); $script:cpuTempMarker=$cpuTempVal; $cpuTempBar.Invalidate() } else { $cpuTempLabel.Text = "Temp: n/a"; $script:cpuTempMarker=$null; $cpuTempBar.Invalidate() }

      $cpuClocks = (Get-Sensors $cpu @("Clock") | Where-Object Name -match 'core')
      $clkAvg = if ($cpuClocks){ ($cpuClocks | Measure-Object -Property Value -Average).Average } else { $null }
      if ($clkAvg -ne $null) { $cpuClockLabel.Text = ("Clock: {0:N0} MHz" -f $clkAvg) } else { $cpuClockLabel.Text = "Clock: n/a" }

      $mbRoots = $roots | Where-Object { $_.HardwareType.ToString() -in @("Mainboard","SuperIO","Cooler") }
      $fans = @(); foreach($m in $mbRoots){ $fans += Get-Sensors $m @("Fan") }
      $cpuFan = $null
      if ($fans){ $cpuFan = ($fans | Where-Object Name -match 'cpu|pump' | Select-Object -First 1) }
      if (-not $cpuFan -and $fans){ $cpuFan = $fans | Select-Object -First 1 }
      if ($cpuFan) { $cpuFanL.Text = ("Fan: {0:N0} RPM" -f ($cpuFan.Value -as [double])) } else { $cpuFanL.Text = "Fan: n/a" }
    } else {
      $cpuFanL.Text="CPU: no sensors"
      Add-CpuChartPoint 0; $cpuTempLabel.Text="Temp: n/a"; $cpuClockLabel.Text="Clock: n/a"
    }

    # ---- GPU (Temp + VRAM Used) ----
    $gpu = $roots | Where-Object { $_.HardwareType.ToString() -like "Gpu*" } | Select-Object -First 1
    if ($gpu) {
      $gLoad = (Get-Sensors $gpu @("Load") | Where-Object Name -match 'core|gpu' | Select-Object -ExpandProperty Value -First 1)
      if ($gLoad -ne $null){ $script:gpuLoadPct = [double]$gLoad } else { $script:gpuLoadPct = -1 }
      $gpuGauge.Invalidate()

      $gTemp = (Get-Sensors $gpu @("Temperature") | Select-Object -ExpandProperty Value -First 1)
      if ($gTemp -ne $null){ $gpuTempLabel.Text = ("Temp: {0:N0}{1}C" -f $gTemp, $deg); $script:gpuTempMarker=$gTemp; $gpuTempBar.Invalidate() } else { $gpuTempLabel.Text = "Temp: n/a"; $script:gpuTempMarker=$null; $gpuTempBar.Invalidate() }

      $gClock = (Get-Sensors $gpu @("Clock") | Where-Object Name -match 'core|graphics' | Select-Object -ExpandProperty Value -First 1)
      if ($gClock -ne $null){ $gpuClockLabel.Text = ("Clock: {0:N0} MHz" -f $gClock) } else { $gpuClockLabel.Text = "Clock: n/a" }

      $gf = Get-Sensors $gpu @("Fan")
      if ($gf -and $gf.Count -gt 0){
        $parts = @()
        foreach($f in ($gf | Select-Object -First 2)){ $parts += ("{0}: {1:N0} RPM" -f $f.Name, ($f.Value -as [double])) }
        $gpuFanL.Text = ($parts -join "   ")
      } else { $gpuFanL.Text = "Fan: n/a" }
    } else { $gpuFanL.Text="GPU: no sensors" }

    # ---- RAM ----
    $ram = $roots | Where-Object { $_.HardwareType.ToString() -eq "RAM" } | Select-Object -First 1
    if ($ram) {
      $rLoad = (Get-Sensors $ram @("Load") | Select-Object -ExpandProperty Value -First 1)
      if ($rLoad -ne $null){ $script:ramLoadPct=[double]$rLoad } else { $script:ramLoadPct=-1 }
      $ramGauge.Invalidate()

      $data = Get-Sensors $ram @("Data")
      $rUsed  = $null; $rAvail = $null
      if ($data){
        $rUsed  = ($data | Where-Object Name -match 'Used Memory'      | Select-Object -ExpandProperty Value -First 1)
        $rAvail = ($data | Where-Object Name -match 'Available Memory' | Select-Object -ExpandProperty Value -First 1)
      }
      if ($rUsed -ne $null -and $rAvail -ne $null){
        $rTot = $rUsed + $rAvail
        $ramStats.Text = ("Used: {0:N1} / {1:N1} GB" -f $rUsed,$rTot)
        $ramGauge.Invalidate()
      } else {
        $ramStats.Text = "Used: n/a"
      }
    } else {
      $ramStats.Text="RAM: no sensors"; $script:ramLoadPct=-1; $ramGauge.Invalidate()
    }

    # ---- Network (download/upload) ----
    try {
      $rx = (Get-Counter "\Network Interface(*)\Bytes Received/sec").CounterSamples
      $tx = (Get-Counter "\Network Interface(*)\Bytes Sent/sec").CounterSamples
      $down = if($rx){ ($rx | Measure-Object -Property CookedValue -Sum).Sum } else { 0 }
      $up   = if($tx){ ($tx | Measure-Object -Property CookedValue -Sum).Sum } else { 0 }
      function Add-SpaceUnit([string]$s){ return ($s -replace '(\d)([A-Za-z])','$1 $2') }
      $downText = Add-SpaceUnit (Format-Size $down)
      $upText   = Add-SpaceUnit (Format-Size $up)
      $netNow.Text = ("{0} {1}/s" -f $arrowDown, $downText)
      $net2.Text   = ("{0} {1}/s" -f $arrowUp,   $upText)
    } catch {
      Write-Log "Network update error" $_
      $netNow.Text = ("{0} n/a" -f $arrowDown); $net2.Text = ("{0} n/a" -f $arrowUp)
    }

    # ---- Storage (dynamic drive usage bars) ----
    foreach($storageItem in $script:storageMeters) {
      $driveLetter = $storageItem.Drive
      $meter = $storageItem.Meter
      
      $driveInfo = Get-DriveUsage "$driveLetter\"
      if ($driveInfo -ne $null) {

        $null = Set-Bar $meter ([double]$driveInfo.Percent) 100
        $free = [double]($driveInfo.Total - $driveInfo.Used)
        $meter.Value.Text = ("Avail: {0}   Used: {1}" -f (Format-Size $free), (Format-Size $driveInfo.Used))
        # Ensure no underlines in the text
        $meter.Value.Font = New-Object System.Drawing.Font($form.Font.FontFamily, $form.Font.Size, [System.Drawing.FontStyle]::Regular)
        # Set chart color for the bar
        $meter.Bar.ForeColor = Resolve-Color ($cfg.Chart.Color) ([System.Drawing.Color]::Blue)
      } else { 
        $meter.Bar.Value=0; 
        $meter.Value.Text="n/a"
        Write-Log "Storage $driveLetter drive not found or not ready"
      }
    }

    # ---- Top processes (CPU%, GPU%, RAM, per-process I/O read/write as Down/Up) ----
    $procList.BeginUpdate()
    try {
      $perf = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process -ErrorAction SilentlyContinue
      $gpuPerf = Get-CimInstance -ClassName Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine -ErrorAction SilentlyContinue
      $pidToGpu = @{}
      if ($gpuPerf) {
        foreach($e in $gpuPerf){
          $procId = $e.ProcessID
          if ($procId -gt 0) { if(-not $pidToGpu.ContainsKey($procId)){ $pidToGpu[$procId]=0 }; $pidToGpu[$procId] += [double]$e.UtilizationPercentage }
        }
      }

      $procs = Get-Process | Where-Object { $_.Id -ne 0 }
      $rows = @()
      foreach($p in $procs){
        $perfRow = $perf | Where-Object { $_.IDProcess -eq $p.Id } | Select-Object -First 1
        if (-not $perfRow) { continue }
        $cpuPct = [double]$perfRow.PercentProcessorTime
        if ($cpuPct -lt 0) { $cpuPct = 0 }
        $gpuPct = if ($pidToGpu.ContainsKey($p.Id)) { [double]$pidToGpu[$p.Id] } else { 0 }
        $downKB = [math]::Round(([double]$perfRow.IOReadBytesPersec)/1024,0)
        $upKB   = [math]::Round(([double]$perfRow.IOWriteBytesPersec)/1024,0)
        $ramMB  = [math]::Round($p.WorkingSet64/1MB,1)
        $rows += [pscustomobject]@{ Name=$p.ProcessName; CPU=$cpuPct; GPU=$gpuPct; RAM=$ramMB; Down=$downKB; Up=$upKB }
      }

      $topRows = $rows | Sort-Object CPU -Descending | Select-Object -First 6
      for($i=0;$i -lt $topRows.Count;$i++){
        if ($i -ge $procList.Items.Count) {
          $row = New-Object System.Windows.Forms.ListViewItem("")
          [void]$row.SubItems.Add("")
          [void]$row.SubItems.Add("")
          [void]$row.SubItems.Add("")
          [void]$row.SubItems.Add("")
          [void]$row.SubItems.Add("")
          [void]$procList.Items.Add($row)
        }
        $item = $procList.Items[$i]
        $r = $topRows[$i]
        $item.Text = $r.Name
        $item.SubItems[1].Text = ("{0:N1}" -f $r.CPU)
        $item.SubItems[2].Text = ("{0:N1}" -f $r.GPU)
        $item.SubItems[3].Text = ("{0:N1}" -f $r.RAM)
        $item.SubItems[4].Text = ("{0:N0}" -f $r.Down)
        $item.SubItems[5].Text = ("{0:N0}" -f $r.Up)
      }
      while($procList.Items.Count -gt $topRows.Count){ $procList.Items.RemoveAt($procList.Items.Count-1) }
    } catch {
      Write-Log "Top processes update error" $_
    }
    $procList.EndUpdate()

    # ---- Status ----
    $status.Text = "Last refresh: {0}" -f (Get-Date).ToString("HH:mm:ss")
  } catch {
    Write-Log "Timer tick error" $_
    $status.Text = "Last refresh: error at {0}" -f (Get-Date).ToString("HH:mm:ss")
  }
})

# ===== RUN =====
$timer.Start()



# Add process exit handler for force-quit
$null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
  $script:shouldExit = $true
  if ($timer) { $timer.Stop(); $timer.Dispose() }
  if ($computer) { try{ $computer.Close() }catch{} }
  if ($tray) { $tray.Dispose() }
}

# Remove duplicate - using the one above

[void]$form.ShowDialog()

# ===== CLEANUP =====
$script:shouldExit = $true
$timer.Stop()
$timer.Dispose()
try{ $computer.Close() }catch{}
$tray.Dispose()

# Force exit if we get here
if ($script:shouldExit) {
  [System.Environment]::Exit(0)
}
