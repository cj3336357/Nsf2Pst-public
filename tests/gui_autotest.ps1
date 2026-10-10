# Automated end-to-end test of the Nsf2Pst window: fills the fields, presses Start, waits, prints the log.
param([string]$Exe, [string]$Nsf, [string]$Pst, [switch]$Anonymize, [int]$TimeoutSec = 600, [string]$Screenshot, [string]$Directory)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$p = Start-Process $Exe -PassThru
$win = $null
for ($i = 0; $i -lt 60 -and -not $win; $i++) {
  Start-Sleep -Milliseconds 500
  $p.Refresh()
  if ($p.MainWindowHandle -ne 0) { $win = [Windows.Automation.AutomationElement]::FromHandle($p.MainWindowHandle) }
}
if (-not $win) { throw "window did not appear" }
function El($id) {
  $c = New-Object Windows.Automation.PropertyCondition ([Windows.Automation.AutomationElement]::AutomationIdProperty), $id
  $win.FindFirst([Windows.Automation.TreeScope]::Descendants, $c)
}
function SetText($id, $v) { (El $id).GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($v) }
SetText 'txtNsf' $Nsf
SetText 'txtPst' $Pst
if ($Directory) { SetText 'txtDir' $Directory }
if ($Anonymize) { (El 'chkAnon').GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern).Toggle() }
"window title: " + $win.Current.Name
(El 'btnStart').GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
$t0 = Get-Date
do {
  Start-Sleep -Seconds 1
  $status = (El 'lblStatus').Current.Name
} while ($status -notmatch '完成|失敗|出錯|已停止|缺少' -and ((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec)
"status: $status   ({0:N0}s)" -f ((Get-Date) - $t0).TotalSeconds
"----- log -----"
(El 'txtLog').GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.Value
if ($Screenshot) {
  Add-Type -AssemblyName System.Drawing
  $r = $win.Current.BoundingRectangle
  $bmp = New-Object Drawing.Bitmap ([int]$r.Width), ([int]$r.Height)
  $g = [Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size); $bmp.Save($Screenshot); $g.Dispose(); $bmp.Dispose()
}
$p.CloseMainWindow() | Out-Null
Start-Sleep -Seconds 1
if (-not $p.HasExited) { $p.Kill() }
