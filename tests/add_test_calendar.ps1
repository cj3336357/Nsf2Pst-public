# Adds test calendar entries (as the Notes client creates them) to a COPY of a mail file.
param([string]$Src, [string]$Out)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
Copy-Item $Src $Out
$s = New-Object -ComObject Lotus.NotesSession
$s.Initialize("")
$db = $s.GetDatabase("", $Out)
$owner = [string]($db.GetProfileDocument("CalendarProfile").GetItemValue("Owner") | Select-Object -First 1)
if (-not $owner) { $owner = $s.UserName }
function DT([string]$t) { $s.CreateDateTime($t) }

function New-Cal($type, $subject, [string[]]$starts, [string[]]$ends, $location, $body, $alarm) {
  $d = $db.CreateDocument()
  $null = $d.ReplaceItemValue("Form", "Appointment")
  $null = $d.ReplaceItemValue("AppointmentType", "$type")
  $null = $d.ReplaceItemValue("Subject", $subject)
  $sd = [object[]]($starts | % { DT $_ }); $ed = [object[]]($ends | % { DT $_ })
  foreach ($i in "StartDateTime", "CalendarDateTime", "StartDate", "StartTime", "RepeatInstanceDates") { $null = $d.ReplaceItemValue($i, $sd) }
  foreach ($i in "EndDateTime", "EndDate", "EndTime") { $null = $d.ReplaceItemValue($i, $ed) }
  if ($starts.Count -gt 1) { $null = $d.ReplaceItemValue("Repeats", "1") }
  $null = $d.ReplaceItemValue("Chair", $owner); $null = $d.ReplaceItemValue("Principal", $owner); $null = $d.ReplaceItemValue("From", $owner)
  $null = $d.ReplaceItemValue("`$BusyName", $owner); $null = $d.ReplaceItemValue("`$BusyPriority", "1")
  $null = $d.ReplaceItemValue("ExcludeFromView", [object[]]@("D", "S"))
  $null = $d.ReplaceItemValue("`$PublicAccess", "1")
  $null = $d.ReplaceItemValue("OrgTable", "C0")
  if ($location) { $null = $d.ReplaceItemValue("Location", $location) }
  if ($type -eq 3) { $null = $d.ReplaceItemValue("RequiredAttendees", [object[]]@($owner, "CN=Alex Chan/OU=Sales/O=ExampleCorp")) }
  if ($alarm) { $null = $d.ReplaceItemValue("`$Alarm", 1); $null = $d.ReplaceItemValue("`$AlarmOffset", -$alarm) }
  $rt = $d.CreateRichTextItem("Body"); $rt.AppendText($body)
  $d.ComputeWithForm($false, $false) | Out-Null
  $null = $d.Save($true, $false)
}
New-Cal 0 "測試約會：灣仔辦事處 工作會議" @("2025/02/10 10:00:00") @("2025/02/10 11:30:00") "會議室 A" "測試用約會，準備年會資料。" 15
New-Cal 2 "測試全日活動：年度預算" @("2025/02/12 00:00:00") @("2025/02/12 00:00:00") "" "全日活動" 0
New-Cal 3 "測試會議：九龍東辦事處" @("2025/02/13 14:00:00") @("2025/02/13 15:00:00") "Room 101" "議程：嘅㗎啲" 30
New-Cal 0 "測試每週例會 (重複)" @("2025/02/17 09:00:00", "2025/02/24 09:00:00", "2025/03/03 09:00:00") @("2025/02/17 09:30:00", "2025/02/24 09:30:00", "2025/03/03 09:30:00") "Teams" "每週一次" 10

$v = $db.GetView('($Calendar)'); $v.Refresh()
"owner: $owner | docs now $($db.AllDocuments.Count) | (`$Calendar) view entries: $($v.AllEntries.Count)"
