# Test mail file with contacts, calendar entries and encrypted mail (for the NSF -> PST calendar / contacts path).
param([string]$Out = "$env:TEMP\nsf2pst-test\testpim.nsf")
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$s = New-Object -ComObject Lotus.NotesSession
$s.Initialize("")
$me = $s.UserName
if (Test-Path $Out) { $Out = $Out -replace '\.nsf$', ("_" + (Get-Date -Format HHmmss) + ".nsf") }
$db = $s.GetDatabase("", "mail12.ntf").CreateFromTemplate("", $Out, $false)
function DT([string]$t) { $s.CreateDateTime($t) }

# --- contacts
function New-Person($first, $last, $mail, $company, $title, $phone, $mobile, $street, $city, $comment) {
  $d = $db.CreateDocument()
  $null = $d.ReplaceItemValue("Form", "Person"); $null = $d.ReplaceItemValue("Type", "Person")
  $null = $d.ReplaceItemValue("FirstName", $first); $null = $d.ReplaceItemValue("LastName", $last)
  $null = $d.ReplaceItemValue("FullName", "CN=$first $last/O=Test")
  $null = $d.ReplaceItemValue("MailAddress", $mail); $null = $d.ReplaceItemValue("InternetAddress", $mail)
  $null = $d.ReplaceItemValue("CompanyName", $company); $null = $d.ReplaceItemValue("JobTitle", $title)
  $null = $d.ReplaceItemValue("OfficePhoneNumber", $phone); $null = $d.ReplaceItemValue("CellPhoneNumber", $mobile)
  $null = $d.ReplaceItemValue("OfficeStreetAddress", $street); $null = $d.ReplaceItemValue("OfficeCity", $city)
  $null = $d.ReplaceItemValue("OfficeCountry", "香港")
  $null = $d.ReplaceItemValue("Comment", $comment)
  $null = $d.ReplaceItemValue("Spouse", "配偶 測試")
  $null = $d.Save($true, $false)
}
New-Person "大文" "陳" "alex.chan@example.com" "灣仔辦事處" "高級主任" "+852 2123 4567" "+852 9123 4567" "彌敦道 100 號" "香港" "年會 聯絡人"
New-Person "Mary" "Wong" "mary.wong@example.org" "Example Ltd" "Manager" "2234 5678" "" "1 Queen's Road" "Central" ""
New-Person "小明" "李" "ming.lee@example.net" "九龍東辦事處" "" "" "6123 4567" "" "" "粵語：嘅㗎啲"

# --- calendar
function New-Cal($type, $subject, [string[]]$starts, [string[]]$ends, $location, $body, $alarm) {
  $d = $db.CreateDocument()
  $null = $d.ReplaceItemValue("Form", "Appointment"); $null = $d.ReplaceItemValue("AppointmentType", "$type")
  $null = $d.ReplaceItemValue("Subject", $subject)
  $null = $d.ReplaceItemValue("StartDateTime", [object[]]($starts | ForEach-Object { DT $_ }))
  $null = $d.ReplaceItemValue("EndDateTime", [object[]]($ends | ForEach-Object { DT $_ }))
  $null = $d.ReplaceItemValue("CalendarDateTime", [object[]]($starts | ForEach-Object { DT $_ }))
  if ($location) { $null = $d.ReplaceItemValue("Location", $location) }
  if ($type -eq 3) {
    $null = $d.ReplaceItemValue("Chair", $me)
    $null = $d.ReplaceItemValue("RequiredAttendees", [object[]]@("CN=Alex Chan/OU=Sales/O=ExampleCorp", "mary.wong@example.org"))
  }
  if ($alarm) { $null = $d.ReplaceItemValue("`$Alarm", 1); $null = $d.ReplaceItemValue("`$AlarmOffset", -$alarm) }
  $rt = $d.CreateRichTextItem("Body"); $rt.AppendText($body)
  $null = $d.Save($true, $false)
}
New-Cal 0 "約會：灣仔辦事處 工作" @("2024/03/11 10:00:00") @("2024/03/11 11:30:00") "會議室 A" "準備年會資料。" 15
New-Cal 2 "全日活動 年度預算" @("2024/03/12 00:00:00") @("2024/03/12 00:00:00") "" "全日" 0
New-Cal 3 "會議：九龍東辦事處" @("2024/03/13 14:00:00") @("2024/03/13 15:00:00") "Room 101" "議程：嘅㗎啲" 30
New-Cal 0 "每週例會 (重複)" @("2024/03/14 09:00:00", "2024/03/21 09:00:00", "2024/03/28 09:00:00") @("2024/03/14 09:30:00", "2024/03/21 09:30:00", "2024/03/28 09:30:00") "Teams" "每週一次" 10
New-Cal 4 "提醒：交報告" @("2024/03/15 17:00:00") @("2024/03/15 17:00:00") "" "" 0
New-Cal 1 "週年紀念" @("2024/03/16 00:00:00") @("2024/03/16 00:00:00") "" "" 0

# --- encrypted mail (encrypted for this ID -> must decrypt)
$m = $db.CreateDocument()
$null = $m.ReplaceItemValue("Form", "Memo"); $null = $m.ReplaceItemValue("Subject", "加密郵件 灣仔辦事處")
$null = $m.ReplaceItemValue("From", "CN=Alex Chan/OU=Sales/O=ExampleCorp"); $null = $m.ReplaceItemValue("SendTo", $me)
$null = $m.ReplaceItemValue("PostedDate", (DT "2024/03/10 09:00:00"))
$rt = $m.CreateRichTextItem("Body"); $rt.AppendText("機密內容：年會 預算")
$att = Join-Path (Split-Path $Out) "機密 附件.txt"; [IO.File]::WriteAllText($att, "機密附件內容 嘅㗎", (New-Object Text.UTF8Encoding $true))
$rt.AddNewLine(1); $null = $rt.EmbedObject(1454, "", $att)
$m.EncryptOnSend = $true
$null = $m.Encrypt()
$null = $m.Save($true, $false)
$m.PutInFolder('($Inbox)', $true)

"created: $($db.FilePath)  docs=$($db.AllDocuments.Count)  user=$me"
