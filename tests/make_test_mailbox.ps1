# Builds a realistic test mail file from mail12.ntf for the NSF -> PST pipeline.
param([string]$Out = "$env:TEMP\nsf2pst-test\testmail.nsf")
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$s = New-Object -ComObject Lotus.NotesSession
$s.Initialize("")
$me = $s.UserName
$db = $s.GetDatabase("", "mail12.ntf").CreateFromTemplate("", $Out, $false)
Set-Content -Path (Join-Path (Split-Path $Out) "unread_noteids.txt") -Value $null

$att = Join-Path (Split-Path $Out) "附件 測試.txt"
[IO.File]::WriteAllText($att, "附件內容：灣仔辦事處 年会 嘅㗎啲 😀", (New-Object Text.UTF8Encoding $true))

function New-Memo($subject, $body, [string[]]$folders, [bool]$unread, [bool]$sent = $false, [string]$attach = $null, [datetime]$when) {
  $d = $db.CreateDocument()
  $null = $d.ReplaceItemValue("Form", "Memo")
  $null = $d.ReplaceItemValue("Subject", $subject)
  $null = $d.ReplaceItemValue("From", $(if ($sent) { $me } else { "CN=Alex Chan/OU=Sales/O=ExampleCorp" }))
  $null = $d.ReplaceItemValue("Principal", $d.GetItemValue("From"))
  $null = $d.ReplaceItemValue("SendTo", $(if ($sent) { "CN=Alex Chan/OU=Sales/O=ExampleCorp" } else { $me }))
  $null = $d.ReplaceItemValue("PostedDate", $s.CreateDateTime($when.ToString("yyyy/MM/dd HH:mm:ss")))
  if (-not $sent) { $null = $d.ReplaceItemValue("DeliveredDate", $s.CreateDateTime($when.AddMinutes(1).ToString("yyyy/MM/dd HH:mm:ss"))) }
  $rt = $d.CreateRichTextItem("Body")
  $rt.AppendText($body)
  if ($attach) { $rt.AddNewLine(1); $null = $rt.EmbedObject(1454, "", $attach) }
  $null = $d.Save($true, $false)
  foreach ($f in $folders) { $d.PutInFolder($f, $true) }
  if ($unread) { Add-Content -Path (Join-Path (Split-Path $Out) "unread_noteids.txt") -Value $d.NoteID }
  return $d
}

$base = Get-Date "2024-03-06 09:30"
$null = New-Memo "灣仔辦事處 工作安排"            "請準備灣仔辦事處下週嘅會議資料。"            @('($Inbox)')                       $true  -when $base
$null = New-Memo "年會 簡報 (2024)"               "年会期间嘅簡報同 第三季 資料，粵語：嘅㗎啲冇。" @('($Inbox)','research(including 年会,第三季）\年會') $false -when $base.AddDays(1)
$null = New-Memo "附件測試 九龍東辦事處"          "請查閱附件。"                                @('項目 (2024)')                    $true  -attach $att -when $base.AddDays(2)
$null = New-Memo "【通告】 年度預算"              "通告內容：年度預算 灣仔辦事處"              @('【通告】')                        $false -when $base.AddDays(3)
$null = New-Memo "Re: 灣仔辦事處 工作安排"        "已收到，謝謝。"                              @()                         $false -sent $true -when $base.AddDays(4)
$null = New-Memo "Unfiled 未分類信件 辦事處"      "呢封信唔喺任何 folder。"                     @()                                 $true  -when $base.AddDays(5)
$null = New-Memo "English only subject"          "Plain English body."                         @('項目 (2024)\Sub A')              $false -when $base.AddDays(6)

# a calendar entry (not a mail) - the exporter should report it as skipped
$c = $db.CreateDocument()
$null = $c.ReplaceItemValue("Form", "Appointment"); $null = $c.ReplaceItemValue("Subject", "會議 灣仔辦事處")
$null = $c.Save($true, $false)

"created: $($db.FilePath)  docs=$($db.AllDocuments.Count)  user=$me"
