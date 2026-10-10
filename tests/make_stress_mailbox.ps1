# Synthetic large mail file for stress testing: N memos, mixed Traditional/Simplified/Cantonese/English text,
# nested folders (with brackets / CJK names), ~10% with attachments, ~20% unread, ~10% in two folders.
param(
  [int]$Count = 1000,
  [string]$Out = "$env:TEMP\nsf2pst-test\stress.nsf"
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$rand = New-Object System.Random 42
$s = New-Object -ComObject Lotus.NotesSession; $s.Initialize("")
$me = $s.UserName
$db = $s.GetDatabase("", "mail12.ntf").CreateFromTemplate("", $Out, $false)
$dir = Split-Path $Out
$unreadFile = Join-Path $dir "stress_unread.txt"; Set-Content $unreadFile $null

$folders = @('($Inbox)', '2019', '2019\HKEx Information', '2021\沈阳论坛（延期）', '2022\宴请申请sample', '項目 (2024)', '項目 (2024)\Sub A',
             '【通告】', 'research(including 年会,第三季）\年會', '(Old) mails', 'Budget （funding ）', 'partners (APAC)', 'Manager''s notes')
$texts = @('灣仔辦事處 工作安排 九龍東辦事處', '年会 第三季 简体中文 湾仔办事处', '粵語 嘅 㗎 啲 冇 乜嘢', 'Quarterly report and budget review',
           '年度預算 通告內容', '会议纪要 出差申请 宴请', 'Mixed 混合 繁體 简体 emoji 😀')
# a 100 KB attachment
$att = Join-Path $dir "stress_附件.bin"; [IO.File]::WriteAllBytes($att, (1..102400 | ForEach-Object { [byte]($_ % 251) }))

$start = Get-Date "2019-01-01"
for ($i = 1; $i -le $Count; $i++) {
  $t = $texts[$rand.Next($texts.Count)]
  $d = $db.CreateDocument()
  $null = $d.ReplaceItemValue("Form", "Memo")
  $null = $d.ReplaceItemValue("Subject", "#$i $t")
  $null = $d.ReplaceItemValue("From", "CN=Alex Chan/OU=Sales/O=ExampleCorp")
  $null = $d.ReplaceItemValue("SendTo", $me)
  $when = $start.AddHours($i * 7)
  $null = $d.ReplaceItemValue("PostedDate", $s.CreateDateTime($when.ToString("yyyy/MM/dd HH:mm:ss")))
  $null = $d.ReplaceItemValue("DeliveredDate", $s.CreateDateTime($when.AddMinutes(1).ToString("yyyy/MM/dd HH:mm:ss")))
  $rt = $d.CreateRichTextItem("Body")
  $rt.AppendText("第 $i 封：$t`n" + ($texts[$rand.Next($texts.Count)] + "。") * (1 + $rand.Next(20)))
  if ($rand.NextDouble() -lt 0.10) { $rt.AddNewLine(1); $null = $rt.EmbedObject(1454, "", $att) }
  $null = $d.Save($true, $false)
  $f1 = $folders[$rand.Next($folders.Count)]; $d.PutInFolder($f1, $true)
  if ($rand.NextDouble() -lt 0.10) { $d.PutInFolder($folders[$rand.Next($folders.Count)], $true) }
  if ($rand.NextDouble() -lt 0.20) { Add-Content $unreadFile $d.NoteID }
  if ($i % 200 -eq 0) { Write-Host "  $i / $Count" }
}
"created: $($db.FilePath)  docs=$($db.AllDocuments.Count)  size=$((Get-Item $Out).Length)"
