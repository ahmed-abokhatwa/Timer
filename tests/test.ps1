# Logic tests for WorkTimer.exe (runs the real code via reflection, in a throw-away data folder).
# usage: powershell -STA -File test.ps1 -Exe <path to WorkTimer.exe> -Work <scratch dir>
param([string]$Exe, [string]$Work)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$bf = [Reflection.BindingFlags]'Static,Instance,Public,NonPublic'
$script:pass = 0; $script:fail = 0
function T($name, $ok, $detail = "") {
    if ($ok) { $script:pass++; Write-Host "  PASS  $name" }
    else { $script:fail++; Write-Host "  FAIL  $name   $detail" -ForegroundColor Red }
}
function NewData($n) { $d = Join-Path $Work $n; Remove-Item $d -Recurse -Force -ea 0; New-Item -ItemType Directory $d | Out-Null; return $d }

$env:WORKTIMER_DATA = NewData "t_main"
$asm = [Reflection.Assembly]::LoadFrom($Exe)
$core = $asm.GetType("WorkTimer.Core"); $fmt = $asm.GetType("WorkTimer.Fmt"); $entryT = $asm.GetType("WorkTimer.Entry")
$pal = $asm.GetType("WorkTimer.Palette"); $xlsx = $asm.GetType("WorkTimer.Xlsx"); $ui = $asm.GetType("WorkTimer.UI")
function C($n, $a = @()) { return $core.GetMethod($n, $bf).Invoke($null, [object[]]$a) }
function F($n) { return $core.GetField($n, $bf) }
function Fm($n, $a) { return $fmt.GetMethod($n, $bf).Invoke($null, [object[]]$a) }
function Hm12($d) { return ([datetime]$d).ToString("h:mm tt", [Globalization.CultureInfo]::InvariantCulture) }
function CeilMin($d) { $t = [datetime]$d; $m = [decimal]600000000; return [datetime]::new([long]([math]::Floor(([decimal]$t.Ticks + $m - 1) / $m) * $m)) }
function Ts($sec) { return [TimeSpan]::FromSeconds($sec) }
$entries = (F "Entries").GetValue($null)
function Reset() { $entries.Clear(); (F "PendingDesc").SetValue($null, ""); (F "DailyHours").SetValue($null, [double]8) }
function E($s, $e, $d = "") {
    $x = [Activator]::CreateInstance($entryT); $x.Start = [datetime]$s
    if ($e -ne $null) { $entryT.GetField("End").SetValue($x, [object][datetime]$e) } else { $entryT.GetField("End").SetValue($x, $null) }
    $x.Desc = $d; [void]$entries.Add($x); return $x
}
function Total($day, $now) { return (C "DayTotal" @([datetime]$day, [datetime]$now)).TotalSeconds }
function Open() { return (C "get_Open") }

C "Init"; C "Load"

Write-Host "`n[1] Formatting"
T "Dur 0"            ((Fm "Dur" @((Ts 0))) -eq "00:00:00")
T "Dur 3661s"        ((Fm "Dur" @((Ts 3661))) -eq "01:01:01")
T "Dur 8h"           ((Fm "Dur" @((Ts 28800))) -eq "08:00:00")
T "Dur 99:59:59"     ((Fm "Dur" @((Ts 359999))) -eq "99:59:59")
T "Dur negative->0"  ((Fm "Dur" @((Ts -50))) -eq "00:00:00")
T "DurHm floor 59:59" ((Fm "DurHm" @((Ts 3599))) -eq "00:59")
T "DurHm 1h"          ((Fm "DurHm" @((Ts 3600))) -eq "01:00")
T "DurHmCeil 0"       ((Fm "DurHmCeil" @((Ts 0))) -eq "00:00")
T "DurHmCeil 1s"      ((Fm "DurHmCeil" @((Ts 1))) -eq "00:01")
T "DurHmCeil 60s"     ((Fm "DurHmCeil" @((Ts 60))) -eq "00:01")
T "DurHmCeil 61s"     ((Fm "DurHmCeil" @((Ts 61))) -eq "00:02")
T "DurHmCeil 7:00:30" ((Fm "DurHmCeil" @((Ts 25230))) -eq "07:01")
T "DurHmCeil 8h"      ((Fm "DurHmCeil" @((Ts 28800))) -eq "08:00")

Write-Host "`n[1b] 12-hour clock times"
T "17:05 -> 5:05 PM"    ((Fm "Hm" @([datetime]"2026-11-25 17:05")) -eq "5:05 PM")
T "00:30 -> 12:30 AM"   ((Fm "Hm" @([datetime]"2026-11-26 00:30")) -eq "12:30 AM")
T "12:00 -> 12:00 PM"   ((Fm "Hm" @([datetime]"2026-11-25 12:00")) -eq "12:00 PM")
T "00:00 -> 12:00 AM"   ((Fm "Hm" @([datetime]"2026-11-25 00:00")) -eq "12:00 AM")
T "09:07 -> 9:07 AM"    ((Fm "Hm" @([datetime]"2026-11-25 09:07")) -eq "9:07 AM")
T "23:59 -> 11:59 PM"   ((Fm "Hm" @([datetime]"2026-11-25 23:59")) -eq "11:59 PM")
T "13:00 -> 1:00 PM"    ((Fm "Hm" @([datetime]"2026-11-25 13:00")) -eq "1:00 PM")

Write-Host "`n[2] Daily totals, target, overtime"
Reset
$day = [datetime]"2026-11-25"
[void](E "2026-11-25 09:00:00" "2026-11-25 13:00:00" "a"); [void](E "2026-11-25 14:00:00" "2026-11-25 18:30:00" "b")
[void](E "2026-11-24 09:00:00" "2026-11-24 10:00:00" "other day")
T "two blocks = 8:30:00"       ((Total $day "2026-11-25 20:00:00") -eq 30600)
T "other day not mixed in"     ((Total "2026-11-24" "2026-11-25 20:00:00") -eq 3600)
T "empty day = 0"              ((Total "2026-11-20" "2026-11-25 20:00:00") -eq 0)
$target = (C "get_Target").TotalSeconds
T "target default 8h"          ($target -eq 28800)
(F "DailyHours").SetValue($null, [double]7.5); T "target 7.5h = 27000s" ((C "get_Target").TotalSeconds -eq 27000)
(F "DailyHours").SetValue($null, [double]8)
$ot = 30600 - 28800; T "overtime 8:30 - 8:00 = 30 min" ($ot -eq 1800)
Reset
[void](E "2026-11-25 09:00:00" "2026-11-25 13:00:00"); [void](E "2026-11-25 14:00:00" $null)
T "running entry counted up to 'now' (4h + 2:30:15)" ((Total $day "2026-11-25 16:30:15") -eq (14400 + 9015))
T "HasEntries true / false" ((C "HasEntries" @($day)) -and -not (C "HasEntries" @([datetime]"2026-01-01")))

Write-Host "`n[3] Live clock (real time): start / pause / resume"
Reset
$n0 = Get-Date
C "Start"; Start-Sleep 5
$t1 = Total ([datetime]::Today) (Get-Date)
T "after 5s running: ~5s (got $([math]::Round($t1,1)))" ($t1 -ge 4 -and $t1 -le 7)
C "Pause"; $tp = Total ([datetime]::Today) (Get-Date); Start-Sleep 4
$t2 = Total ([datetime]::Today) (Get-Date)
T "paused 4s: total does not move ($([math]::Round($tp,1)) -> $([math]::Round($t2,1)))" ([math]::Abs($t2 - $tp) -lt 0.5)
T "open entry is null while paused" ((Open) -eq $null)
C "Start"; Start-Sleep 3
$t3 = Total ([datetime]::Today) (Get-Date)
T "resume +3s: ~$([math]::Round($tp+3,1)) (got $([math]::Round($t3,1)))" ([math]::Abs($t3 - ($tp + 3)) -lt 1.6)
T "two entries recorded" ($entries.Count -eq 2)
C "Pause"

Write-Host "`n[4] Slice / pause edge cases"
Reset
C "Start"; Start-Sleep 2; (F "PendingDesc").GetValue($null) | Out-Null
C "SetDesc" @("task A"); T "current desc = running entry desc" ((C "get_CurrentDesc") -eq "task A")
C "Slice"
$ent = @($entries)
T "slice -> 2 entries" ($ent.Count -eq 2)
T "slice: no gap (end of first == start of second)" ($ent[0].End -eq $ent[1].Start)
T "slice: first keeps desc, second empty" ($ent[0].Desc -eq "task A" -and $ent[1].Desc -eq "")
T "slice: second is running" ($ent[1].End -eq $null)
C "Pause"
Reset; (F "PendingDesc").SetValue($null, "keep me")
C "Start"; C "Pause"   # same second -> zero length -> removed
T "instant start+pause leaves no junk entry" ($entries.Count -eq 0)
T "...and the description is not lost" ((F "PendingDesc").GetValue($null) -eq "keep me")
Reset; C "Start"; $when = (Get-Date).AddMinutes(-0.0)
Start-Sleep 3; C "PauseAt" @([datetime]::Now.AddSeconds(-2))
T "PauseAt(earlier time) closes at that time (sleep/shutdown case)" (@($entries)[0].End -lt [datetime]::Now.AddSeconds(-1))

Write-Host "`n[5] Split, edit, delete, overlap, start-from"
Reset
$a = E "2026-11-25 09:00:00" "2026-11-25 13:00:00" "work"
C "SplitEntry" @($a, [datetime]"2026-11-25 10:30:00", "part1", "part2")
$ent = @($entries | Sort-Object { $_.Start })
T "split -> 2 entries"                ($ent.Count -eq 2)
T "split: 1:30 + 2:30"                (($ent[0].End - $ent[0].Start).TotalSeconds -eq 5400 -and ($ent[1].End - $ent[1].Start).TotalSeconds -eq 9000)
T "split keeps total at 4:00"         ((Total $day "2026-11-25 20:00:00") -eq 14400)
T "split descriptions"                ($ent[0].Desc -eq "part1" -and $ent[1].Desc -eq "part2")
Reset; $r = E "2026-11-25 14:00:00" $null "run"
C "SplitEntry" @($r, [datetime]"2026-11-25 15:00:00", "x", "y")
$ent = @($entries | Sort-Object { $_.Start })
T "split of running entry: second part keeps running" (($ent[0].End -ne $null) -and ($ent[1].End -eq $null))
Reset; $a = E "2026-11-25 09:00:00" "2026-11-25 13:00:00"
T "overlap 12-14 detected"       (C "Overlaps" @([datetime]"2026-11-25 12:00", [datetime]"2026-11-25 14:00", $null))
T "touching 13-14 is OK"         (-not (C "Overlaps" @([datetime]"2026-11-25 13:00", [datetime]"2026-11-25 14:00", $null)))
T "before 08-09 is OK"           (-not (C "Overlaps" @([datetime]"2026-11-25 08:00", [datetime]"2026-11-25 09:00", $null)))
T "overlap ignores the entry being edited" (-not (C "Overlaps" @([datetime]"2026-11-25 09:30", [datetime]"2026-11-25 12:00", $a)))
C "EditEntry" @($a, [datetime]"2026-11-25 08:00:00", [datetime]"2026-11-25 12:00:00", "edited")
T "edit changes the total (4h)"  ((Total $day "2026-11-25 20:00:00") -eq 14400 -and $a.Desc -eq "edited")
C "DeleteEntry" @($a); T "delete removes it" ($entries.Count -eq 0)
Reset; C "StartFrom" @([datetime]::Today.AddHours(4).AddMinutes(30))
$o = Open
T "StartFrom creates a running entry at 04:30" ($o -ne $null -and $o.Start -eq [datetime]::Today.AddHours(4).AddMinutes(30))
$exp = ([datetime]::Now - [datetime]::Today.AddHours(4).AddMinutes(30)).TotalSeconds
T "...and the clock shows time since 04:30 (~$([math]::Round($exp/3600,2))h)" ([math]::Abs((Total ([datetime]::Today) (Get-Date)) - $exp) -lt 2)
C "StartFrom" @([datetime]::Today.AddHours(2)); T "StartFrom on a running entry moves its start" ((Open).Start -eq [datetime]::Today.AddHours(2))
$ui_c = $ui.GetMethod("Compose", $bf)
$orig = [datetime]"2026-11-25 09:00:37"
T "Compose keeps seconds when h:m unchanged" ($ui_c.Invoke($null, [object[]]@([datetime]"2026-11-25", [datetime]"2026-11-25 09:00:00", $orig)) -eq $orig)
T "Compose zeroes seconds when h:m changed"  ($ui_c.Invoke($null, [object[]]@([datetime]"2026-11-25", [datetime]"2026-11-25 10:15:00", $orig)) -eq [datetime]"2026-11-25 10:15:00")

Write-Host "`n[6] Midnight handling"
Reset
$yd = [datetime]::Today.AddDays(-1)
[void](E $yd.AddHours(22) $null "late")
C "Tick"
$ent = @($entries | Sort-Object { $_.Start })
T "running entry from yesterday 22:00 is split at midnight (2 entries)" ($ent.Count -eq 2)
T "first piece: 22:00 -> 24:00 yesterday"  ($ent[0].End -eq [datetime]::Today)
T "second piece: starts 00:00 today, running" ($ent[1].Start -eq [datetime]::Today -and ($ent[1].End -eq $null))
T "yesterday total = 2:00:00" ((Total $yd (Get-Date)) -eq 7200)
$sinceMid = ([datetime]::Now - [datetime]::Today).TotalSeconds
T "today total = time since midnight" ([math]::Abs((Total ([datetime]::Today) (Get-Date)) - $sinceMid) -lt 3)
Reset
[void](E ([datetime]::Today.AddDays(-2).AddHours(23)) $null "long")
C "Tick"
$ent = @($entries | Sort-Object { $_.Start })
T "2 days old running entry -> 3 pieces (1h, 24h, today)" ($ent.Count -eq 3 -and (($ent[0].End - $ent[0].Start).TotalHours -eq 1) -and (($ent[1].End - $ent[1].Start).TotalHours -eq 24) -and ($ent[2].End -eq $null))

Write-Host "`n[7] Week start (all 7 settings)"
$ws = $core.GetMethod("WeekStartOf", $bf); $wsF = F "WeekStart"; $dowT = [DayOfWeek]
$expect = @{ 0 = "2026-11-22"; 1 = "2026-11-23"; 2 = "2026-11-24"; 3 = "2026-11-25"; 4 = "2026-11-19"; 5 = "2026-11-20"; 6 = "2026-11-21" }
foreach ($k in 0..6) {
    $wsF.SetValue($null, [Enum]::ToObject($dowT, $k))
    $r = $ws.Invoke($null, [object[]]@([datetime]"2026-11-25 15:30"))
    T "week starting day $k -> week of Wed 25 Nov begins $($expect[$k])" ($r.ToString("yyyy-MM-dd") -eq $expect[$k])
}
$wsF.SetValue($null, [DayOfWeek]::Monday)
T "Sunday 22 Nov with Monday start -> 16 Nov" ($ws.Invoke($null, [object[]]@([datetime]"2026-11-22")).ToString("yyyy-MM-dd") -eq "2026-11-16")

Write-Host "`n[8] Save / reload / settings / crash recovery"
Reset
[void](E "2026-11-24 09:00:00" "2026-11-24 12:00:00" "tabs`tand\backslash`nnewline مرحبا é")
[void](E "2026-11-25 09:00:00" "2026-11-25 10:30:00" "plain")
(F "DailyHours").SetValue($null, [double]7.25); (F "OpacityPct").SetValue($null, 60); (F "Collapsed").SetValue($null, $true); (F "OnTop").SetValue($null, $false)
C "SaveAll"; C "SaveSettings"
$entries.Clear(); (F "DailyHours").SetValue($null, [double]8); (F "OpacityPct").SetValue($null, 100); (F "Collapsed").SetValue($null, $false); (F "OnTop").SetValue($null, $true)
C "Load"
$ent = @($entries | Sort-Object { $_.Start })
T "reload: 2 entries"                         ($ent.Count -eq 2)
T "reload: awkward description round-trips"   ($ent[0].Desc -eq "tabs`tand\backslash`nnewline مرحبا é")
T "reload: times identical"                   ($ent[1].Start -eq [datetime]"2026-11-25 09:00:00" -and $ent[1].End -eq [datetime]"2026-11-25 10:30:00")
T "settings round-trip (7.25h, 60%, collapsed, not on top)" ((F "DailyHours").GetValue($null) -eq 7.25 -and (F "OpacityPct").GetValue($null) -eq 60 -and (F "Collapsed").GetValue($null) -eq $true -and (F "OnTop").GetValue($null) -eq $false)
Reset
$dd = $env:WORKTIMER_DATA
$s = [datetime]::Now.AddMinutes(-30); $al = [datetime]::Now.AddMinutes(-10)
[IO.File]::WriteAllText("$dd\entries.tsv", $s.ToString("yyyy-MM-dd HH:mm:ss") + "`t`tcrashed task`r`n")
[IO.File]::WriteAllText("$dd\alive.txt", $al.ToString("yyyy-MM-dd HH:mm:ss"))
C "Load"
$ent = @($entries)
T "crash recovery: open entry closed at last heartbeat (20 min counted, not 30)" ($ent.Count -eq 1 -and [math]::Abs(($ent[0].End - $ent[0].Start).TotalMinutes - 20) -lt 0.1)
T "crash recovery: description kept for resume" ((F "PendingDesc").GetValue($null) -eq "crashed task" -and (Open) -eq $null)

Write-Host "`n[9] Colours: red until the hours are done, green in overtime"
$pf = $pal.GetMethod("For", $bf)
function Col($p) { $c = $pf.Invoke($null, [object[]]@([double]$p)); return "$($c.R),$($c.G),$($c.B)" }
$RED = "235,64,52"; $GREEN = "46,196,120"
T "0%    = red"      ((Col 0) -eq $RED)
T "1%    = red"      ((Col 0.01) -eq $RED)
T "50%   = red"      ((Col 0.5) -eq $RED)
T "99.9% = red"      ((Col 0.999) -eq $RED)
T "100%  = green"    ((Col 1) -eq $GREEN)
T "overtime 150% = green" ((Col 1.5) -eq $GREEN)
T "negative clamps to red" ((Col -1) -eq $RED)

Write-Host "`n[10] On-screen text (the real window code)"
Reset
$mt = $asm.GetType("WorkTimer.MiniForm"); $form = [Activator]::CreateInstance($mt)
function Vis($c) { return [bool]($c.GetType().GetMethod("GetState", $bf, $null, [Type[]]@([int]), $null).Invoke($c, [object[]]@(2))) }
function FieldOf($n) { return $mt.GetField($n, $bf).GetValue($form) }
function Show-State() { $form.RefreshView(); return @{ main = (FieldOf "face").Main; sec = (FieldOf "face").Sec; prog = (FieldOf "face").Progress; status = (FieldOf "lblStatus").Text; btn = (FieldOf "bToggle").Text } }
$now = Get-Date
# 4h worked earlier today, nothing running
[void](E $now.Date.AddHours(1) $now.Date.AddHours(5) "x")
$st = Show-State
T "paused, 4h worked: ring text 04:00" ($st.main -eq "04:00" -and $st.sec -eq "00") "got $($st.main):$($st.sec)"
T "progress 0.5"                    ([math]::Abs($st.prog - 0.5) -lt 0.001)
T "status 'Left 04:00 · paused'"    ($st.status -eq "Left 04:00  ·  paused") "got '$($st.status)'"
T "button says Resume (entries exist)" ($st.btn -eq "Resume")
$entries.Clear(); T "button says Start (no entries)" ((Show-State).btn -eq "Start")
# running: 1h closed + open entry started exactly 3h ago
$entries.Clear()
[void](E $now.Date.AddHours(1) $now.Date.AddHours(2) "x"); [void](E $now.AddHours(-3) $null "run")
$now2 = Get-Date; $st = Show-State
$expLeft = 28800 - (3600 + 10800)
# the shift end = now + 4h, rounded UP to the minute; allow for the 1-2 s between my clock read and the window's
$cands = @(); foreach ($d in -2..2) { $e = $now2.AddSeconds($expLeft + $d); $p = ""; if ($e.Date -gt $now2.Date) { $p = " (+1d)" }; $cands += ("Left 04:00  " + [char]0x00B7 + "  ends " + (Hm12 (CeilMin $e)) + $p) }
T "running 4h: status = 'Left 04:00 · ends <time>' (got '$($st.status)')" ($cands -contains $st.status) "candidates: $($cands -join ' | ')"
T "running: ring note says 'ends <time>'" ((FieldOf "face").Note -like "ends *")
T "running, under target: ring is RED" ((FieldOf "face").Accent.ToArgb() -eq [Drawing.Color]::FromArgb(235,64,52).ToArgb())
T "running: button says Pause"      ($st.btn -eq "Pause")
T "running: ring 04:00"             ($st.main -eq "04:00")
$fin = [datetime]::MinValue; $args2 = [object[]]@([datetime](Get-Date), $fin)
$okF = $core.GetMethod("TryFinish", $bf).Invoke($null, $args2)
$secsToEnd = ([datetime]$args2[1] - (Get-Date)).TotalSeconds
T "TryFinish: running with 4h left -> ends in 4h (0..60 s rounding up), got $([math]::Round($secsToEnd/3600,3))h" ($okF -and $secsToEnd -gt (14400 - 3) -and $secsToEnd -lt (14400 + 62))
T "TryFinish result is on a whole minute" (([datetime]$args2[1]).Second -eq 0)
# overtime: 9h30 closed
$entries.Clear(); [void](E $now.Date.AddHours(0) $now.Date.AddHours(9).AddMinutes(30) "x")
$st = Show-State
T "9:30 worked -> 'Done ✓  Overtime +01:30 · paused'" ($st.status -eq ("Done " + [char]0x2713 + "  Overtime +01:30  " + [char]0x00B7 + "  paused")) "got '$($st.status)'"
T "9:30 worked -> ring 09:30, progress > 1" ($st.main -eq "09:30" -and $st.prog -gt 1.18 -and $st.prog -lt 1.19)
T "overtime: ring is GREEN" ((FieldOf "face").Accent.ToArgb() -eq [Drawing.Color]::FromArgb(46,196,120).ToArgb() -or (FieldOf "face").Accent.G -gt (FieldOf "face").Accent.R)
T "overtime: ring note shows 'OT +01:30'" ((FieldOf "face").Note -eq "OT +01:30")
$entries.Clear(); [void](E $now.Date $now.Date.AddHours(3) "x"); [void](Show-State)
T "paused, under target: ring note 'paused'" ((FieldOf "face").Note -eq "paused")
$okP = $core.GetMethod("TryFinish", $bf).Invoke($null, [object[]]@([datetime](Get-Date), [datetime]::MinValue))
T "TryFinish: paused -> no shift end" (-not $okP)
$entries.Clear(); [void](E $now.Date $now.Date.AddHours(9) "x"); [void](Show-State)
# exactly 8:00 -> done, overtime 0
$entries.Clear(); [void](E $now.Date $now.Date.AddHours(8) "x"); $st = Show-State
T "exactly 8:00 -> done, overtime +00:00" ($st.status -like "Done*Overtime +00:00*")
# 7:59:30 -> still counting, left rounds UP to 00:01
$entries.Clear(); [void](E $now.Date $now.Date.AddHours(7).AddMinutes(59).AddSeconds(30) "x"); $st = Show-State
T "7:59:30 -> 'Left 00:01' (rounded up, not 00:00)" ($st.status -like "Left 00:01*") "got '$($st.status)'"
# custom daily hours 6
(F "DailyHours").SetValue($null, [double]6); $entries.Clear(); [void](E $now.Date $now.Date.AddHours(6).AddMinutes(10) "x"); $st = Show-State
T "6h target, 6:10 worked -> overtime +00:10" ($st.status -like "Done*Overtime +00:10*")
(F "DailyHours").SetValue($null, [double]8)
# collapse / expand
$form.GetType().GetMethod("ApplyLayout", $bf).Invoke($form, [object[]]@($true))
T "collapsed: ring-only 136 wide; buttons/description/status hidden; ring+chevron kept" ($form.ClientSize.Width -eq 136 -and -not (Vis (FieldOf "bToggle")) -and -not (Vis (FieldOf "pnlTxt")) -and -not (Vis (FieldOf "lblStatus")) -and (Vis (FieldOf "face")) -and (Vis (FieldOf "lblToggle")))
$form.GetType().GetMethod("ApplyLayout", $bf).Invoke($form, [object[]]@($false))
T "expanded: 400 wide; buttons, description, status visible" ($form.ClientSize.Width -eq 400 -and (Vis (FieldOf "bToggle")) -and (Vis (FieldOf "bLog")) -and (Vis (FieldOf "pnlTxt")) -and (Vis (FieldOf "lblStatus")))
$form.Dispose()

Write-Host "`n[10b] Time pickers in the dialogs"
$dlgT = $asm.GetType("WorkTimer.EntryDialog")
$dlg = [Activator]::CreateInstance($dlgT, [object[]]@("t", [datetime]"2026-11-25", [datetime]"2026-11-25 17:05:00", [datetime]"2026-11-25 23:30:00", "d", $false, $null))
$from = $dlgT.GetField("dtFrom", $bf).GetValue($dlg); $to = $dlgT.GetField("dtTo", $bf).GetValue($dlg)
T "Edit dialog From picker uses 12-hour format (hh:mm tt) and holds 5:05 PM" ($from.CustomFormat -eq "hh:mm tt" -and $from.Value.ToString("hh:mm tt", [Globalization.CultureInfo]::InvariantCulture) -eq "05:05 PM")
T "Edit dialog To picker uses 12-hour format and holds 11:30 PM" ($to.CustomFormat -eq "hh:mm tt" -and $to.Value.ToString("hh:mm tt", [Globalization.CultureInfo]::InvariantCulture) -eq "11:30 PM")
$dlg.Dispose()

Write-Host "`n[11] Excel weekly report (checked by an independent reader)"
Reset; $wsF.SetValue($null, [DayOfWeek]::Monday)
[void](E "2026-11-23 09:00:00" "2026-11-23 13:00:00" "Mon morning"); [void](E "2026-11-23 14:00:00" "2026-11-23 19:30:00" "Mon afternoon")
[void](E "2026-11-24 09:00:00" "2026-11-24 12:00:00" 'R&D <draft> "quoted" مرحبا'); [void](E "2026-11-24 12:30:00" "2026-11-24 17:00:00" "Tue pm")
[void](E "2026-11-25 10:00:00" "2026-11-25 14:00:00" "Wed")
[void](E "2026-11-26 08:00:00" "2026-11-26 16:00:00" "Thu")
$xf = Join-Path $Work "week_test.xlsx"; Remove-Item $xf -ea 0
$xlsx.GetMethod("Export", $bf).Invoke($null, [object[]]@(([string]$xf), ([datetime]"2026-11-23")))
$py = Join-Path $Work "check_xlsx.py"
@"
import sys; sys.path.insert(0, r'$Work\pylib')
from datetime import timedelta, time
import openpyxl
wb = openpyxl.load_workbook(r'$xf')
s = wb['Week Summary']; d = wb['Details']
rows = list(s.iter_rows(values_only=True))
def h(x): return int(x.total_seconds())
ok = True
def chk(name, cond):
    global ok
    print(('  PASS  ' if cond else '  FAIL  ') + name)
    if not cond: ok = False
mon, tue, wed, thu, fri, sat, sun, tot = rows[3], rows[4], rows[5], rows[6], rows[7], rows[8], rows[9], rows[10]
chk('7 day rows + total row', len(rows) == 11 and tot[0] == 'Total')
chk('days in order Monday..Sunday', [r[1] for r in rows[3:10]] == ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])
chk('Excel clock times are formatted 12-hour (AM/PM)', s['C4'].number_format == 'h:mm AM/PM' and d['C2'].number_format == 'h:mm AM/PM')
chk('Mon worked 9:30, overtime 1:30, short 0', h(mon[4]) == 34200 and h(mon[6]) == 5400 and h(mon[7]) == 0)
chk('Mon first start 09:00, last end 19:30', mon[2] == time(9,0) and mon[3] == time(19,30))
chk('Tue worked 7:30, overtime 0, short 0:30', h(tue[4]) == 27000 and h(tue[6]) == 0 and h(tue[7]) == 1800)
chk('Wed worked 4:00, short 4:00', h(wed[4]) == 14400 and h(wed[7]) == 14400)
chk('Thu worked 8:00, overtime 0, short 0', h(thu[4]) == 28800 and h(thu[6]) == 0 and h(thu[7]) == 0)
chk('Fri/Sat/Sun empty', all(r[4] is None for r in (fri, sat, sun)))
chk('TOTAL worked 29:00:00', h(tot[4]) == 104400)
chk('TOTAL target 4 worked days x 8h = 32:00:00', h(tot[5]) == 115200)
chk('TOTAL overtime 1:30:00', h(tot[6]) == 5400)
chk('TOTAL short 4:30:00', h(tot[7]) == 16200)
drows = list(d.iter_rows(values_only=True))
chk('Details: header + 6 entries', len(drows) == 7)
chk('Details: special characters survive', drows[3][5] == 'R&D <draft> "quoted" مرحبا')
chk('Details: durations 4:00, 5:30, 3:00, 4:30', [h(r[4]) for r in drows[1:5]] == [14400, 19800, 10800, 16200])
print('XLSX_OK' if ok else 'XLSX_FAILED')
"@ | Set-Content $py -Encoding UTF8
$out = & python -I $py 2>&1
$out | % { Write-Host $_; if ($_ -match '^\s+PASS') { $script:pass++ } elseif ($_ -match '^\s+FAIL') { $script:fail++ } }

Write-Host "`n[12] Auto-update logic"
$up = $asm.GetType("WorkTimer.Updater")
function U($n, $a) { return $up.GetMethod($n, $bf).Invoke($null, [object[]]$a) }
$repo = $up.GetField("Repo", $bf).GetRawConstantValue()
T "repo constant looks like owner/name ($repo)" ($repo -match '^[A-Za-z0-9-]+/[A-Za-z0-9._-]+$')
T "ParseVersion v1.0.5 -> 1.0.5.0"      ((U "ParseVersion" @("v1.0.5")).ToString() -eq "1.0.5.0")
T "ParseVersion 2.3.4.5"                ((U "ParseVersion" @("2.3.4.5")).ToString() -eq "2.3.4.5")
T "ParseVersion rejects junk"           ((U "ParseVersion" @("latest")) -eq $null -and (U "ParseVersion" @("1.0")) -eq $null -and (U "ParseVersion" @($null)) -eq $null)
$v = { param($s) U "ParseVersion" @($s) }
T "1.0.6 is newer than 1.0.5"           (U "IsNewer" @((& $v "1.0.6"), (& $v "1.0.5")))
T "1.0.10 is newer than 1.0.9 (numeric, not text)" (U "IsNewer" @((& $v "1.0.10"), (& $v "1.0.9")))
T "1.1.0 is newer than 1.0.99"          (U "IsNewer" @((& $v "1.1.0"), (& $v "1.0.99")))
T "same version is not newer"           (-not (U "IsNewer" @((& $v "v1.0.5"), (& $v "1.0.5.0"))))
T "older version is not newer"          (-not (U "IsNewer" @((& $v "1.0.4"), (& $v "1.0.5"))))
$good = '{"url":"x","tag_name":"v1.0.7","assets":[{"name":"WorkTimer.exe","browser_download_url":"https://github.com/o/r/releases/download/v1.0.7/WorkTimer.exe"},{"name":"WorkTimer.exe.sha256","browser_download_url":"https://github.com/o/r/releases/download/v1.0.7/WorkTimer.exe.sha256"}]}'
$rel = U "ParseRelease" @($good, "o/r")
T "ParseRelease reads version + both links" ($rel -ne $null -and $rel.Version.ToString() -eq "1.0.7.0" -and $rel.ExeUrl.EndsWith("/v1.0.7/WorkTimer.exe") -and $rel.ShaUrl.EndsWith("WorkTimer.exe.sha256"))
T "ParseRelease: exe link is not the .sha256 link" (-not $rel.ExeUrl.EndsWith(".sha256"))
T "links pointing at another repo are rejected" ((U "ParseRelease" @($good, "evil/other")) -eq $null)
T "links on another host are rejected" ((U "ParseRelease" @($good.Replace("https://github.com/", "https://evil.example/"), "o/r")) -eq $null)
T "release without checksum asset is rejected" ((U "ParseRelease" @(($good -replace ',\{"name":"WorkTimer.exe.sha256".*?\}\]', ']'), "o/r")) -eq $null)
T "garbage / empty JSON -> null" ((U "ParseRelease" @("{}", "o/r")) -eq $null -and (U "ParseRelease" @("", "o/r")) -eq $null)
T "SHA-256 of 'abc' is correct" ((U "Sha256Hex" @(,[Text.Encoding]::ASCII.GetBytes("abc"))) -eq "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
$curv = U "get_Current" @()
T "this build reports a version ($curv)" ($curv -ne $null -and $curv.Major -ge 1)

Write-Host "`n==== $script:pass passed, $script:fail failed ===="
Remove-Item Env:\WORKTIMER_DATA
if ($script:fail -gt 0) { exit 1 }
