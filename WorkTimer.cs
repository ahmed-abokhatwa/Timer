// Work Timer - small tray time tracker: 8h target, overtime, slices with descriptions, weekly Excel export.
// Targets C# 5 / .NET Framework 4.x (built into Windows), so the EXE needs no installs.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("Work Timer")]
[assembly: System.Reflection.AssemblyProduct("Work Timer")]

namespace WorkTimer
{
    // ------------------------------------------------------------------ data
    class Entry
    {
        public DateTime Start;
        public DateTime? End;          // null = still running
        public string Desc = "";
        public TimeSpan Dur(DateTime now) { return (End.HasValue ? End.Value : now) - Start; }
    }

    static class Fmt
    {
        public static string Dur(TimeSpan t)
        {
            if (t < TimeSpan.Zero) t = TimeSpan.Zero;
            long s = (long)Math.Round(t.TotalSeconds);
            return (s / 3600).ToString("00") + ":" + ((s / 60) % 60).ToString("00") + ":" + (s % 60).ToString("00");
        }
        public static string DurHm(TimeSpan t)
        {
            if (t < TimeSpan.Zero) t = TimeSpan.Zero;
            long m = (long)Math.Floor(t.TotalMinutes);
            return (m / 60).ToString("00") + ":" + (m % 60).ToString("00");
        }
        // minutes rounded UP, so "time left" never shows 00:00 while there is still time to go
        public static string DurHmCeil(TimeSpan t)
        {
            if (t < TimeSpan.Zero) t = TimeSpan.Zero;
            long m = (long)Math.Ceiling(Math.Round(t.TotalSeconds) / 60.0);
            return (m / 60).ToString("00") + ":" + (m % 60).ToString("00");
        }
        // clock times are shown in 12-hour format, e.g. "5:30 PM", "12:30 AM"
        public static string Hm(DateTime d) { return d.ToString("h:mm tt", CultureInfo.InvariantCulture); }
    }

    static class Core
    {
        public static List<Entry> Entries = new List<Entry>();
        public static string PendingDesc = "";
        public static double DailyHours = 8;
        public static DayOfWeek WeekStart = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
        public static bool AutoStart = true;
        public static int OpacityPct = 100;       // 25..100, window opacity
        public static bool OnTop = true;
        public static bool Collapsed = false;     // true = only the ring clock is shown
        public static bool AutoUpdate = true;     // check GitHub for new versions and install them
        public static int WinX = int.MinValue, WinY = int.MinValue;
        public static DateTime? RecoveredAt;      // set when the app found a running timer from a previous session
        public static bool FirstRun;
        public static string Dir;
        public static event Action Changed;

        static string entriesFile, settingsFile, aliveFile, errorFile;
        static bool dirty, settingsDirty;
        static DateTime lastAlive = DateTime.MinValue;
        const string DtFmt = "yyyy-MM-dd HH:mm:ss";

        public static TimeSpan Target { get { return TimeSpan.FromHours(DailyHours); } }

        public static DateTime Now()
        {
            DateTime n = DateTime.Now;
            return new DateTime(n.Ticks - n.Ticks % TimeSpan.TicksPerSecond, n.Kind);
        }

        public static Entry Open
        {
            get { foreach (Entry e in Entries) if (!e.End.HasValue) return e; return null; }
        }
        public static string CurrentDesc { get { Entry o = Open; return o != null ? o.Desc : PendingDesc; } }

        // ---- init / load / save
        public static void Init()
        {
            string env = Environment.GetEnvironmentVariable("WORKTIMER_DATA");
            Dir = !string.IsNullOrEmpty(env) ? env : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WorkTimer");
            Directory.CreateDirectory(Dir);
            entriesFile = Path.Combine(Dir, "entries.tsv");
            settingsFile = Path.Combine(Dir, "settings.ini");
            aliveFile = Path.Combine(Dir, "alive.txt");
            errorFile = Path.Combine(Dir, "error.log");
        }

        public static void LogError(Exception ex)
        {
            try { File.AppendAllText(errorFile, DateTime.Now.ToString(DtFmt) + "  " + (ex == null ? "unknown" : ex.ToString()) + "\r\n\r\n"); } catch { }
        }

        static string D(DateTime d) { return d.ToString(DtFmt, CultureInfo.InvariantCulture); }
        static bool P(string s, out DateTime d)
        {
            return DateTime.TryParseExact(s, DtFmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out d);
        }
        static string Esc(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\").Replace("\t", "\\t").Replace("\r", "").Replace("\n", "\\n");
        }
        static string Unesc(string s)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\\' && i + 1 < s.Length)
                {
                    char n = s[++i];
                    sb.Append(n == 't' ? '\t' : n == 'n' ? '\n' : n);
                }
                else sb.Append(c);
            }
            return sb.ToString();
        }

        static void WriteAtomic(string path, string text)
        {
            string tmp = path + ".tmp";
            File.WriteAllText(tmp, text, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(tmp, path, path + ".bak");
            else File.Move(tmp, path);
        }

        public static void Load()
        {
            // settings
            FirstRun = !File.Exists(settingsFile);
            try
            {
                if (!FirstRun)
                {
                    foreach (string line in File.ReadAllLines(settingsFile))
                    {
                        int i = line.IndexOf('=');
                        if (i < 1) continue;
                        string k = line.Substring(0, i).Trim(), v = line.Substring(i + 1).Trim();
                        double d; int n;
                        if (k == "DailyHours" && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d > 0 && d <= 24) DailyHours = d;
                        else if (k == "WeekStart" && int.TryParse(v, out n) && n >= 0 && n <= 6) WeekStart = (DayOfWeek)n;
                        else if (k == "AutoStart") AutoStart = v == "1";
                        else if (k == "OnTop") OnTop = v == "1";
                        else if (k == "Collapsed") Collapsed = v == "1";
                        else if (k == "AutoUpdate") AutoUpdate = v == "1";
                        else if (k == "Opacity" && int.TryParse(v, out n) && n >= 25 && n <= 100) OpacityPct = n;
                        else if (k == "WinX" && int.TryParse(v, out n)) WinX = n;
                        else if (k == "WinY" && int.TryParse(v, out n)) WinY = n;
                    }
                }
            }
            catch (Exception ex) { LogError(ex); }

            // entries
            Entries.Clear();
            try
            {
                if (File.Exists(entriesFile))
                {
                    foreach (string line in File.ReadAllLines(entriesFile, Encoding.UTF8))
                    {
                        string[] p = line.Split('\t');
                        if (p.Length < 3) continue;
                        DateTime s, e;
                        if (!P(p[0], out s)) continue;
                        Entry en = new Entry();
                        en.Start = s;
                        if (p[1].Length > 0) { if (!P(p[1], out e)) continue; en.End = e; }
                        en.Desc = Unesc(p[2]);
                        Entries.Add(en);
                    }
                }
            }
            catch (Exception ex) { LogError(ex); }

            // an entry still open means the app was killed / PC lost power: stop it at the last heartbeat
            Entry o = Open;
            if (o != null)
            {
                DateTime alive = o.Start, a;
                try { if (File.Exists(aliveFile) && P(File.ReadAllText(aliveFile).Trim(), out a)) alive = a; } catch { }
                DateTime now = Now();
                if (alive < o.Start) alive = o.Start;
                if (alive > now) alive = now;
                PendingDesc = o.Desc;
                CloseEntry(o, alive);
                // keep every entry inside one calendar day
                Entry cur = o;
                while (cur.End.HasValue && Entries.Contains(cur) && cur.End.Value > cur.Start.Date.AddDays(1))
                {
                    DateTime mid = cur.Start.Date.AddDays(1);
                    Entry nx = new Entry { Start = mid, End = cur.End, Desc = cur.Desc };
                    cur.End = mid;
                    Entries.Add(nx);
                    cur = nx;
                }
                RecoveredAt = alive;
                Entries.Sort(delegate(Entry x, Entry y) { return x.Start.CompareTo(y.Start); });
                SaveAll();
            }
            if (FirstRun) SaveSettings();
        }

        public static void SaveSettings()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("DailyHours=" + DailyHours.ToString(CultureInfo.InvariantCulture));
                sb.AppendLine("WeekStart=" + (int)WeekStart);
                sb.AppendLine("AutoStart=" + (AutoStart ? "1" : "0"));
                sb.AppendLine("OnTop=" + (OnTop ? "1" : "0"));
                sb.AppendLine("Collapsed=" + (Collapsed ? "1" : "0"));
                sb.AppendLine("AutoUpdate=" + (AutoUpdate ? "1" : "0"));
                sb.AppendLine("Opacity=" + OpacityPct);
                if (WinX != int.MinValue) { sb.AppendLine("WinX=" + WinX); sb.AppendLine("WinY=" + WinY); }
                WriteAtomic(settingsFile, sb.ToString());
                settingsDirty = false;
            }
            catch (Exception ex) { LogError(ex); }
        }

        public static void SaveAll()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (Entry e in Entries)
                    sb.Append(D(e.Start)).Append('\t').Append(e.End.HasValue ? D(e.End.Value) : "").Append('\t').Append(Esc(e.Desc)).Append("\r\n");
                WriteAtomic(entriesFile, sb.ToString());
                lastAlive = DateTime.Now;
                File.WriteAllText(aliveFile, D(Now()));
                dirty = false;
            }
            catch (Exception ex) { LogError(ex); }
        }

        public static void MarkSettingsDirty() { settingsDirty = true; }
        public static void RaiseChanged() { if (Changed != null) Changed(); }

        static void Commit()
        {
            Entries.Sort(delegate(Entry x, Entry y) { return x.Start.CompareTo(y.Start); });
            SaveAll();
            RaiseChanged();
        }

        static void CloseEntry(Entry o, DateTime end)
        {
            if (end <= o.Start) Entries.Remove(o); else o.End = end;
        }

        // ---- timer actions
        public static void Start()
        {
            if (Open != null) return;
            Entries.Add(new Entry { Start = Now(), Desc = PendingDesc });
            Commit();
        }

        // start the timer as if it had begun at 'at' (or move the start of the running entry to 'at')
        public static void StartFrom(DateTime at)
        {
            DateTime t = new DateTime(at.Ticks - at.Ticks % TimeSpan.TicksPerSecond);
            Entry o = Open;
            if (o != null) o.Start = t;
            else Entries.Add(new Entry { Start = t, Desc = PendingDesc });
            Commit();
        }

        public static void PauseAt(DateTime when)
        {
            Entry o = Open;
            if (o == null) return;
            DateTime t = new DateTime(when.Ticks - when.Ticks % TimeSpan.TicksPerSecond, when.Kind);
            PendingDesc = o.Desc;
            CloseEntry(o, t);
            Commit();
        }
        public static void Pause() { PauseAt(Now()); }

        public static void Slice()
        {
            Entry o = Open;
            if (o == null) return;
            DateTime n = Now();
            CloseEntry(o, n);
            PendingDesc = "";
            Entries.Add(new Entry { Start = n });
            Commit();
        }

        public static void SetDesc(string s)
        {
            Entry o = Open;
            if (o != null) o.Desc = s; else PendingDesc = s;
            dirty = true;
        }

        public static void Tick()
        {
            DateTime n = Now();
            Entry o = Open;
            bool rolled = false;
            while (o != null && o.Start.Date < n.Date)       // keep every entry inside one calendar day
            {
                DateTime mid = o.Start.Date.AddDays(1);
                o.End = mid;
                Entry nx = new Entry { Start = mid, Desc = o.Desc };
                Entries.Add(nx);
                o = nx;
                rolled = true;
            }
            if (rolled) { Commit(); return; }
            if (dirty || (DateTime.Now - lastAlive).TotalSeconds >= 20) SaveAll();
            if (settingsDirty) SaveSettings();
        }

        // ---- editing
        public static bool Overlaps(DateTime s, DateTime e, Entry ignore)
        {
            DateTime now = DateTime.Now;
            foreach (Entry x in Entries)
            {
                if (x == ignore) continue;
                DateTime xe = x.End.HasValue ? x.End.Value : now;
                if (s < xe && x.Start < e) return true;
            }
            return false;
        }
        public static void AddEntry(DateTime s, DateTime e, string d)
        {
            Entries.Add(new Entry { Start = s, End = e, Desc = d });
            Commit();
        }
        public static void EditEntry(Entry en, DateTime s, DateTime? e, string d)
        {
            en.Start = s; en.End = e; en.Desc = d;
            Commit();
        }
        public static void DeleteEntry(Entry en)
        {
            if (!en.End.HasValue) PendingDesc = en.Desc;
            Entries.Remove(en);
            Commit();
        }
        public static void SplitEntry(Entry en, DateTime at, string d1, string d2)
        {
            Entry second = new Entry { Start = at, End = en.End, Desc = d2 };
            en.End = at; en.Desc = d1;
            Entries.Add(second);
            Commit();
        }

        // ---- totals
        public static TimeSpan DayTotal(DateTime day, DateTime now)
        {
            TimeSpan t = TimeSpan.Zero;
            foreach (Entry e in Entries) if (e.Start.Date == day.Date) t += e.Dur(now);
            return t;
        }
        // when will the daily target be reached if you keep working from now? (rounded UP to the minute)
        public static bool TryFinish(DateTime now, out DateTime end)
        {
            end = now;
            if (Open == null) return false;
            TimeSpan t = DayTotal(now, now), tg = Target;
            if (t >= tg) return false;
            DateTime e = now + (tg - t);
            end = new DateTime((e.Ticks + TimeSpan.TicksPerMinute - 1) / TimeSpan.TicksPerMinute * TimeSpan.TicksPerMinute);
            return true;
        }

        // used after an automatic update: continue the task that was running
        public static void ResumeLast()
        {
            if (Open != null || Entries.Count == 0) return;
            Entry last = Entries[Entries.Count - 1];
            PendingDesc = last.Desc;
            Start();
        }

        public static bool HasEntries(DateTime day)
        {
            foreach (Entry e in Entries) if (e.Start.Date == day.Date) return true;
            return false;
        }
        public static DateTime WeekStartOf(DateTime d)
        {
            int diff = (7 + (int)d.DayOfWeek - (int)WeekStart) % 7;
            return d.Date.AddDays(-diff);
        }

        // ---- start with Windows
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        public static void ApplyAutoStart()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RunKey, true))
                {
                    if (k == null) return;
                    if (AutoStart) k.SetValue("WorkTimer", "\"" + Application.ExecutablePath + "\"");
                    else k.DeleteValue("WorkTimer", false);
                }
            }
            catch (Exception ex) { LogError(ex); }
        }
        public static void SetAutoStart(bool on)
        {
            AutoStart = on;
            ApplyAutoStart();
            SaveSettings();
            RaiseChanged();
        }
    }

    // ------------------------------------------------------------------ Excel export (.xlsx, no libraries)
    static class Xlsx
    {
        // style ids (see Styles())
        const int SDef = 0, SHead = 1, SDate = 2, STime = 3, SDur = 4, STotText = 5, STotDur = 6, SOver = 7, STitle = 8;

        class Sheet
        {
            public StringBuilder Rows = new StringBuilder();
            StringBuilder cur = new StringBuilder();
            int row = 0;
            public double[] Widths;
            public void Begin() { row++; cur = new StringBuilder(); }
            public void End() { Rows.Append("<row r=\"").Append(row).Append("\">").Append(cur).Append("</row>"); }
            public void Skip() { row++; }
            static string Ref(int col, int row) { return ((char)('A' + col)).ToString() + row; }
            public void Str(int col, string s, int style)
            {
                cur.Append("<c r=\"").Append(Ref(col, row)).Append("\" t=\"inlineStr\" s=\"").Append(style)
                   .Append("\"><is><t xml:space=\"preserve\">").Append(Clean(s)).Append("</t></is></c>");
            }
            public void Num(int col, double v, int style)
            {
                cur.Append("<c r=\"").Append(Ref(col, row)).Append("\" s=\"").Append(style).Append("\"><v>")
                   .Append(v.ToString("R", CultureInfo.InvariantCulture)).Append("</v></c>");
            }
            public void Blank(int col, int style)
            {
                cur.Append("<c r=\"").Append(Ref(col, row)).Append("\" s=\"").Append(style).Append("\"/>");
            }
            public string Xml()
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                sb.Append("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><cols>");
                for (int i = 0; i < Widths.Length; i++)
                    sb.Append("<col min=\"").Append(i + 1).Append("\" max=\"").Append(i + 1).Append("\" width=\"")
                      .Append(Widths[i].ToString(CultureInfo.InvariantCulture)).Append("\" customWidth=\"1\"/>");
                sb.Append("</cols><sheetData>").Append(Rows).Append("</sheetData></worksheet>");
                return sb.ToString();
            }
        }

        static string Clean(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in (s ?? ""))
                if (c == '\t' || c == '\n' || c == '\r' || c >= 0x20) sb.Append(c);
            return SecurityElement.Escape(sb.ToString());
        }
        static double Serial(DateTime d) { return (d.Date - new DateTime(1899, 12, 30)).TotalDays; }
        static double Frac(TimeSpan t) { return Math.Round(t.TotalSeconds) / 86400.0; }
        static double TimeOfDay(DateTime d) { return d.TimeOfDay.TotalSeconds / 86400.0; }

        public static void Export(string path, DateTime weekStart)
        {
            DateTime now = DateTime.Now;
            TimeSpan target = Core.Target;
            Sheet sum = new Sheet { Widths = new double[] { 13, 13, 12, 12, 13, 13, 13, 13 } };
            Sheet det = new Sheet { Widths = new double[] { 13, 13, 10, 10, 12, 70 } };

            sum.Begin(); sum.Str(0, "Weekly report  " + weekStart.ToString("d MMM yyyy") + " \u2013 " + weekStart.AddDays(6).ToString("d MMM yyyy"), STitle); sum.End();
            sum.Skip();
            sum.Begin();
            string[] h1 = { "Date", "Day", "First start", "Last end", "Worked", "Daily target", "Overtime", "Short by" };
            for (int i = 0; i < h1.Length; i++) sum.Str(i, h1[i], SHead);
            sum.End();

            det.Begin();
            string[] h2 = { "Date", "Day", "From", "To", "Duration", "Description" };
            for (int i = 0; i < h2.Length; i++) det.Str(i, h2[i], SHead);
            det.End();

            TimeSpan totW = TimeSpan.Zero, totO = TimeSpan.Zero, totS = TimeSpan.Zero, totT = TimeSpan.Zero;
            for (int i = 0; i < 7; i++)
            {
                DateTime d = weekStart.AddDays(i);
                List<Entry> list = Core.Entries.Where(e => e.Start.Date == d.Date).OrderBy(e => e.Start).ToList();
                TimeSpan w = TimeSpan.Zero;
                foreach (Entry e in list)
                {
                    TimeSpan du = e.Dur(now);
                    w += du;
                    DateTime end = e.End.HasValue ? e.End.Value : now;
                    det.Begin();
                    det.Num(0, Serial(d), SDate);
                    det.Str(1, d.ToString("dddd"), SDef);
                    det.Num(2, TimeOfDay(e.Start), STime);
                    if (end.Date > e.Start.Date) det.Num(3, 1.0, STime); else det.Num(3, TimeOfDay(end), STime);
                    det.Num(4, Frac(du), SDur);
                    det.Str(5, e.Desc + (e.End.HasValue ? "" : "  (still running)"), SDef);
                    det.End();
                }
                sum.Begin();
                sum.Num(0, Serial(d), SDate);
                sum.Str(1, d.ToString("dddd"), SDef);
                if (list.Count == 0)
                {
                    sum.Blank(2, SDef); sum.Blank(3, SDef); sum.Blank(4, SDef); sum.Blank(5, SDef); sum.Blank(6, SDef); sum.Blank(7, SDef);
                }
                else
                {
                    DateTime first = list[0].Start;
                    Entry lastE = list[list.Count - 1];
                    DateTime last = lastE.End.HasValue ? lastE.End.Value : now;
                    sum.Num(2, TimeOfDay(first), STime);
                    if (last.Date > first.Date) sum.Num(3, 1.0, STime); else sum.Num(3, TimeOfDay(last), STime);
                    sum.Num(4, Frac(w), SDur);
                    sum.Num(5, Frac(target), SDur);
                    TimeSpan ot = w > target ? w - target : TimeSpan.Zero;
                    TimeSpan sh = w < target ? target - w : TimeSpan.Zero;
                    sum.Num(6, Frac(ot), ot > TimeSpan.Zero ? SOver : SDur);
                    sum.Num(7, Frac(sh), SDur);
                    totW += w; totT += target; totO += ot; totS += sh;
                }
                sum.End();
            }
            sum.Begin();
            sum.Str(0, "Total", STotText); sum.Blank(1, STotText); sum.Blank(2, STotText); sum.Blank(3, STotText);
            sum.Num(4, Frac(totW), STotDur); sum.Num(5, Frac(totT), STotDur); sum.Num(6, Frac(totO), STotDur); sum.Num(7, Frac(totS), STotDur);
            sum.End();

            string ct = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>";
            string rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
            string wb = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<sheets><sheet name=\"Week Summary\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Details\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>";
            string wbRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
                + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>";

            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                Put(zip, "[Content_Types].xml", ct);
                Put(zip, "_rels/.rels", rels);
                Put(zip, "xl/workbook.xml", wb);
                Put(zip, "xl/_rels/workbook.xml.rels", wbRels);
                Put(zip, "xl/styles.xml", Styles());
                Put(zip, "xl/worksheets/sheet1.xml", sum.Xml());
                Put(zip, "xl/worksheets/sheet2.xml", det.Xml());
            }
        }

        static void Put(ZipArchive zip, string name, string content)
        {
            ZipArchiveEntry en = zip.CreateEntry(name);
            using (Stream s = en.Open())
            {
                byte[] b = new UTF8Encoding(false).GetBytes(content);
                s.Write(b, 0, b.Length);
            }
        }

        static string Styles()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<numFmts count=\"3\"><numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/><numFmt numFmtId=\"165\" formatCode=\"h:mm AM/PM\"/><numFmt numFmtId=\"166\" formatCode=\"[h]:mm:ss\"/></numFmts>"
                + "<fonts count=\"5\">"
                + "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
                + "<font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/></font>"
                + "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>"
                + "<font><b/><sz val=\"11\"/><color rgb=\"FFC55A11\"/><name val=\"Calibri\"/></font>"
                + "<font><b/><sz val=\"14\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"4\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
                + "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF2F5D50\"/><bgColor indexed=\"64\"/></patternFill></fill>"
                + "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE3ECE8\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"9\">"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>"
                + "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                + "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                + "<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>"
                + "<xf numFmtId=\"166\" fontId=\"2\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\"/>"
                + "<xf numFmtId=\"166\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
        }
    }

    // ------------------------------------------------------------------ small UI helpers
    static class UI
    {
        public static Label Lbl(string t, int x, int y, int w, int h)
        {
            return new Label { Text = t, Left = x, Top = y, Width = w, Height = h, TextAlign = ContentAlignment.MiddleLeft };
        }
        public static Button Btn(string t, int x, int y, int w, int h, EventHandler click)
        {
            Button b = new Button { Text = t, Left = x, Top = y, Width = w, Height = h, UseVisualStyleBackColor = true };
            if (click != null) b.Click += click;
            return b;
        }
        public static void Prep(Form f, string title, int w, int h)
        {
            f.AutoScaleDimensions = new SizeF(96F, 96F);
            f.AutoScaleMode = AutoScaleMode.Dpi;
            f.Font = new Font("Segoe UI", 9F);
            f.Text = title;
            f.ClientSize = new Size(w, h);
        }
        public static DateTimePicker TimePicker(DateTime v, int x, int y, int w)
        {
            return new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "hh:mm tt", ShowUpDown = true, Left = x, Top = y, Width = w + 20, Value = v };
        }
        // keep seconds if the user did not touch the hour/minute
        public static DateTime Compose(DateTime day, DateTime picked, DateTime original)
        {
            if (picked.Hour == original.Hour && picked.Minute == original.Minute) return original;
            return new DateTime(day.Year, day.Month, day.Day, picked.Hour, picked.Minute, 0);
        }
    }

    class EntryDialog : Form
    {
        DateTimePicker dtFrom, dtTo; TextBox txt;
        Entry editing; DateTime day, origStart, origEnd; bool running;
        public DateTime ResStart; public DateTime? ResEnd; public string ResDesc;

        public EntryDialog(string title, DateTime day, DateTime start, DateTime end, string desc, bool running, Entry editing)
        {
            this.day = day; this.origStart = start; this.origEnd = end; this.running = running; this.editing = editing;
            UI.Prep(this, title, 380, 170);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Controls.Add(UI.Lbl("From", 14, 16, 50, 24));
            dtFrom = UI.TimePicker(start, 70, 16, 80); Controls.Add(dtFrom);
            Controls.Add(UI.Lbl("To", 180, 16, 30, 24));
            dtTo = UI.TimePicker(end, 214, 16, 80); Controls.Add(dtTo);
            if (running) { dtTo.Enabled = false; Controls.Add(UI.Lbl("running", 322, 16, 52, 24)); }
            Controls.Add(UI.Lbl("What", 14, 56, 50, 24));
            txt = new TextBox { Left = 70, Top = 57, Width = 296, Text = desc };
            Controls.Add(txt);
            Button ok = UI.Btn("OK", 196, 120, 80, 28, OkClick);
            Button cancel = UI.Btn("Cancel", 286, 120, 80, 28, null);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(ok); Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
            Shown += delegate { txt.Focus(); txt.SelectAll(); };
        }

        void OkClick(object s, EventArgs e)
        {
            DateTime st = UI.Compose(day, dtFrom.Value, origStart);
            DateTime en = running ? DateTime.Now : UI.Compose(day, dtTo.Value, origEnd);
            if (en <= st) { MessageBox.Show(this, "The end time must be after the start time.", "Work Timer"); return; }
            if (running && st > DateTime.Now) { MessageBox.Show(this, "The start time is in the future.", "Work Timer"); return; }
            if (Core.Overlaps(st, en, editing)) { MessageBox.Show(this, "This overlaps another entry. Adjust the times.", "Work Timer"); return; }
            ResStart = st; ResEnd = running ? (DateTime?)null : en; ResDesc = txt.Text.Trim();
            DialogResult = DialogResult.OK;
        }
    }

    class SplitDialog : Form
    {
        DateTimePicker dt; TextBox t1, t2; Entry en; DateTime day, mid;
        public DateTime ResAt; public string D1, D2;

        public SplitDialog(Entry en, DateTime day)
        {
            this.en = en; this.day = day;
            DateTime end = en.End.HasValue ? en.End.Value : DateTime.Now;
            mid = en.Start.AddTicks((end - en.Start).Ticks / 2);
            mid = new DateTime(mid.Ticks - mid.Ticks % TimeSpan.TicksPerSecond);
            UI.Prep(this, "Split entry", 400, 200);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Controls.Add(UI.Lbl(Fmt.Hm(en.Start) + " \u2013 " + (en.End.HasValue ? Fmt.Hm(end) : "now") + "  will be cut at:", 14, 14, 250, 24));
            dt = UI.TimePicker(mid, 270, 14, 80); Controls.Add(dt);
            Controls.Add(UI.Lbl("Before", 14, 54, 60, 24));
            t1 = new TextBox { Left = 80, Top = 55, Width = 306, Text = en.Desc }; Controls.Add(t1);
            Controls.Add(UI.Lbl("After", 14, 90, 60, 24));
            t2 = new TextBox { Left = 80, Top = 91, Width = 306, Text = en.Desc }; Controls.Add(t2);
            Button ok = UI.Btn("Split", 216, 150, 80, 28, OkClick);
            Button cancel = UI.Btn("Cancel", 306, 150, 80, 28, null);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(ok); Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
        }

        void OkClick(object s, EventArgs e)
        {
            DateTime at = UI.Compose(day, dt.Value, mid);
            DateTime end = en.End.HasValue ? en.End.Value : DateTime.Now;
            if (at <= en.Start || at >= end) { MessageBox.Show(this, "Pick a time inside the entry.", "Work Timer"); return; }
            ResAt = at; D1 = t1.Text.Trim(); D2 = t2.Text.Trim();
            DialogResult = DialogResult.OK;
        }
    }

    class StartFromDialog : Form
    {
        DateTimePicker dt; DateTime init; bool running;
        public DateTime Res;

        public StartFromDialog(DateTime init, bool running)
        {
            this.init = init; this.running = running;
            UI.Prep(this, "Start from a chosen time", 340, 150);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen; TopMost = true;
            Controls.Add(UI.Lbl(running ? "Change today's start time to:" : "I started working today at:", 14, 18, 200, 24));
            dt = UI.TimePicker(init, 220, 18, 80); Controls.Add(dt);
            Controls.Add(UI.Lbl("Pick today's time. Click AM/PM to switch it.", 14, 52, 320, 22));
            Button ok = UI.Btn(running ? "Change" : "Start", 154, 100, 80, 28, OkClick);
            Button cancel = UI.Btn("Cancel", 244, 100, 80, 28, null);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(ok); Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
        }

        void OkClick(object s, EventArgs e)
        {
            DateTime today = DateTime.Today, now = DateTime.Now;
            DateTime t = UI.Compose(today, dt.Value, init);
            if (t > now) { MessageBox.Show(this, "That time is in the future.", "Work Timer"); return; }
            if (Core.Overlaps(t, now, Core.Open)) { MessageBox.Show(this, "That overlaps an entry you already have today. Pick a later time, or fix the entry in the Log.", "Work Timer"); return; }
            Res = t;
            DialogResult = DialogResult.OK;
        }
    }

    // ------------------------------------------------------------------ log window
    class LogForm : Form
    {
        DateTimePicker dtp; Button bPrev, bNext, bToday, bAdd, bEdit, bSplit, bDel, bExport;
        Label lblSum, lblWeek; ListView lv; NumericUpDown nud; ComboBox cmb; CheckBox chk;
        bool loading;
        static readonly Color Orange = Color.FromArgb(200, 85, 20), Green = Color.FromArgb(30, 130, 76);

        public LogForm()
        {
            UI.Prep(this, "Work Timer \u2013 Log", 780, 520);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(700, 440);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            Panel top = new Panel { Dock = DockStyle.Top, Size = new Size(780, 96) };
            bPrev = UI.Btn("\u25C0", 10, 10, 34, 28, delegate { Shift(-1); });
            dtp = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "ddd  dd MMM yyyy", Left = 50, Top = 11, Width = 170 };
            bNext = UI.Btn("\u25B6", 226, 10, 34, 28, delegate { Shift(1); });
            bToday = UI.Btn("Today", 266, 10, 70, 28, delegate { dtp.Value = DateTime.Today; });
            lblSum = UI.Lbl("", 10, 46, 750, 24); lblSum.Font = new Font("Segoe UI Semibold", 11F); lblSum.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblWeek = UI.Lbl("", 10, 70, 750, 20); lblWeek.ForeColor = Color.FromArgb(90, 90, 90); lblWeek.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            top.Controls.AddRange(new Control[] { bPrev, dtp, bNext, bToday, lblSum, lblWeek });

            Panel bottom = new Panel { Dock = DockStyle.Bottom, Size = new Size(780, 92) };
            bAdd = UI.Btn("Add entry", 10, 10, 90, 30, delegate { AddClick(); });
            bEdit = UI.Btn("Edit", 106, 10, 70, 30, delegate { EditClick(); });
            bSplit = UI.Btn("Split", 182, 10, 70, 30, delegate { SplitClick(); });
            bDel = UI.Btn("Delete", 258, 10, 70, 30, delegate { DeleteClick(); });
            bExport = UI.Btn("Export week to Excel", 580, 10, 190, 30, delegate { ExportClick(); });
            bExport.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Label l1 = UI.Lbl("Daily hours", 10, 54, 70, 24);
            nud = new NumericUpDown { Left = 84, Top = 55, Width = 60, DecimalPlaces = 2, Increment = 0.25M, Minimum = 0.25M, Maximum = 24 };
            Label l2 = UI.Lbl("Week starts on", 164, 54, 90, 24);
            cmb = new ComboBox { Left = 258, Top = 55, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            for (int i = 0; i < 7; i++) cmb.Items.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetDayName((DayOfWeek)i));
            chk = new CheckBox { Left = 380, Top = 54, Width = 170, Height = 24, Text = "Start with Windows" };
            bottom.Controls.AddRange(new Control[] { bAdd, bEdit, bSplit, bDel, bExport, l1, nud, l2, cmb, chk });

            lv = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, MultiSelect = false, HideSelection = false };
            lv.Columns.Add("From", 70); lv.Columns.Add("To", 70); lv.Columns.Add("Duration", 90); lv.Columns.Add("What I was doing", 500);
            lv.DoubleClick += delegate { EditClick(); };
            lv.SelectedIndexChanged += delegate { UpdateButtons(); };
            lv.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Delete) DeleteClick(); };
            lv.Resize += delegate { FitColumn(); };

            Controls.Add(lv); Controls.Add(bottom); Controls.Add(top);

            loading = true;
            dtp.Value = DateTime.Today;
            LoadSettingsToUi();
            loading = false;
            dtp.ValueChanged += delegate { Reload(); };
            nud.ValueChanged += delegate { if (loading) return; Core.DailyHours = (double)nud.Value; Core.SaveSettings(); Core.RaiseChanged(); };
            cmb.SelectedIndexChanged += delegate { if (loading) return; Core.WeekStart = (DayOfWeek)cmb.SelectedIndex; Core.SaveSettings(); Core.RaiseChanged(); };
            chk.CheckedChanged += delegate { if (loading) return; Core.SetAutoStart(chk.Checked); };

            Core.Changed += OnCoreChanged;
            FormClosed += delegate { Core.Changed -= OnCoreChanged; };
            Shown += delegate { FitColumn(); };
            Reload();
        }

        void LoadSettingsToUi()
        {
            bool l = loading; loading = true;
            nud.Value = (decimal)Core.DailyHours;
            cmb.SelectedIndex = (int)Core.WeekStart;
            chk.Checked = Core.AutoStart;
            loading = l;
        }

        void OnCoreChanged()
        {
            LoadSettingsToUi();
            Reload();
        }

        public void GoToday() { dtp.Value = DateTime.Today; }

        void FitColumn()
        {
            int w = lv.ClientSize.Width - lv.Columns[0].Width - lv.Columns[1].Width - lv.Columns[2].Width;
            if (w > 100) lv.Columns[3].Width = w;
        }

        void Shift(int d) { dtp.Value = dtp.Value.Date.AddDays(d); }
        DateTime Day { get { return dtp.Value.Date; } }
        Entry Selected { get { return lv.SelectedItems.Count > 0 ? lv.SelectedItems[0].Tag as Entry : null; } }

        void UpdateButtons()
        {
            bool s = Selected != null;
            bEdit.Enabled = s; bDel.Enabled = s; bSplit.Enabled = s;
        }

        void Reload()
        {
            Entry sel = Selected;
            DateTime now = DateTime.Now, day = Day;
            lv.BeginUpdate();
            lv.Items.Clear();
            foreach (Entry e in Core.Entries.Where(x => x.Start.Date == day).OrderBy(x => x.Start))
            {
                string to = !e.End.HasValue ? "running" : Fmt.Hm(e.End.Value);
                ListViewItem li = new ListViewItem(Fmt.Hm(e.Start));
                li.SubItems.Add(to); li.SubItems.Add(Fmt.Dur(e.Dur(now))); li.SubItems.Add(e.Desc);
                li.Tag = e;
                if (!e.End.HasValue) li.ForeColor = Green;
                lv.Items.Add(li);
                if (e == sel) li.Selected = true;
            }
            lv.EndUpdate();
            UpdateSummary();
            UpdateButtons();
        }

        public void Tick()
        {
            DateTime now = DateTime.Now;
            foreach (ListViewItem li in lv.Items)
            {
                Entry e = li.Tag as Entry;
                if (e != null && !e.End.HasValue)
                {
                    li.SubItems[2].Text = Fmt.Dur(e.Dur(now));
                    if (li.SubItems[3].Text != e.Desc) li.SubItems[3].Text = e.Desc;
                }
            }
            UpdateSummary();
        }

        void UpdateSummary()
        {
            DateTime now = DateTime.Now, day = Day;
            TimeSpan total = Core.DayTotal(day, now), target = Core.Target;
            string s = "Worked " + Fmt.Dur(total) + "      Target " + Fmt.Dur(target) + "      ";
            if (total >= target) { s += "Overtime +" + Fmt.Dur(total - target); lblSum.ForeColor = Orange; }
            else { s += "Left " + Fmt.Dur(target - total); lblSum.ForeColor = SystemColors.ControlText; }
            lblSum.Text = s;

            DateTime ws = Core.WeekStartOf(day);
            TimeSpan wt = TimeSpan.Zero, ot = TimeSpan.Zero;
            for (int i = 0; i < 7; i++)
            {
                TimeSpan t = Core.DayTotal(ws.AddDays(i), now);
                wt += t;
                if (t > target) ot += t - target;
            }
            lblWeek.Text = "Week " + ws.ToString("d MMM") + " \u2013 " + ws.AddDays(6).ToString("d MMM") + "      Worked " + Fmt.Dur(wt) + "      Overtime " + Fmt.Dur(ot);
        }

        void AddClick()
        {
            DateTime day = Day, start = day.AddHours(9);
            DateTime last = DateTime.MinValue;
            foreach (Entry x in Core.Entries) if (x.Start.Date == day && x.End.HasValue && x.End.Value > last && x.End.Value.Date == day) last = x.End.Value;
            if (last > DateTime.MinValue) start = last;
            if (start.Date != day) start = day.AddHours(9);
            DateTime end = start.AddHours(1);
            if (end.Date != day) end = day.AddHours(23).AddMinutes(59);
            using (EntryDialog dlg = new EntryDialog("Add entry", day, start, end, "", false, null))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    DateTime ns = dlg.ResStart;
                    DateTime ne = dlg.ResEnd ?? ns.AddHours(1);
                    Core.AddEntry(ns, ne, dlg.ResDesc);
                }
            }
        }

        void EditClick()
        {
            Entry en = Selected;
            if (en == null) return;
            bool running = !en.End.HasValue;
            using (EntryDialog dlg = new EntryDialog("Edit entry", en.Start.Date, en.Start, running ? DateTime.Now : en.End.Value, en.Desc, running, en))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK) Core.EditEntry(en, dlg.ResStart, dlg.ResEnd, dlg.ResDesc);
            }
        }

        void SplitClick()
        {
            Entry en = Selected;
            if (en == null) return;
            using (SplitDialog dlg = new SplitDialog(en, en.Start.Date))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK) Core.SplitEntry(en, dlg.ResAt, dlg.D1, dlg.D2);
            }
        }

        void DeleteClick()
        {
            Entry en = Selected;
            if (en == null) return;
            if (MessageBox.Show(this, "Delete this entry?\n\n" + Fmt.Hm(en.Start) + "  " + en.Desc, "Work Timer", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                Core.DeleteEntry(en);
        }

        void ExportClick()
        {
            DateTime ws = Core.WeekStartOf(Day);
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel workbook (*.xlsx)|*.xlsx";
                sfd.FileName = "WorkTimer_Week_" + ws.ToString("yyyy-MM-dd") + ".xlsx";
                sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (sfd.ShowDialog(this) != DialogResult.OK) return;
                try { Xlsx.Export(sfd.FileName, ws); }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Could not save the file.\n(Is it open in Excel?)\n\n" + ex.Message, "Work Timer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try { Process.Start(sfd.FileName); }
                catch { MessageBox.Show(this, "Saved to:\n" + sfd.FileName, "Work Timer"); }
            }
        }
    }

    // colour of the day: RED until the daily hours are done, GREEN once you are in overtime
    static class Palette
    {
        public static readonly Color Red = Color.FromArgb(235, 64, 52);
        public static readonly Color Green = Color.FromArgb(46, 196, 120);
        public static Color For(double progress) { return progress >= 1.0 ? Green : Red; }
        public static Color Mix(Color a, Color b, double t)
        {
            return Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));
        }
    }

    // ------------------------------------------------------------------ modern ring clock face
    class ClockFace : Control
    {
        public double Progress;
        public bool Running;
        public string Main = "00:00", Sec = "00";
        public int SecondOfMinute;
        public Color Accent = Color.Gray;
        public string Note = "";                       // small line under the seconds: "ends 12:30 AM" / "OT +01:30" / "paused"
        public Color NoteColor = Color.Gray;
        Font fMain, fSec, fNote;

        public ClockFace()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            fMain = new Font("Segoe UI Semibold", 17F);
            fSec = new Font("Segoe UI", 8.5F);
            fNote = new Font("Segoe UI Semibold", 7.5F);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { fMain.Dispose(); fSec.Dispose(); fNote.Dispose(); }
            base.Dispose(disposing);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float th = Math.Max(4f, Width * 0.085f);
            RectangleF r = new RectangleF(th / 2 + 2, th / 2 + 2, Width - th - 4, Height - th - 4);
            using (Pen track = new Pen(Color.FromArgb(48, 55, 64), th)) g.DrawEllipse(track, r);
            if (Progress > 0.002)
            {
                using (Pen pen = new Pen(Accent, th))
                {
                    pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round;
                    if (Progress >= 0.999) g.DrawEllipse(pen, r);
                    else g.DrawArc(pen, r, -90f, (float)(Progress * 360.0));
                }
            }
            if (Running)    // little "second hand" dot travelling around the ring
            {
                double a = SecondOfMinute / 60.0 * Math.PI * 2 - Math.PI / 2;
                float cx = r.X + r.Width / 2, cy = r.Y + r.Height / 2, rad = r.Width / 2, d = th * 0.55f;
                using (SolidBrush b = new SolidBrush(Color.White))
                    g.FillEllipse(b, cx + (float)Math.Cos(a) * rad - d / 2, cy + (float)Math.Sin(a) * rad - d / 2, d, d);
            }
            Rectangle top = new Rectangle(0, (int)(Height * 0.19), Width, (int)(Height * 0.30));
            Rectangle bot = new Rectangle(0, (int)(Height * 0.48), Width, (int)(Height * 0.17));
            Rectangle note = new Rectangle(0, (int)(Height * 0.66), Width, (int)(Height * 0.17));
            TextFormatFlags f = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
            TextRenderer.DrawText(g, Main, fMain, top, Color.FromArgb(245, 247, 250), f);
            TextRenderer.DrawText(g, ":" + Sec, fSec, bot, Accent, f);
            if (!string.IsNullOrEmpty(Note)) TextRenderer.DrawText(g, Note, fNote, note, NoteColor, f);
        }
    }

    // ------------------------------------------------------------------ the small always-available timer window
    class MiniForm : Form
    {
        ClockFace face; Label lblStatus, lblGear, lblClose, lblToggle; Panel pnlTxt; TextBox txt;
        Button bToggle, bFrom, bSlice, bLog;
        ContextMenuStrip gear; TrackBar tb; Label lblTr; ToolStripMenuItem miTop, miPause, miControls;
        ToolTip tip = new ToolTip();
        string lastTip = "";
        bool sync, hidTold, collapsed;
        float sc = 1f;
        public event Action ToggleRequested, LogRequested, HidFirstTime, FromRequested;
        public bool HotkeyOk;
        public bool Quiet;      // show without grabbing keyboard focus (used at startup)
        protected override bool ShowWithoutActivation { get { return Quiet; } }

        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr h, int id, uint mod, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr h, int id);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
        [DllImport("user32.dll", EntryPoint = "SendMessage")] static extern IntPtr SendMessageI(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("gdi32.dll")] static extern IntPtr CreateRoundRectRgn(int l, int t, int r, int b, int w, int h);
        [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr o);

        static readonly Color BgC = Color.FromArgb(27, 31, 36), Card = Color.FromArgb(38, 44, 52), BtnC = Color.FromArgb(52, 60, 70),
            Muted = Color.FromArgb(145, 155, 166), Green = Color.FromArgb(46, 196, 120),
            Grey = Color.FromArgb(110, 120, 132), Light = Color.FromArgb(245, 247, 250);

        static Color Lighten(Color c, int n) { return Color.FromArgb(Math.Min(255, c.R + n), Math.Min(255, c.G + n), Math.Min(255, c.B + n)); }
        static void Tint(Button b, Color back, Color fore)
        {
            b.BackColor = back; b.ForeColor = fore;
            b.FlatAppearance.MouseOverBackColor = Lighten(back, 16);
            b.FlatAppearance.MouseDownBackColor = Lighten(back, 30);
        }
        static Button Flat(string t, int x, int w, EventHandler click)
        {
            Button b = new Button { Text = t, Left = x, Top = 82, Width = w, Height = 36, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false, Font = new Font("Segoe UI Semibold", 9F), TabStop = false };
            b.FlatAppearance.BorderSize = 0;
            Tint(b, BtnC, Light);
            if (click != null) b.Click += click;
            return b;
        }

        public MiniForm()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) sc = g.DpiX / 96f;
            UI.Prep(this, "Work Timer", 400, 136);
            FormBorderStyle = FormBorderStyle.None;
            ClientSize = new Size(400, 136);
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false; TopMost = Core.OnTop;
            BackColor = BgC; ForeColor = Light;
            StartPosition = FormStartPosition.Manual;
            Opacity = Core.OpacityPct / 100.0;

            face = new ClockFace { Left = 12, Top = 16, Width = 104, Height = 104, BackColor = BgC };
            lblToggle = new Label { Left = 326, Top = 10, Width = 24, Height = 24, Text = "«", ForeColor = Muted, BackColor = BgC, Font = new Font("Segoe UI", 12F), TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            lblToggle.MouseEnter += delegate { lblToggle.ForeColor = Light; };
            lblToggle.MouseLeave += delegate { lblToggle.ForeColor = Muted; };
            lblToggle.Click += delegate { SetCollapsed(!collapsed, true); };
            lblStatus = new Label { Left = 128, Top = 12, Width = 194, Height = 22, ForeColor = Muted, Font = new Font("Segoe UI", 8.5F), TextAlign = ContentAlignment.MiddleLeft };
            lblGear = new Label { Left = 350, Top = 10, Width = 24, Height = 24, Text = "⚙", ForeColor = Muted, Font = new Font("Segoe UI Symbol", 12F), TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            lblClose = new Label { Left = 374, Top = 10, Width = 22, Height = 24, Text = "✕", ForeColor = Muted, Font = new Font("Segoe UI Symbol", 9F), TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            foreach (Label l in new Label[] { lblGear, lblClose })
            {
                Label x = l;
                x.MouseEnter += delegate { x.ForeColor = Light; };
                x.MouseLeave += delegate { x.ForeColor = Muted; };
            }
            lblClose.Click += delegate { Close(); };     // OnFormClosing turns this into "hide to tray"

            pnlTxt = new Panel { Left = 128, Top = 38, Width = 260, Height = 32, BackColor = Card };
            txt = new TextBox { Left = 10, Top = 7, Width = 240, BorderStyle = BorderStyle.None, BackColor = Card, ForeColor = Light, Font = new Font("Segoe UI", 9.5F), MaxLength = 300 };
            pnlTxt.Controls.Add(txt);

            bToggle = Flat("Start", 128, 88, delegate { if (ToggleRequested != null) ToggleRequested(); });
            bFrom = Flat("From…", 220, 62, delegate { if (FromRequested != null) FromRequested(); });
            bSlice = Flat("Slice", 286, 54, delegate { Core.Slice(); txt.Focus(); });
            bLog = Flat("Log", 344, 44, delegate { if (LogRequested != null) LogRequested(); });

            txt.TextChanged += delegate { if (!sync) Core.SetDesc(txt.Text); };
            txt.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; bToggle.Focus(); } };
            Controls.AddRange(new Control[] { face, lblStatus, lblGear, lblClose, pnlTxt, bToggle, bFrom, bSlice, bLog, lblToggle });
            lblToggle.BringToFront();

            // drag the window from any empty area; right-click anywhere on it = menu
            foreach (Control c in new Control[] { this, face, lblStatus })
            {
                c.MouseDown += delegate(object s, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left) return;
                    if (e.Clicks == 2) { SetCollapsed(!collapsed, true); return; }      // double-click = collapse / expand
                    ReleaseCapture(); SendMessageI(Handle, 0xA1, (IntPtr)2, IntPtr.Zero);   // otherwise drag the window
                };
                c.MouseUp += delegate(object s, MouseEventArgs e) { if (e.Button == MouseButtons.Right) gear.Show(Cursor.Position); };
            }

            BuildGearMenu();
            lblGear.Click += delegate { gear.Show(lblGear, new Point(-(int)(150 * sc), lblGear.Height)); };

            ApplyLayout(Core.Collapsed);
            Core.Changed += OnCoreChanged;
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            Point p = new Point(wa.Right - Width - 20, wa.Top + 20);
            if (Core.WinX != int.MinValue)
            {
                Rectangle r = new Rectangle(Core.WinX, Core.WinY, Width, Height);
                foreach (Screen scr in Screen.AllScreens) if (scr.WorkingArea.IntersectsWith(r)) { p = new Point(Core.WinX, Core.WinY); break; }
            }
            Location = p;
            SyncText();
            RefreshView();
        }

        void BuildGearMenu()
        {
            gear = new ContextMenuStrip();
            miPause = new ToolStripMenuItem("Pause", null, delegate { if (ToggleRequested != null) ToggleRequested(); });
            miPause.Font = new Font(miPause.Font, FontStyle.Bold);
            miControls = new ToolStripMenuItem("Show buttons && description", null, delegate { SetCollapsed(!collapsed, true); });
            gear.Items.Add(miPause);
            gear.Items.Add(miControls);
            gear.Items.Add(new ToolStripSeparator());
            gear.Opening += delegate
            {
                miPause.Text = Core.Open != null ? "Pause" : (Core.HasEntries(DateTime.Now) ? "Resume" : "Start");
                miControls.Checked = !collapsed;
            };
            Panel host = new Panel { Width = (int)(230 * sc), Height = (int)(56 * sc), BackColor = Color.White };
            lblTr = new Label { Left = (int)(12 * sc), Top = (int)(6 * sc), Width = (int)(200 * sc), Height = (int)(18 * sc), Font = new Font("Segoe UI", 9F) };
            tb = new TrackBar { Left = (int)(6 * sc), Top = (int)(26 * sc), Width = (int)(218 * sc), Height = (int)(28 * sc), Minimum = 0, Maximum = 75, TickStyle = TickStyle.None, SmallChange = 5, LargeChange = 10 };
            tb.Value = Math.Max(0, Math.Min(75, 100 - Core.OpacityPct));
            lblTr.Text = "Transparency   " + tb.Value + "%";
            host.Controls.Add(lblTr); host.Controls.Add(tb);
            tb.ValueChanged += delegate
            {
                Core.OpacityPct = 100 - tb.Value;
                Opacity = Core.OpacityPct / 100.0;
                lblTr.Text = "Transparency   " + tb.Value + "%";
                Core.MarkSettingsDirty();
            };
            gear.Items.Add(new ToolStripControlHost(host) { AutoSize = false, Size = host.Size, Margin = Padding.Empty, Padding = Padding.Empty });
            gear.Items.Add(new ToolStripSeparator());
            miTop = new ToolStripMenuItem("Always on top") { CheckOnClick = true, Checked = Core.OnTop };
            miTop.CheckedChanged += delegate { Core.OnTop = miTop.Checked; TopMost = Core.OnTop; Core.MarkSettingsDirty(); };
            gear.Items.Add(miTop);
            gear.Items.Add("Hide to tray", null, delegate { Close(); });
        }

        void OnCoreChanged() { SyncText(); RefreshView(); }

        // current scale factor, measured from a control we know the design size of (works before and after auto-scaling)
        int S(int v) { return (int)Math.Round(v * (face.Width / 104.0)); }

        // collapsed = only the ring clock; expanded = ring + status + description + buttons
        void ApplyLayout(bool col)
        {
            collapsed = col;
            bool show = !col;
            lblStatus.Visible = show; lblGear.Visible = show; lblClose.Visible = show; pnlTxt.Visible = show;
            bToggle.Visible = show; bFrom.Visible = show; bSlice.Visible = show; bLog.Visible = show;
            lblToggle.Text = col ? "»" : "«";       // » = expand, « = collapse
            face.Left = S(col ? 16 : 12);                      // centred in the narrow ring-only window
            lblToggle.Left = S(col ? 108 : 326);
            lblToggle.Top = S(col ? 2 : 10);
            ClientSize = new Size(S(col ? 136 : 400), S(136));
            lblToggle.BringToFront();
        }

        void SetCollapsed(bool col, bool persist)
        {
            ApplyLayout(col);
            if (Visible)       // keep it fully on screen after growing
            {
                Rectangle wa = Screen.FromControl(this).WorkingArea;
                int x = Left, y = Top;
                if (x + Width > wa.Right) x = wa.Right - Width - 8;
                if (y + Height > wa.Bottom) y = wa.Bottom - Height - 8;
                if (x != Left || y != Top) Location = new Point(Math.Max(wa.Left, x), Math.Max(wa.Top, y));
            }
            if (persist) { Core.Collapsed = col; Core.MarkSettingsDirty(); }
            RefreshView();
        }

        public void SyncText()
        {
            string d = Core.CurrentDesc;
            if (txt.Text != d) { sync = true; txt.Text = d; sync = false; }
        }

        public void RefreshView()
        {
            DateTime now = DateTime.Now;
            TimeSpan total = Core.DayTotal(now, now), target = Core.Target;
            bool run = Core.Open != null, over = total >= target;
            long secs = (long)total.TotalSeconds;
            face.Main = (secs / 3600).ToString("00") + ":" + ((secs / 60) % 60).ToString("00");
            face.Sec = (secs % 60).ToString("00");
            face.SecondOfMinute = (int)(secs % 60);
            face.Running = run;
            face.Progress = target.TotalSeconds > 0 ? total.TotalSeconds / target.TotalSeconds : 1;
            Color pc = Palette.For(face.Progress);
            face.Accent = run ? pc : Palette.Mix(pc, Grey, 0.45);    // paused = same colour, washed out
            face.Invalidate();

            if (!over)
            {
                TimeSpan left = target - total;
                string s = "Left " + Fmt.DurHmCeil(left);
                DateTime end;
                if (Core.TryFinish(now, out end))
                {
                    s += "  ·  ends " + Fmt.Hm(end) + (end.Date > now.Date ? " (+1d)" : "");
                    face.Note = "ends " + Fmt.Hm(end);
                }
                else { s += "  ·  paused"; face.Note = "paused"; }
                lblStatus.Text = s; lblStatus.ForeColor = Muted;
                face.NoteColor = Muted;
            }
            else
            {
                string ot = Fmt.DurHm(total - target);
                lblStatus.Text = "Done ✓  Overtime +" + ot + (run ? "" : "  ·  paused");
                lblStatus.ForeColor = Green;
                face.Note = "OT +" + ot;
                face.NoteColor = Green;
            }
            if (lblStatus.Text != lastTip) { lastTip = lblStatus.Text; tip.SetToolTip(face, lastTip); }   // hover the ring when collapsed
            if (run) { bToggle.Text = "Pause"; Tint(bToggle, Color.FromArgb(74, 84, 98), Light); }
            else { bToggle.Text = Core.HasEntries(now) ? "Resume" : "Start"; Tint(bToggle, Green, Color.FromArgb(12, 28, 20)); }
            bSlice.Enabled = run;
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            int d = (int)(24 * sc);
            IntPtr h = CreateRoundRectRgn(0, 0, Width + 1, Height + 1, d, d);
            Region = Region.FromHrgn(h);
            DeleteObject(h);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
                if (!hidTold) { hidTold = true; if (HidFirstTime != null) HidFirstTime(); }
            }
            base.OnFormClosing(e);
        }

        protected override void OnMove(EventArgs e)
        {
            base.OnMove(e);
            if (Visible && WindowState == FormWindowState.Normal) { Core.WinX = Left; Core.WinY = Top; Core.MarkSettingsDirty(); }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            HotkeyOk = RegisterHotKey(Handle, 1, 0x0001 | 0x0002 | 0x4000, (uint)Keys.P);   // Ctrl+Alt+P
            SendMessage(txt.Handle, 0x1501, (IntPtr)1, "What are you working on?");
        }
        protected override void OnHandleDestroyed(EventArgs e)
        {
            UnregisterHotKey(Handle, 1);
            base.OnHandleDestroyed(e);
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == 1) { if (ToggleRequested != null) ToggleRequested(); return; }
            base.WndProc(ref m);
        }
    }

    // ------------------------------------------------------------------ automatic updates (GitHub Releases)
    // Every push to the repository is built + published as a release by GitHub Actions. The app checks the
    // latest release, downloads WorkTimer.exe, verifies its SHA-256 and version, swaps itself and restarts.
    static class Updater
    {
        public const string Repo = "ahmed-abokhatwa/Timer";

        public class Release { public Version Version; public string ExeUrl, ShaUrl; }

        class TimedClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest r = base.GetWebRequest(address);
                r.Timeout = 30000;
                return r;
            }
        }

        // For testing the update flow against a local fake server (only used when this environment variable is set).
        static string TestBase { get { string b = Environment.GetEnvironmentVariable("WORKTIMER_UPDATE_TEST_BASE"); return string.IsNullOrEmpty(b) ? null : b.TrimEnd('/'); } }
        static string ApiUrl { get { return TestBase != null ? TestBase + "/latest.json" : "https://api.github.com/repos/" + Repo + "/releases/latest"; } }
        static string DownloadPrefix { get { return TestBase != null ? TestBase + "/" : "https://github.com/" + Repo + "/releases/download/"; } }

        public static Version Current { get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version; } }
        public static string CurrentText { get { Version v = Current; return v.Major + "." + v.Minor + "." + v.Build; } }

        // automatic checks can be switched off for tests / portable use with WORKTIMER_NO_UPDATE=1
        public static bool EnvDisabled { get { return !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WORKTIMER_NO_UPDATE")); } }

        public static Version ParseVersion(string s)
        {
            if (s == null) return null;
            Match m = Regex.Match(s.Trim(), @"^v?(\d+)\.(\d+)\.(\d+)(?:\.(\d+))?$");
            if (!m.Success) return null;
            return new Version(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value),
                               m.Groups[4].Success ? int.Parse(m.Groups[4].Value) : 0);
        }

        public static bool IsNewer(Version latest, Version current) { return latest != null && current != null && latest > current; }

        // pulls the tag + the two asset links out of GitHub's "latest release" JSON; only links inside this repo's releases are accepted
        public static Release ParseRelease(string json, string repo)
        {
            if (string.IsNullOrEmpty(json)) return null;
            Match t = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
            if (!t.Success) return null;
            Version v = ParseVersion(t.Groups[1].Value);
            if (v == null) return null;
            string prefix = Regex.Escape(repo == Repo ? DownloadPrefix : "https://github.com/" + repo + "/releases/download/");
            Match e = Regex.Match(json, "\"browser_download_url\"\\s*:\\s*\"(" + prefix + "[^\"]+/WorkTimer\\.exe)\"", RegexOptions.IgnoreCase);
            Match h = Regex.Match(json, "\"browser_download_url\"\\s*:\\s*\"(" + prefix + "[^\"]+/WorkTimer\\.exe\\.sha256)\"", RegexOptions.IgnoreCase);
            if (!e.Success || !h.Success) return null;
            return new Release { Version = v, ExeUrl = e.Groups[1].Value, ShaUrl = h.Groups[1].Value };
        }

        static WebClient NewClient()
        {
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;      // TLS 1.2 (HTTPS only; the test server is plain http on localhost)
            WebClient wc = new TimedClient();
            wc.Headers[HttpRequestHeader.UserAgent] = "WorkTimer-Updater";
            wc.Headers[HttpRequestHeader.Accept] = "application/vnd.github+json";
            return wc;
        }

        public static Release FetchLatest(out string error)
        {
            error = null;
            try
            {
                using (WebClient wc = NewClient())
                {
                    Release r = ParseRelease(wc.DownloadString(ApiUrl), Repo);
                    if (r == null) error = "no usable release found";
                    return r;
                }
            }
            catch (Exception ex) { error = ex.Message; return null; }
        }

        public static string Sha256Hex(byte[] bytes)
        {
            using (SHA256 h = SHA256.Create()) return BitConverter.ToString(h.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }

        // downloads next to the running EXE (so replacing it is a plain rename) and verifies checksum, EXE header and version
        public static string Download(Release r, out string error)
        {
            error = null;
            string dest = Application.ExecutablePath + ".new";
            try
            {
                using (WebClient wc = NewClient())
                {
                    Match m = Regex.Match(wc.DownloadString(r.ShaUrl).Trim(), "^[0-9a-fA-F]{64}");
                    if (!m.Success) { error = "bad checksum file"; return null; }
                    if (File.Exists(dest)) File.Delete(dest);
                    wc.DownloadFile(r.ExeUrl, dest);
                    byte[] b = File.ReadAllBytes(dest);
                    if (b.Length < 20000 || b[0] != (byte)'M' || b[1] != (byte)'Z') { error = "downloaded file is not a program"; File.Delete(dest); return null; }
                    if (Sha256Hex(b) != m.Value.ToLowerInvariant()) { error = "checksum mismatch"; File.Delete(dest); return null; }
                    Version fv = ParseVersion(System.Diagnostics.FileVersionInfo.GetVersionInfo(dest).FileVersion);
                    if (fv == null || fv != r.Version) { error = "version mismatch"; File.Delete(dest); return null; }
                }
                return dest;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                try { if (File.Exists(dest)) File.Delete(dest); } catch { }
                return null;
            }
        }

        // a running EXE can be renamed (not overwritten): move it aside, put the new one in its place
        public static bool Swap(string newPath)
        {
            string cur = Application.ExecutablePath, old = cur + ".old";
            try { if (File.Exists(old)) File.Delete(old); File.Move(cur, old); }
            catch { return false; }
            try { File.Move(newPath, cur); }
            catch { try { File.Move(old, cur); } catch { } return false; }
            return true;
        }

        public static void CleanupLeftovers()
        {
            try { File.Delete(Application.ExecutablePath + ".old"); } catch { }
            try { File.Delete(Application.ExecutablePath + ".new"); } catch { }
        }
    }

    // ------------------------------------------------------------------ tray + glue
    class TrayApp : ApplicationContext
    {
        NotifyIcon ni; MiniForm mini; LogForm log; System.Windows.Forms.Timer timer;
        ToolStripMenuItem miToggle, miSlice, miAuto, miAutoUpd;
        System.Windows.Forms.Timer updTimer;
        bool updating;
        Icon icoPause, lastIcon;
        Dictionary<int, Icon> runIcons = new Dictionary<int, Icon>();     // one play-icon per colour step
        string lastTip = "";
        DateTime notifiedDay = DateTime.MinValue;
        bool sleepPaused, exiting;

        [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr h);

        static Icon MakeIcon(Color c, bool pause)
        {
            using (Bitmap bmp = new Bitmap(32, 32))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (SolidBrush b = new SolidBrush(c)) g.FillEllipse(b, 1, 1, 30, 30);
                using (SolidBrush w = new SolidBrush(Color.White))
                {
                    if (!pause) g.FillPolygon(w, new Point[] { new Point(12, 8), new Point(12, 24), new Point(24, 16) });
                    else { g.FillRectangle(w, 10, 8, 4, 16); g.FillRectangle(w, 18, 8, 4, 16); }
                }
                IntPtr h = bmp.GetHicon();
                try { using (Icon tmp = Icon.FromHandle(h)) return (Icon)tmp.Clone(); }
                finally { DestroyIcon(h); }
            }
        }

        public TrayApp()
        {
            Core.Load();
            if (Core.FirstRun) Core.ApplyAutoStart();
            else if (Core.AutoStart) Core.ApplyAutoStart();      // keeps the path right if the EXE was moved

            icoPause = MakeIcon(Color.FromArgb(130, 130, 130), true);

            ContextMenuStrip menu = new ContextMenuStrip();
            miToggle = new ToolStripMenuItem("Pause", null, delegate { Toggle(); });
            miToggle.Font = new Font(miToggle.Font, FontStyle.Bold);
            miSlice = new ToolStripMenuItem("Slice \u2013 start next task", null, delegate { Core.Slice(); });
            menu.Items.Add(miToggle);
            menu.Items.Add(miSlice);
            menu.Items.Add("Start from a chosen time…", null, delegate { StartFrom(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Show timer", null, delegate { ShowMini(); });
            menu.Items.Add("Open log / export week", null, delegate { ShowLog(); });
            menu.Items.Add(new ToolStripSeparator());
            miAuto = new ToolStripMenuItem("Start with Windows", null, delegate { Core.SetAutoStart(!Core.AutoStart); });
            menu.Items.Add(miAuto);
            miAutoUpd = new ToolStripMenuItem("Update automatically", null, delegate { Core.AutoUpdate = !Core.AutoUpdate; Core.SaveSettings(); });
            menu.Items.Add(miAutoUpd);
            menu.Items.Add("Check for updates now", null, delegate { CheckForUpdates(true); });
            menu.Items.Add(new ToolStripSeparator());
            ToolStripMenuItem miVer = new ToolStripMenuItem("Work Timer v" + Updater.CurrentText) { Enabled = false };
            menu.Items.Add(miVer);
            menu.Items.Add("Exit", null, delegate { ExitApp(); });
            menu.Opening += delegate { UpdateMenu(); };

            ni = new NotifyIcon();
            ni.ContextMenuStrip = menu;
            ni.Icon = icoPause; lastIcon = icoPause;
            ni.Text = "Work Timer";
            ni.MouseClick += delegate(object s, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left) { if (mini.Visible) mini.Hide(); else ShowMini(); }
            };
            ni.Visible = true;

            mini = new MiniForm();
            mini.ToggleRequested += Toggle;
            mini.LogRequested += ShowLog;
            mini.FromRequested += StartFrom;
            mini.HidFirstTime += delegate { ni.ShowBalloonTip(3000, "Work Timer", "Still running in the tray. Click the tray icon to bring it back (Ctrl+Alt+P pauses/resumes).", ToolTipIcon.Info); };
            mini.Quiet = true; mini.Show(); mini.Quiet = false;

            DateTime now = DateTime.Now;
            if (Core.DayTotal(now, now) >= Core.Target) notifiedDay = now.Date;

            timer = new System.Windows.Forms.Timer { Interval = 1000 };
            timer.Tick += delegate { OnTick(); };
            timer.Start();
            Core.Changed += delegate { UpdateTray(); };
            UpdateTray();

            SystemEvents.PowerModeChanged += OnPower;
            SystemEvents.SessionEnding += OnSessionEnding;

            // look for a new version shortly after start, then every 6 hours
            updTimer = new System.Windows.Forms.Timer { Interval = 30000 };
            updTimer.Tick += delegate
            {
                updTimer.Interval = 6 * 3600 * 1000;
                if (Core.AutoUpdate && !Updater.EnvDisabled) CheckForUpdates(false);
            };
            updTimer.Start();

            if (Core.RecoveredAt.HasValue)
                ni.ShowBalloonTip(6000, "Work Timer", "The timer was running when the app closed. It was stopped at " + Fmt.Hm(Core.RecoveredAt.Value) + ". Press Resume to continue.", ToolTipIcon.Warning);
            else if (Core.FirstRun)
                ni.ShowBalloonTip(6000, "Work Timer", "Ready. It lives in the tray and starts with Windows. Press Start when you begin work.", ToolTipIcon.Info);
        }

        void Toggle() { if (Core.Open != null) Core.Pause(); else Core.Start(); }

        // ---- automatic updates
        void UiDo(Action a)
        {
            try { if (mini != null && mini.IsHandleCreated && !mini.IsDisposed) mini.BeginInvoke(a); } catch { }
        }

        void CheckForUpdates(bool manual)
        {
            if (updating) return;
            updating = true;
            ThreadPool.QueueUserWorkItem(delegate
            {
                string err = null, dlErr = null, file = null;
                Updater.Release r = Updater.FetchLatest(out err);
                bool newer = r != null && Updater.IsNewer(r.Version, Updater.Current);
                if (newer) file = Updater.Download(r, out dlErr);
                UiDo(delegate
                {
                    updating = false;
                    if (r == null) { if (manual) MessageBox.Show("Could not check for updates.\n\n" + err, "Work Timer"); return; }
                    if (!newer) { if (manual) MessageBox.Show("You have the latest version (v" + Updater.CurrentText + ").", "Work Timer"); return; }
                    if (file == null)
                    {
                        ni.ShowBalloonTip(8000, "Work Timer", "Version " + r.Version.Major + "." + r.Version.Minor + "." + r.Version.Build +
                            " is available but could not be installed automatically (" + dlErr + "). Get it from github.com/" + Updater.Repo + "/releases", ToolTipIcon.Info);
                        return;
                    }
                    ApplyUpdate(file);
                });
            });
        }

        void ApplyUpdate(string newFile)
        {
            if (exiting) return;
            if (!Updater.Swap(newFile))
            {
                try { File.Delete(newFile); } catch { }
                ni.ShowBalloonTip(8000, "Work Timer", "An update is available but this folder is read-only. Move WorkTimer.exe to a normal folder (e.g. Desktop or Documents).", ToolTipIcon.Warning);
                return;
            }
            bool wasRunning = Core.Open != null;
            try
            {
                // the new copy waits for this process to finish, then continues the running task
                Process.Start(new ProcessStartInfo(Application.ExecutablePath,
                    "--wait-pid " + Process.GetCurrentProcess().Id + " --updated" + (wasRunning ? " --resume" : "")) { UseShellExecute = false });
            }
            catch (Exception ex) { Core.LogError(ex); return; }     // swapped on disk; it will be used at the next start
            exiting = true;
            try
            {
                timer.Stop(); updTimer.Stop();
                SystemEvents.PowerModeChanged -= OnPower;
                SystemEvents.SessionEnding -= OnSessionEnding;
                Core.Pause(); Core.SaveAll(); Core.SaveSettings();
            }
            catch (Exception ex) { Core.LogError(ex); }
            ni.Visible = false; ni.Dispose();
            if (log != null && !log.IsDisposed) log.Close();
            mini.Dispose();
            ExitThread();
        }

        // called once after start-up (arguments come from an update restart)
        public void AfterStart(bool resume, bool updated)
        {
            if (resume) Core.ResumeLast();
            if (updated) ni.ShowBalloonTip(4000, "Work Timer", "Updated to v" + Updater.CurrentText + (resume ? " – timer is still running." : "."), ToolTipIcon.Info);
        }

        void StartFrom()
        {
            Entry o = Core.Open;
            DateTime init = o != null ? o.Start : DateTime.Now;
            using (StartFromDialog dlg = new StartFromDialog(init, o != null))
            {
                if (dlg.ShowDialog() == DialogResult.OK) Core.StartFrom(dlg.Res);
            }
        }

        public void RequestShow()
        {
            try { if (mini != null && mini.IsHandleCreated) mini.BeginInvoke(new Action(ShowMini)); } catch { }
        }

        void ShowMini()
        {
            if (!mini.Visible) mini.Show();
            if (mini.WindowState == FormWindowState.Minimized) mini.WindowState = FormWindowState.Normal;
            mini.TopMost = Core.OnTop;
            mini.SyncText(); mini.RefreshView();
            mini.Activate();
        }

        void ShowLog()
        {
            if (log == null || log.IsDisposed) { log = new LogForm(); log.Show(); }
            else
            {
                if (!log.Visible) log.Show();
                if (log.WindowState == FormWindowState.Minimized) log.WindowState = FormWindowState.Normal;
            }
            log.GoToday();
            log.Activate();
        }

        void UpdateMenu()
        {
            bool run = Core.Open != null;
            miToggle.Text = run ? "Pause" : (Core.HasEntries(DateTime.Now) ? "Resume" : "Start");
            miToggle.ShortcutKeyDisplayString = mini.HotkeyOk ? "Ctrl+Alt+P" : "";
            miSlice.Enabled = run;
            miAuto.Checked = Core.AutoStart;
            miAutoUpd.Checked = Core.AutoUpdate;
        }

        void OnTick()
        {
            try
            {
                Core.Tick();
                UpdateTray();
                if (mini.Visible) mini.RefreshView();
                if (log != null && !log.IsDisposed && log.Visible) log.Tick();

                DateTime now = DateTime.Now;
                TimeSpan total = Core.DayTotal(now, now);
                if (total < Core.Target) notifiedDay = DateTime.MinValue;
                else if (Core.Open != null && notifiedDay != now.Date && Core.Target > TimeSpan.Zero)
                {
                    notifiedDay = now.Date;
                    ni.ShowBalloonTip(5000, "Work Timer", "Your " + Fmt.DurHm(Core.Target) + " are done – shift complete! Everything from now on is overtime.", ToolTipIcon.Info);
                }
            }
            catch (Exception ex) { Core.LogError(ex); }
        }

        void UpdateTray()
        {
            DateTime now = DateTime.Now;
            TimeSpan total = Core.DayTotal(now, now);
            bool run = Core.Open != null;
            Icon want = icoPause;
            if (run)
            {
                double p = Core.Target.TotalSeconds > 0 ? total.TotalSeconds / Core.Target.TotalSeconds : 1;
                int step = p >= 1 ? 1 : 0;                         // 0 = red (still working towards the target), 1 = green (overtime)
                if (!runIcons.TryGetValue(step, out want))
                {
                    want = MakeIcon(step == 1 ? Palette.Green : Palette.Red, false);
                    runIcons[step] = want;
                }
            }
            if (want != lastIcon) { ni.Icon = want; lastIcon = want; }
            string tip = "Work Timer \u00B7 " + (run ? "running" : "paused") + "\n" + Fmt.DurHm(total) + " / " + Fmt.DurHm(Core.Target);
            DateTime fin;
            if (Core.TryFinish(now, out fin)) tip += " \u00B7 ends " + Fmt.Hm(fin);
            else if (total >= Core.Target) tip += " \u00B7 OT +" + Fmt.DurHm(total - Core.Target);
            if (tip != lastTip)
            {
                lastTip = tip;
                ni.Text = tip.Length > 63 ? tip.Substring(0, 63) : tip;
            }
        }

        void OnPower(object sender, PowerModeChangedEventArgs e)
        {
            try
            {
                if (e.Mode == PowerModes.Suspend)
                {
                    DateTime t = DateTime.Now;
                    if (mini.IsHandleCreated) mini.BeginInvoke(new Action(delegate { if (Core.Open != null) { Core.PauseAt(t); sleepPaused = true; } }));
                }
                else if (e.Mode == PowerModes.Resume && mini.IsHandleCreated)
                {
                    mini.BeginInvoke(new Action(delegate
                    {
                        if (sleepPaused)
                        {
                            sleepPaused = false;
                            ni.ShowBalloonTip(5000, "Work Timer", "The timer was paused when the PC went to sleep. Press Resume to continue.", ToolTipIcon.Info);
                        }
                    }));
                }
            }
            catch (Exception ex) { Core.LogError(ex); }
        }

        void OnSessionEnding(object sender, SessionEndingEventArgs e)
        {
            try
            {
                DateTime t = DateTime.Now;
                if (mini.IsHandleCreated) mini.BeginInvoke(new Action(delegate { Core.PauseAt(t); Core.SaveSettings(); }));
            }
            catch (Exception ex) { Core.LogError(ex); }
        }

        void ExitApp()
        {
            if (exiting) return;
            string msg = Core.Open != null ? "Exit Work Timer?\n\nThe timer will be paused and saved." : "Exit Work Timer?";
            if (MessageBox.Show(msg, "Work Timer", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            exiting = true;
            try
            {
                timer.Stop();
                SystemEvents.PowerModeChanged -= OnPower;
                SystemEvents.SessionEnding -= OnSessionEnding;
                Core.Pause();
                Core.SaveAll();
                Core.SaveSettings();
            }
            catch (Exception ex) { Core.LogError(ex); }
            ni.Visible = false;
            ni.Dispose();
            if (log != null && !log.IsDisposed) log.Close();
            mini.Dispose();
            ExitThread();
        }
    }

    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Core.Init();
            // arguments used by the automatic updater: --wait-pid N  --resume  --updated
            bool resume = false, updated = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--resume") resume = true;
                else if (args[i] == "--updated") updated = true;
                else if (args[i] == "--wait-pid" && i + 1 < args.Length)
                {
                    int pid;
                    if (int.TryParse(args[++i], out pid)) { try { Process.GetProcessById(pid).WaitForExit(20000); } catch { } }
                }
            }
            if (updated) Updater.CleanupLeftovers();
            string key = "WorkTimer_" + Core.Dir.Replace('\\', '_').Replace(':', '_').Replace('/', '_');
            bool first;
            Mutex mtx = new Mutex(true, key, out first);
            EventWaitHandle ev = new EventWaitHandle(false, EventResetMode.AutoReset, key + "_show");
            if (!first) { ev.Set(); return; }       // already running: just bring its window forward

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object s, ThreadExceptionEventArgs e) { Core.LogError(e.Exception); };
            AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e) { Core.LogError(e.ExceptionObject as Exception); };

            TrayApp app = new TrayApp();
            if (resume || updated) app.AfterStart(resume, updated);
            Thread t = new Thread(delegate() { while (true) { ev.WaitOne(); app.RequestShow(); } });
            t.IsBackground = true;
            t.Start();
            Application.Run(app);
            GC.KeepAlive(mtx);
        }
    }
}
