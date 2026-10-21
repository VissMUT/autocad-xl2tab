using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;

[assembly: CommandClass(typeof(Xl2Tab.Plugin))]
[assembly: ExtensionApplication(typeof(Xl2Tab.Plugin))]

namespace Xl2Tab
{
    public class Plugin : IExtensionApplication
    {
        static object Miss { get { return Type.Missing; } }

        static bool _toolbarReady;

        public void Initialize()
        {
            try
            {
                if (Autodesk.Windows.ComponentManager.Ribbon == null)
                    Autodesk.Windows.ComponentManager.ItemInitialized += OnItemInitialized;
                else
                    AddRibbonTab();
            }
            catch (System.Exception ex) { Log("Инициализация: " + ex.Message); }
            try { Autodesk.AutoCAD.ApplicationServices.Application.Idle += OnIdle; } catch { }
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.DocumentActivated += OnDocActivated;
            }
            catch { }
            EnsureClassicToolbar("старт");
        }

        static void OnIdle(object sender, EventArgs e)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnIdle;
            }
            catch { }
            EnsureClassicToolbar("idle");
        }

        static void OnDocActivated(object sender, Autodesk.AutoCAD.ApplicationServices.DocumentCollectionEventArgs e)
        {
            try
            {
                if (_toolbarReady)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.DocumentActivated -= OnDocActivated;
                    return;
                }
                EnsureClassicToolbar("документ");
            }
            catch { }
        }

        public void Terminate() { }

        static object NewExcel()
        {
            Type t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null) throw new InvalidOperationException("Microsoft Excel не найден в системе");
            return Activator.CreateInstance(t);
        }

        static int[] GetExcelPids()
        {
            try { return Array.ConvertAll(Process.GetProcessesByName("EXCEL"), p => p.Id); }
            catch { return new int[0]; }
        }

        // жёстко убивает Excel-процессы, появившиеся после снятия before
        // (Quit по COM не всегда гасит процесс - остаются невидимые сироты)
        static void KillNewExcel(int[] before)
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName("EXCEL"))
                    if (Array.IndexOf(before, p.Id) < 0)
                        try { p.Kill(); } catch { }
            }
            catch { }
        }
        static object GetP(object o, string n) { return o.GetType().InvokeMember(n, BindingFlags.GetProperty, null, o, null); }
        static object GetPI(object o, string n, object[] a) { return o.GetType().InvokeMember(n, BindingFlags.GetProperty, null, o, a); }
        static void SetP(object o, string n, object v) { o.GetType().InvokeMember(n, BindingFlags.SetProperty, null, o, new object[] { v }); }
        static object Call(object o, string n, object[] a) { return o.GetType().InvokeMember(n, BindingFlags.InvokeMethod, null, o, a); }

        static string LogPath { get { return Path.Combine(Path.GetTempPath(), "xltabdiag_log.txt"); } }

        static void Log(string m)
        {
            try { File.AppendAllText(LogPath, m + "\r\n", Encoding.UTF8); } catch { }
            Document d = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (d != null) d.Editor.WriteMessage("\n" + m);
        }

        // номер колонки - буквы (1-A, 26-Z, 27-AA)
        static string ColLetter(int n)
        {
            string s = "";
            while (n > 0)
            {
                int r = (n - 1) % 26;
                s = (char)('A' + r) + s;
                n = (n - 1) / 26;
            }
            return s;
        }

        // буквы колонки - номер (A-1, AA-27)
        static int ColNo(string letters)
        {
            int n = 0;
            foreach (char ch in letters) n = n * 26 + (ch - 'A' + 1);
            return n;
        }

        // "A5:N40" - (c1,r1,c2,r2)
        static bool ParseRange(string rng, out int c1, out int r1, out int c2, out int r2)
        {
            c1 = c2 = r1 = r2 = 0;
            if (string.IsNullOrEmpty(rng)) return false;
            string s = rng.Trim().ToUpperInvariant();
            int p = s.IndexOf(':');
            if (p <= 0 || p == s.Length - 1) return false;
            string a = s.Substring(0, p), b = s.Substring(p + 1);
            int i1 = 0;
            while (i1 < a.Length && char.IsLetter(a[i1])) i1++;
            if (i1 == 0 || i1 == a.Length) return false;
            if (!int.TryParse(a.Substring(i1), out r1)) return false;
            c1 = ColNo(a.Substring(0, i1));
            int i2 = 0;
            while (i2 < b.Length && char.IsLetter(b[i2])) i2++;
            if (i2 == 0 || i2 == b.Length) return false;
            if (!int.TryParse(b.Substring(i2), out r2)) return false;
            c2 = ColNo(b.Substring(0, i2));
            return c1 > 0 && c2 > 0 && r1 > 0 && r2 > 0 && c2 >= c1 && r2 >= r1;
        }

        // ── чтение xlsx напрямую из ZIP (без Excel) ──────────────────

        class XlsxSheet
        {
            public string Name;
            public bool Hidden;
            public int C1 = 1, R1 = 1, C2 = 1, R2 = 1;                  // dimension (1-based)
            public Dictionary<int, double> ColW = new Dictionary<int, double>(); // абс. колонка - символы
            public Dictionary<int, double> RowH = new Dictionary<int, double>(); // абс. строка - пункты
            public double DefaultColW = 8.43;
            public double DefaultRowH = 15.0;
        }

        static void ParseSheet(XmlDocument doc, XlsxSheet sh)
        {
            XmlNodeList dims = doc.GetElementsByTagName("dimension");
            if (dims.Count > 0)
            {
                string r = ((XmlElement)dims[0]).GetAttribute("ref");
                int cc1, rr1, cc2, rr2;
                if (!string.IsNullOrEmpty(r) && r != "A1" &&
                    ParseRange(r, out cc1, out rr1, out cc2, out rr2))
                {
                    sh.C1 = cc1; sh.R1 = rr1; sh.C2 = cc2; sh.R2 = rr2;
                }
            }
            XmlNodeList sfps = doc.GetElementsByTagName("sheetFormatPr");
            if (sfps.Count > 0)
            {
                XmlElement e = sfps[0] as XmlElement;
                double d;
                if (double.TryParse(e.GetAttribute("defaultColWidth"), NumberStyles.Any,
                                    CultureInfo.InvariantCulture, out d) && d > 0) sh.DefaultColW = d;
                if (double.TryParse(e.GetAttribute("defaultRowHeight"), NumberStyles.Any,
                                    CultureInfo.InvariantCulture, out d) && d > 0) sh.DefaultRowH = d;
            }
            foreach (XmlNode n in doc.GetElementsByTagName("col"))
            {
                XmlElement e = n as XmlElement;
                if (e == null) continue;
                int min, max;
                double w;
                if (!int.TryParse(e.GetAttribute("min"), out min)) continue;
                if (!int.TryParse(e.GetAttribute("max"), out max)) max = min;
                if (!double.TryParse(e.GetAttribute("width"), NumberStyles.Any,
                                     CultureInfo.InvariantCulture, out w) || w <= 0) continue;
                for (int c = min; c <= max; c++) sh.ColW[c] = w;
            }
            foreach (XmlNode n in doc.GetElementsByTagName("row"))
            {
                XmlElement e = n as XmlElement;
                if (e == null) continue;
                int r;
                if (!int.TryParse(e.GetAttribute("r"), out r)) continue;
                double h;
                if (double.TryParse(e.GetAttribute("ht"), NumberStyles.Any,
                                    CultureInfo.InvariantCulture, out h) && h > 0)
                    sh.RowH[r] = h;
            }
        }

        // ключ: порядковый номер листа (1-based). Бросает исключение при нечитаемом файле.
        static Dictionary<int, XlsxSheet> ReadXlsx(string path)
        {
            Dictionary<int, XlsxSheet> sheets = new Dictionary<int, XlsxSheet>();
            using (ZipArchive zip = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry == null)
                {
                    foreach (ZipArchiveEntry z in zip.Entries)
                        if (z.FullName.ToLowerInvariant() == "xl/workbook.xml") { wbEntry = z; break; }
                }
                if (wbEntry == null) throw new InvalidOperationException("в архиве нет xl/workbook.xml");

                XmlDocument wbDoc = new XmlDocument();
                using (Stream s = wbEntry.Open()) wbDoc.Load(s);

                List<string> names = new List<string>();
                List<string> rids = new List<string>();
                List<bool> hidden = new List<bool>();
                foreach (XmlNode n in wbDoc.GetElementsByTagName("sheet"))
                {
                    XmlElement e = n as XmlElement;
                    if (e == null) continue;
                    names.Add(e.GetAttribute("name"));
                    string state = e.GetAttribute("state");
                    hidden.Add(state == "hidden" || state == "veryHidden");
                    string rid = null;
                    foreach (XmlAttribute a in e.Attributes)
                        if (a.LocalName == "id") { rid = a.Value; break; }
                    rids.Add(rid);
                }

                Dictionary<string, string> rels = new Dictionary<string, string>();
                ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry != null)
                {
                    XmlDocument relDoc = new XmlDocument();
                    using (Stream s = relsEntry.Open()) relDoc.Load(s);
                    foreach (XmlNode n in relDoc.GetElementsByTagName("Relationship"))
                    {
                        XmlElement e = n as XmlElement;
                        if (e == null) continue;
                        string id = null, target = null;
                        foreach (XmlAttribute a in e.Attributes)
                        {
                            if (a.LocalName == "Id") id = a.Value;
                            else if (a.LocalName == "Target") target = a.Value;
                        }
                        if (id != null && target != null) rels[id] = target;
                    }
                }

                for (int i = 0; i < names.Count; i++)
                {
                    XlsxSheet sh = new XlsxSheet();
                    sh.Name = names[i];
                    sh.Hidden = hidden[i];
                    string entryPath = null;
                    if (rids[i] != null && rels.ContainsKey(rids[i]))
                    {
                        entryPath = rels[rids[i]].TrimStart('/');
                        if (!entryPath.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                            entryPath = "xl/" + entryPath;
                    }
                    if (entryPath == null) entryPath = "xl/worksheets/sheet" + (i + 1) + ".xml";
                    ZipArchiveEntry shEntry = zip.GetEntry(entryPath);
                    if (shEntry == null)
                    {
                        foreach (ZipArchiveEntry z in zip.Entries)
                            if (z.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase)) { shEntry = z; break; }
                    }
                    if (shEntry != null)
                    {
                        XmlDocument doc = new XmlDocument();
                        using (Stream s = shEntry.Open()) doc.Load(s);
                        ParseSheet(doc, sh);
                    }
                    sheets[i + 1] = sh;
                }
            }
            return sheets;
        }

        static double RowHmmZip(XlsxSheet sh, int row, double scale)
        {
            double h = sh.DefaultRowH;
            if (sh.RowH.ContainsKey(row)) h = sh.RowH[row];
            if (h <= 0) h = sh.DefaultRowH;
            double mm = h * 25.4 / 72.0 * scale;
            return mm < 5.0 ? 5.0 : mm;
        }

        static double ColWmmZip(XlsxSheet sh, int col, double scale)
        {
            double w = sh.DefaultColW;
            if (sh.ColW.ContainsKey(col)) w = sh.ColW[col];
            if (w <= 0) w = sh.DefaultColW;
            double px = w > 1.0 ? w * 7.0 + 5.0 : w * 12.0;
            double mm = px * 25.4 / 96.0 * scale;
            return mm < 5.0 ? 5.0 : mm;
        }

        // обновление одной связи (вместо глобального dlm.Update - не растёт квадратично).
        // UpdateOption.None: тянем только данные - высоты/ширины применяем сами
        // (после масштаба), поэтому флаги UpdateRowHeight/UpdateColumnWidth не нужны.
        static void UpdateLink(Transaction tr, ObjectId id)
        {
            DataLink dl = tr.GetObject(id, OpenMode.ForWrite) as DataLink;
            if (dl != null)
                dl.Update(UpdateDirection.SourceToData, UpdateOption.None);
        }

        // высота строки Excel (пункты) - мм с масштабом, мин 5
        static double RowHmm(object ws, int row, double scale)
        {
            double h = 15.0;
            try { h = (double)GetP(GetPI(GetP(ws, "Rows"), "Item", new object[] { row }), "RowHeight"); } catch { }
            if (h <= 0) h = 15.0;
            double mm = h * 25.4 / 72.0 * scale;
            return mm < 5.0 ? 5.0 : mm;
        }

        // ширина колонки Excel (символы) - мм с масштабом, мин 5
        static double ColWmm(object ws, int col, double scale)
        {
            double w = 8.43;
            try { w = (double)GetP(GetPI(GetP(ws, "Columns"), "Item", new object[] { col }), "ColumnWidth"); } catch { }
            if (w <= 0) w = 8.43;
            double px = w > 1.0 ? w * 7.0 + 5.0 : w * 12.0;
            double mm = px * 25.4 / 96.0 * scale;
            return mm < 5.0 ? 5.0 : mm;
        }

        // текст ячейки таблицы AutoCAD
        static string CellStr(Table t, int row, int col)
        {
            try { return t.GetTextString(row, col, 0); }
            catch { return "<err>"; }
        }

        static DataLink MakeLink(string name, string adapter, string conn, string desc)
        {
            DataLink dl = new DataLink();
            dl.Name = name;
            dl.Description = desc;
            dl.DataAdapterId = adapter;
            dl.ConnectionString = conn;
            dl.DataLinkOption = DataLinkOption.PersistCache;
            return dl;
        }

        // современный формат (как пишет диалог DATALINK). НЕ ИСПОЛЬЗУЕТСЯ:
        // через API даёт eDataLinkSourceNotFound на AutoCAD 2019 (проверено диагностикой)
        static ObjectId CreateDataLinkModern(Database db, Transaction tr, string name, string xlsxPath)
        {
            DataLink dl = MakeLink(name, "ACEXCEL",
                "ACEXCEL;DataSource=" + xlsxPath + ";SourceType=ExcelLink",
                "XL2TAB: " + xlsxPath);
            ObjectId id = db.DataLinkManager.AddDataLink(dl);
            tr.AddNewlyCreatedDBObject(dl, true);
            return id;
        }

        // рабочий формат: адаптер ACEXCEL + синтаксис <путь>!<лист>!<диапазон>
        // (подтверждено диагностикой TF: работает на AutoCAD 2019, идёт через ACE-провайдер)
        static ObjectId CreateDataLinkLegacy(Database db, Transaction tr, string name, string xlsxPath,
                                             string sheetName, string range)
        {
            DataLink dl = MakeLink(name, "ACEXCEL",
                xlsxPath + "!" + sheetName + "!" + range,
                "XL2TAB: " + xlsxPath);
            ObjectId id = db.DataLinkManager.AddDataLink(dl);
            tr.AddNewlyCreatedDBObject(dl, true);
            return id;
        }

        static Table AddTable(Database db, Transaction tr, BlockTableRecord btr,
                              ObjectId dlId, Point3d pt, int rows, int cols)
        {
            return AddTableAnchor(db, tr, btr, dlId, pt, rows, cols, 0);
        }

        // anchorRow: строка таблицы, с которой стартуют данные связи
        // (0 = вся таблица данные; 1 = строка 0 - ручная шапка)
        static Table AddTableAnchor(Database db, Transaction tr, BlockTableRecord btr,
                                    ObjectId dlId, Point3d pt, int rows, int cols, int anchorRow)
        {
            Table t = new Table();
            t.TableStyle = db.Tablestyle;
            t.Position = pt;
            t.NumRows = rows;
            t.NumColumns = cols;
            btr.AppendEntity(t);
            tr.AddNewlyCreatedDBObject(t, true);
            t.SetDataLink(anchorRow, 0, dlId, true);
            return t;
        }

        // две связи в одной таблице: шапка (строка 0) + данные (со строки 1)
        static Table AddTable2Links(Database db, Transaction tr, BlockTableRecord btr,
                                    ObjectId hdrId, ObjectId dataId, Point3d pt, int rows, int cols)
        {
            Table t = new Table();
            t.TableStyle = db.Tablestyle;
            t.Position = pt;
            t.NumRows = rows;
            t.NumColumns = cols;
            btr.AppendEntity(t);
            tr.AddNewlyCreatedDBObject(t, true);
            t.SetDataLink(0, 0, hdrId, true);
            t.SetDataLink(1, 0, dataId, true);
            return t;
        }

        static string MakeTestXlsx(string path)
        {
            object app = null;
            int[] pidsBefore = GetExcelPids();
            try
            {
                app = NewExcel();
                SetP(app, "Visible", 0);
                SetP(app, "DisplayAlerts", false);
                object wbs = GetP(app, "Workbooks");
                object wb = Call(wbs, "Add", new object[] { Miss });
                object ws1 = GetP(wb, "ActiveSheet");
                SetP(ws1, "Name", "DiagnList1");
                object r = GetPI(ws1, "Range", new object[] { "A1" });
                SetP(r, "Value2", "S1_HEADER");
                r = GetPI(ws1, "Range", new object[] { "A2" });
                SetP(r, "Value2", "S1_dataA");
                r = GetPI(ws1, "Range", new object[] { "B2" });
                SetP(r, "Value2", "S1_dataB");
                object ws2 = Call(GetP(wb, "Worksheets"), "Add", new object[] { Miss, Miss, Miss, Miss });
                SetP(ws2, "Name", "DiagnList2");
                r = GetPI(ws2, "Range", new object[] { "A1" });
                SetP(r, "Value2", "S2_HEADER");
                r = GetPI(ws2, "Range", new object[] { "A2" });
                SetP(r, "Value2", "S2_dataA");
                object ws3 = Call(GetP(wb, "Worksheets"), "Add", new object[] { Miss, Miss, Miss, Miss });
                SetP(ws3, "Name", "Диаг_Лист3");
                r = GetPI(ws3, "Range", new object[] { "A1" });
                SetP(r, "Value2", "S3_HEADER");
                r = GetPI(ws3, "Range", new object[] { "A2" });
                SetP(r, "Value2", "S3_dataA");
                object ws4 = Call(GetP(wb, "Worksheets"), "Add", new object[] { Miss, Miss, Miss, Miss });
                SetP(ws4, "Name", "Диаг Лист4");
                r = GetPI(ws4, "Range", new object[] { "A1" });
                SetP(r, "Value2", "S4_HEADER");
                r = GetPI(ws4, "Range", new object[] { "A2" });
                SetP(r, "Value2", "S4_dataA");
                Call(ws2, "Activate", new object[0]);
                Call(wb, "SaveCopyAs", new object[] { path });
                Call(wb, "Close", new object[] { false, Miss, Miss });
                try { Call(app, "Quit", new object[0]); } catch { }
                KillNewExcel(pidsBefore);
                return path;
            }
            catch
            {
                try { if (app != null) Call(app, "Quit", new object[0]); } catch { }
                KillNewExcel(pidsBefore);
                throw;
            }
        }

        public class SheetInfo
        {
            public int Index;
            public string Name;
            public override string ToString() { return Name; }
        }

        static void LogTable(Table t, string tag)
        {
            if (t == null) { Log(tag + ": таблицы нет"); return; }
            string s00 = "", s10 = "";
            try { s00 = t.GetTextString(0, 0, 0); } catch { s00 = "<err>"; }
            try { s10 = t.GetTextString(1, 0, 0); } catch { s10 = "<err>"; }
            Log(tag + ": rows=" + t.NumRows + " cols=" + t.NumColumns +
                " текст(0,0)=\"" + s00 + "\" текст(1,0)=\"" + s10 + "\"");
        }

        // один тест связи с пошаговым логом
        static void TestLink(string tag, Transaction tr, Database db, DataLinkManager dlm,
                             BlockTableRecord btr, DataLink dlTemplate, Point3d pt)
        {
            try
            {
                ObjectId oldId = dlm.GetDataLink(dlTemplate.Name);
                if (oldId != ObjectId.Null) { dlm.RemoveDataLink(oldId); Log(tag + ".0: старая связь удалена"); }
            }
            catch (System.Exception e) { Log(tag + ".0: RemoveDataLink - " + e.Message); }

            ObjectId dlId = ObjectId.Null;
            try
            {
                dlId = dlm.AddDataLink(dlTemplate);
                Log(tag + ".1: AddDataLink ok");
            }
            catch (System.Exception e) { Log(tag + ".1: AddDataLink ОШИБКА: " + e.Message); return; }

            try
            {
                tr.AddNewlyCreatedDBObject(dlTemplate, true);
                Log(tag + ".2: AddNewlyCreatedDBObject ok");
            }
            catch (System.Exception e) { Log(tag + ".2: AddNewlyCreatedDBObject ОШИБКА: " + e.Message); }

            Table t = null;
            try
            {
                t = AddTable(db, tr, btr, dlId, pt, 2, 2);
                Log(tag + ".3: таблица создана");
            }
            catch (System.Exception e) { Log(tag + ".3: AddTable ОШИБКА: " + e.Message); return; }

            Stopwatch swU = Stopwatch.StartNew();
            try
            {
                dlm.Update(UpdateDirection.SourceToData,
                           UpdateOption.UpdateRowHeight | UpdateOption.UpdateColumnWidth);
                swU.Stop();
                Log(tag + ".4: Update ok, " + swU.ElapsedMilliseconds + " мс");
            }
            catch (System.Exception e)
            {
                swU.Stop();
                Log(tag + ".4: Update ОШИБКА (" + swU.ElapsedMilliseconds + " мс): " + e.Message);
            }

            LogTable(t, tag);
        }

        // ───────────────────────────────────────────────────────────
        [CommandMethod("XLTABDIAG2")]
        public void Diag()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;

            try
            {
                Log("=== XLTABDIAG2 (.NET) ===");
                Log("AutoCAD: " + Autodesk.AutoCAD.ApplicationServices.Application.Version.ToString());

                string testPath = Path.Combine(Path.GetTempPath(), "xltabdiag_test.xlsx");
                try { if (File.Exists(testPath)) File.Delete(testPath); } catch { }
                Log("S1: создаю тестовый xlsx (4 листа)...");
                MakeTestXlsx(testPath);
                Log("S1: тестовый xlsx: " + testPath);

                Log("SZ: чтение xlsx напрямую (ZIP, без Excel)...");
                try
                {
                    Dictionary<int, XlsxSheet> xz = ReadXlsx(testPath);
                    foreach (KeyValuePair<int, XlsxSheet> kv in xz)
                        Log("SZ: лист " + kv.Key + " '" + kv.Value.Name + "' скрыт=" + kv.Value.Hidden +
                            " диапазон " + ColLetter(kv.Value.C1) + kv.Value.R1 + ":" +
                            ColLetter(kv.Value.C2) + kv.Value.R2 +
                            " колонок=" + kv.Value.ColW.Count + " строк=" + kv.Value.RowH.Count +
                            " шрифтовая ширина по умолч.=" + kv.Value.DefaultColW.ToString("0.##", CultureInfo.InvariantCulture));
                }
                catch (System.Exception eZ) { Log("SZ: ОШИБКА: " + eZ.Message); }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DataLinkManager dlm = db.DataLinkManager;
                    Log("S2: DataLinkManager: DataLinkCount=" + dlm.DataLinkCount);

                    foreach (ObjectId id in dlm.GetDataLink())
                    {
                        DataLink x = tr.GetObject(id, OpenMode.ForRead) as DataLink;
                        if (x == null) continue;
                        Log("S2: существующая связь: '" + x.Name + "' adapter=" + x.DataAdapterId +
                            " option=" + x.DataLinkOption + " updateOption=" + x.UpdateOption);
                        Log("    conn: " + x.ConnectionString);
                        try
                        {
                            ArrayList srcs = x.GetSourceFiles(DataLinkGetSourceContext.Other);
                            if (srcs != null) foreach (object s in srcs) Log("    source: " + s);
                        }
                        catch (System.Exception e2) { Log("    GetSourceFiles: " + e2.Message); }
                    }

                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    Log("TB: формат 'путь!лист!диапазон' - DiagnList2!A1:B2 (латиница)");
                    TestLink("TB", tr, db, dlm, btr,
                             MakeLink("XLTABDIAG_LINK_B", "AcExcel",
                                      testPath + "!DiagnList2!A1:B2", "test B"),
                             new Point3d(1000, 1000, 0));

                    Log("TC: формат 'путь!лист!диапазон' - кириллический лист Диаг_Лист3!A1:B2");
                    TestLink("TC", tr, db, dlm, btr,
                             MakeLink("XLTABDIAG_LINK_C", "AcExcel",
                                      testPath + "!Диаг_Лист3!A1:B2", "test C"),
                             new Point3d(1000, 700, 0));

                    Log("TD1: формат 'путь!лист!диапазон' - лист с пробелом БЕЗ кавычек: Диаг Лист4!A1:B2");
                    TestLink("TD1", tr, db, dlm, btr,
                             MakeLink("XLTABDIAG_LINK_D1", "AcExcel",
                                      testPath + "!Диаг Лист4!A1:B2", "test D1"),
                             new Point3d(1000, 400, 0));

                    Log("TF: адаптер ACEXCEL + старый синтаксис 'путь!лист!диапазон' - DiagnList2!A1:B2");
                    TestLink("TF", tr, db, dlm, btr,
                             MakeLink("XLTABDIAG_LINK_F", "ACEXCEL",
                                      testPath + "!DiagnList2!A1:B2", "test F"),
                             new Point3d(1500, 700, 0));

                    Log("TE: шапка-связь + данные-связь в одной таблице (якоря 0,0 и 1,0), двойной Update");
                    try
                    {
                        ObjectId oldEh = dlm.GetDataLink("XLTABDIAG_LINK_E_HDR");
                        if (oldEh != ObjectId.Null) dlm.RemoveDataLink(oldEh);
                        ObjectId oldEd = dlm.GetDataLink("XLTABDIAG_LINK_E_DATA");
                        if (oldEd != ObjectId.Null) dlm.RemoveDataLink(oldEd);

                        DataLink dlEh = MakeLink("XLTABDIAG_LINK_E_HDR", "AcExcel",
                                                 testPath + "!DiagnList1!A1:B1", "test E header");
                        ObjectId idEh = dlm.AddDataLink(dlEh);
                        tr.AddNewlyCreatedDBObject(dlEh, true);
                        DataLink dlEd = MakeLink("XLTABDIAG_LINK_E_DATA", "AcExcel",
                                                 testPath + "!DiagnList1!A2:B2", "test E data");
                        ObjectId idEd = dlm.AddDataLink(dlEd);
                        tr.AddNewlyCreatedDBObject(dlEd, true);

                        Table tE = new Table();
                        tE.TableStyle = db.Tablestyle;
                        tE.Position = new Point3d(1500, 1000, 0);
                        tE.NumRows = 2;
                        tE.NumColumns = 2;
                        btr.AppendEntity(tE);
                        tr.AddNewlyCreatedDBObject(tE, true);
                        tE.SetDataLink(0, 0, idEh, true);
                        tE.SetDataLink(1, 0, idEd, true);
                        dlm.Update(UpdateDirection.SourceToData,
                                   UpdateOption.UpdateRowHeight | UpdateOption.UpdateColumnWidth);
                        Log("TE.1: rows=" + tE.NumRows + " (0,0)=\"" + CellStr(tE, 0, 0) +
                            "\" (1,0)=\"" + CellStr(tE, 1, 0) + "\" (1,1)=\"" + CellStr(tE, 1, 1) + "\"");
                        dlm.Update(UpdateDirection.SourceToData,
                                   UpdateOption.UpdateRowHeight | UpdateOption.UpdateColumnWidth);
                        Log("TE.2: rows=" + tE.NumRows + " (0,0)=\"" + CellStr(tE, 0, 0) +
                            "\" (1,0)=\"" + CellStr(tE, 1, 0) + "\" (1,1)=\"" + CellStr(tE, 1, 1) + "\"");
                        Log("TE: ожидается (0,0)=S1_HEADER (шапка жива), (1,0)=S1_dataA, (1,1)=S1_dataB");
                    }
                    catch (System.Exception eE) { Log("TE: ОШИБКА: " + eE.Message); }

                    Log("ВЕРДИКТ: рабочий формат - 'путь!Лист!A1:N40', адаптер ACEXCEL," +
                        " имя листа БЕЗ кавычек. Кириллица и пробелы в имени листа - OK.");
                    Log("Современный формат 'ACEXCEL;DataSource=...' через API НЕ работает (eDataLinkSourceNotFound).");

                    tr.Commit();
                }
                Log("=== XLTABDIAG2 завершён ===");
            }
            catch (System.Exception ex)
            {
                Log("НЕПОЙМАННАЯ ОШИБКА: " + ex.Message);
            }
        }

        // ── ядро: создание таблиц по настройкам ─────────────────────
        static int CreateTables(Database db, Editor ed, string path, List<SheetInfo> picked,
                                Dictionary<int, XlsxSheet> xlsx, object wsc,
                                string rngIn, double scale, bool split,
                                double maxH1, double maxH2, Point3d pt)
        {
            const double gap = 15.0;
            int ok = 0;
            Stopwatch swAll = Stopwatch.StartNew();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DataLinkManager dlm = db.DataLinkManager;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                foreach (SheetInfo si in picked)
                {
                    XlsxSheet sh = null;
                    object ws = null;
                    if (xlsx != null) sh = xlsx[si.Index];
                    else ws = GetPI(wsc, "Item", new object[] { si.Index });

                    // границы диапазона
                    int c1 = 1, r1 = 1, c2 = 1, r2 = 1, rows = 1, cols = 1;
                    if (rngIn != "")
                    {
                        if (!ParseRange(rngIn, out c1, out r1, out c2, out r2))
                        {
                            ed.WriteMessage("\n  ПРОПУСК листа \"" + si.Name +
                                            "\": неверный диапазон \"" + rngIn + "\"");
                            continue;
                        }
                        rows = r2 - r1 + 1;
                        cols = c2 - c1 + 1;
                    }
                    else if (xlsx != null)
                    {
                        c1 = sh.C1; r1 = sh.R1; c2 = sh.C2; r2 = sh.R2;
                        rows = r2 - r1 + 1;
                        cols = c2 - c1 + 1;
                    }
                    else
                    {
                        object ur = GetP(ws, "UsedRange");
                        if (ur != null)
                        {
                            rows = (int)GetP(GetP(ur, "Rows"), "Count");
                            cols = (int)GetP(GetP(ur, "Columns"), "Count");
                            try { r1 = (int)GetP(ur, "Row"); } catch { }
                            try { c1 = (int)GetP(ur, "Column"); } catch { }
                        }
                        c2 = c1 + cols - 1;
                        r2 = r1 + rows - 1;
                    }

                    if (split && rows > 1)
                    {
                        double totalW = 0;
                        for (int c = c1; c <= c2; c++)
                            totalW += (xlsx != null) ? ColWmmZip(sh, c, scale) : ColWmm(ws, c, scale);
                        if (totalW > 390.0)
                            ed.WriteMessage("\n  ВНИМАНИЕ: лист \"" + si.Name + "\" - ширина " +
                                            totalW.ToString("0") + " мм > 390 (не влезет в А3 альбомную)");
                    }

                    if (!split || rows <= 1)
                    {
                        // без разбивки (одна таблица на весь диапазон)
                        string rangeStr = ColLetter(c1) + r1 + ":" + ColLetter(c2) + r2;
                        string linkName = "XL2TAB_" + si.Index + "_" + si.Name;
                        ObjectId oldId = dlm.GetDataLink(linkName);
                        if (oldId != ObjectId.Null) dlm.RemoveDataLink(oldId);

                        Stopwatch swL = Stopwatch.StartNew();
                        ObjectId dlId = CreateDataLinkLegacy(db, tr, linkName, path, si.Name, rangeStr);
                        Table t = AddTable(db, tr, btr, dlId, pt, rows, cols);
                        UpdateLink(tr, dlId);
                        swL.Stop();

                        string probe = null;
                        try { probe = t.GetTextString(0, 0, 0); } catch { }
                        if (probe == null || probe == "")
                            Log("  ВНИМАНИЕ: лист \"" + si.Name +
                                "\" дал пустую таблицу - запустите XLTABDIAG2 для диагностики");

                        Stopwatch swS = Stopwatch.StartNew();
                        t.SuppressRegenerateTable(true);
                        for (int c = 0; c < cols; c++)
                            t.SetColumnWidth(c, (xlsx != null) ? ColWmmZip(sh, c1 + c, scale) : ColWmm(ws, c1 + c, scale));
                        for (int r = 0; r < rows; r++)
                            t.SetRowHeight(r, (xlsx != null) ? RowHmmZip(sh, r1 + r, scale) : RowHmm(ws, r1 + r, scale));
                        t.SuppressRegenerateTable(false);
                        swS.Stop();
                        Log("  [" + si.Name + "] связь+обновление " + swL.ElapsedMilliseconds +
                            " мс, размеры " + swS.ElapsedMilliseconds + " мс");

                        pt = new Point3d(pt.X, pt.Y - t.Height - gap, pt.Z);
                        ok++;
                        ed.WriteMessage("\n  OK: лист \"" + si.Name + "\" - связь \"" + linkName + "\"");
                    }
                    else
                    {
                        // разбивка по высоте: шапка = первая строка диапазона отдельной связью,
                        // форматирование шапки приходит из Excel (не копируется текстом)
                        string hdrRange = ColLetter(c1) + r1 + ":" + ColLetter(c2) + r1;
                        string hdrLinkName = "XL2TAB_" + si.Index + "_" + si.Name + "_ШАПКА";
                        ObjectId hdrOld = dlm.GetDataLink(hdrLinkName);
                        if (hdrOld != ObjectId.Null) dlm.RemoveDataLink(hdrOld);
                        ObjectId hdrId = CreateDataLinkLegacy(db, tr, hdrLinkName, path, si.Name, hdrRange);

                        double headerH = (xlsx != null) ? RowHmmZip(sh, r1, scale) : RowHmm(ws, r1, scale);
                        int dataStart = r1 + 1;
                        int chunkNo = 1;
                        int cur = dataStart;
                        while (cur <= r2)
                        {
                            double limit = (chunkNo == 1) ? maxH1 : maxH2;
                            int chunkStart = cur;
                            double sum = headerH;
                            sum += (xlsx != null) ? RowHmmZip(sh, cur, scale) : RowHmm(ws, cur, scale);
                            cur++;
                            while (cur <= r2 &&
                                   sum + ((xlsx != null) ? RowHmmZip(sh, cur, scale) : RowHmm(ws, cur, scale)) <= limit)
                            {
                                sum += (xlsx != null) ? RowHmmZip(sh, cur, scale) : RowHmm(ws, cur, scale);
                                cur++;
                            }
                            int chunkRows = cur - chunkStart;

                            string chunkRange = ColLetter(c1) + chunkStart + ":" + ColLetter(c2) + (cur - 1);
                            string linkName = "XL2TAB_" + si.Index + "_" + si.Name + "_Ч" + chunkNo;
                            ObjectId oldId = dlm.GetDataLink(linkName);
                            if (oldId != ObjectId.Null) dlm.RemoveDataLink(oldId);
                            ObjectId dataId = CreateDataLinkLegacy(db, tr, linkName, path, si.Name, chunkRange);

                            Stopwatch swL = Stopwatch.StartNew();
                            Table t = AddTable2Links(db, tr, btr, hdrId, dataId, pt, chunkRows + 1, cols);
                            UpdateLink(tr, hdrId);
                            UpdateLink(tr, dataId);
                            swL.Stop();

                            Stopwatch swS = Stopwatch.StartNew();
                            t.SuppressRegenerateTable(true);
                            t.SetRowHeight(0, headerH);
                            for (int c = 0; c < cols; c++)
                                t.SetColumnWidth(c, (xlsx != null) ? ColWmmZip(sh, c1 + c, scale) : ColWmm(ws, c1 + c, scale));
                            for (int j = 0; j < chunkRows; j++)
                                t.SetRowHeight(1 + j, (xlsx != null) ? RowHmmZip(sh, chunkStart + j, scale)
                                                                     : RowHmm(ws, chunkStart + j, scale));
                            t.SuppressRegenerateTable(false);
                            swS.Stop();
                            Log("  [" + si.Name + " ч." + chunkNo + "] связи+обновление " +
                                swL.ElapsedMilliseconds + " мс, размеры " + swS.ElapsedMilliseconds + " мс");

                            pt = new Point3d(pt.X, pt.Y - t.Height - gap, pt.Z);
                            ok++;
                            ed.WriteMessage("\n  OK: лист \"" + si.Name + "\" ч." + chunkNo +
                                            " (строки " + chunkStart + "-" + (cur - 1) +
                                            ") - связь \"" + linkName + "\"");
                            chunkNo++;
                        }
                    }
                }
                tr.Commit();
            }

            swAll.Stop();
            Log("XL2TAB v11: ВСЕГО " + swAll.ElapsedMilliseconds + " мс (таблиц: " + ok + ")");
            return ok;
        }

        // ── пошаговый режим (запасной) ──────────────────────────────
        [CommandMethod("XL2TABPROMPT")]
        public void PromptRun()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string path = null;
            using (System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog())
            {
                ofd.Title = "Выберите файл Excel";
                ofd.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*";
                if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    ed.WriteMessage("\nXL2TAB: отменено.");
                    return;
                }
                path = ofd.FileName;
            }

            object app = null;
            object wsc = null;
            int[] excelPidsBefore = null;
            Dictionary<int, XlsxSheet> xlsx = null;
            try
            {
                if (path.ToLowerInvariant().EndsWith(".xlsx"))
                {
                    Stopwatch swZip = Stopwatch.StartNew();
                    try
                    {
                        xlsx = ReadXlsx(path);
                        swZip.Stop();
                        Log("XL2TAB v11: ZIP-чтение " + swZip.ElapsedMilliseconds +
                            " мс (листов: " + xlsx.Count + ")");
                    }
                    catch (System.Exception ez)
                    {
                        swZip.Stop();
                        Log("XL2TAB v11: ZIP-чтение не удалось за " + swZip.ElapsedMilliseconds +
                            " мс (" + ez.Message + ") - переключаюсь на Excel COM");
                        xlsx = null;
                    }
                }
                if (xlsx == null)
                {
                    Stopwatch swCom = Stopwatch.StartNew();
                    excelPidsBefore = GetExcelPids();
                    app = NewExcel();
                    SetP(app, "Visible", 0);
                    SetP(app, "DisplayAlerts", false);
                    try { SetP(app, "ScreenUpdating", false); } catch { }
                    try { SetP(app, "Calculation", -4135); } catch { } // xlCalculationManual - быстрее открытие
                    object wbs = GetP(app, "Workbooks");
                    object wb = Call(wbs, "Open", new object[] { path, Miss, true, Miss, Miss, Miss, Miss,
                                                                Miss, Miss, Miss, Miss, Miss, Miss, Miss, Miss });
                    wsc = GetP(wb, "Worksheets");
                    swCom.Stop();
                    Log("XL2TAB v11: Excel COM-открытие " + swCom.ElapsedMilliseconds + " мс");
                }

                List<SheetInfo> sheets = new List<SheetInfo>();
                if (xlsx != null)
                {
                    foreach (KeyValuePair<int, XlsxSheet> kv in xlsx)
                        if (!kv.Value.Hidden)
                            sheets.Add(new SheetInfo { Index = kv.Key, Name = kv.Value.Name });
                }
                else
                {
                    int cnt = (int)GetP(wsc, "Count");
                    for (int i = 1; i <= cnt; i++)
                    {
                        object ws = GetPI(wsc, "Item", new object[] { i });
                        int vis = (int)GetP(ws, "Visible");
                        if (vis == -1) sheets.Add(new SheetInfo { Index = i, Name = (string)GetP(ws, "Name") });
                    }
                }

                ed.WriteMessage("\nЛистов (видимых): " + sheets.Count);
                foreach (SheetInfo si in sheets)
                    ed.WriteMessage("\n  " + si.Index + ". " + si.Name);

                PromptResult pr = ed.GetString("\nНомера листов через запятую (Enter = все): ");
                List<SheetInfo> picked = new List<SheetInfo>();
                if (string.IsNullOrEmpty(pr.StringResult) || pr.StringResult.Trim() == "")
                {
                    picked.AddRange(sheets);
                }
                else
                {
                    string[] parts = pr.StringResult.Split(new char[] { ',', ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string p in parts)
                    {
                        int n;
                        if (int.TryParse(p, out n))
                            foreach (SheetInfo si in sheets)
                                if (si.Index == n) picked.Add(si);
                    }
                }
                if (picked.Count == 0)
                {
                    ed.WriteMessage("\nXL2TAB: листы не выбраны.");
                    return;
                }

                string rngIn = "";
                PromptResult rr = ed.GetString("\nДиапазон (Enter = весь лист, напр. A5:N40): ");
                if (rr.Status == PromptStatus.OK) rngIn = rr.StringResult.Trim();

                double scale = 1.0;
                PromptDoubleOptions pdo = new PromptDoubleOptions("\nМасштаб размеров");
                pdo.DefaultValue = 1.0;
                pdo.AllowNone = true;
                PromptDoubleResult pdr = ed.GetDouble(pdo);
                if (pdr.Status == PromptStatus.OK) scale = pdr.Value;
                if (scale <= 0) scale = 1.0;

                bool split = false;
                double maxH1 = 232.0, maxH2 = 247.0;
                PromptKeywordOptions pko = new PromptKeywordOptions("\nРазбивать по высоте? [Да/Нет] <Нет>: ");
                pko.Keywords.Add("Да");
                pko.Keywords.Add("Нет");
                pko.Keywords.Default = "Нет";
                pko.AllowNone = true;
                PromptResult pkr = ed.GetKeywords(pko);
                if (pkr.Status == PromptStatus.OK && pkr.StringResult == "Да")
                {
                    split = true;
                    PromptDoubleOptions pho1 = new PromptDoubleOptions("\nМакс. высота таблицы для 1-го листа, мм (штамп больше)");
                    pho1.DefaultValue = 232.0;
                    pho1.AllowNone = true;
                    PromptDoubleResult phr1 = ed.GetDouble(pho1);
                    if (phr1.Status == PromptStatus.OK) maxH1 = phr1.Value;
                    if (maxH1 <= 0) maxH1 = 232.0;

                    PromptDoubleOptions pho2 = new PromptDoubleOptions("\nМакс. высота таблицы для последующих листов, мм (штамп меньше)");
                    pho2.DefaultValue = 247.0;
                    pho2.AllowNone = true;
                    PromptDoubleResult phr2 = ed.GetDouble(pho2);
                    if (phr2.Status == PromptStatus.OK) maxH2 = phr2.Value;
                    if (maxH2 <= 0) maxH2 = 247.0;
                }

                Point3d pt = new Point3d(0, 0, 0);
                PromptPointOptions ppo = new PromptPointOptions("\nТочка вставки первой таблицы (верхний левый угол) <0,0>:");
                ppo.AllowNone = true;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.OK) pt = ppr.Value;

                int ok = CreateTables(db, ed, path, picked, xlsx, wsc, rngIn, scale, split, maxH1, maxH2, pt);

                ed.WriteMessage("\nXL2TAB: создано таблиц: " + ok +
                                ". Обновление данных: DATALINKUPDATE (связи ссылаются на исходный файл).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nXL2TAB ОШИБКА: " + ex.Message);
                Log("XL2TAB ОШИБКА: " + ex.Message);
            }
            finally
            {
                try { if (app != null) Call(app, "Quit", new object[0]); } catch { }
                if (excelPidsBefore != null) KillNewExcel(excelPidsBefore);
                GC.Collect();
            }
        }

        // ── лента: вкладка и кнопки ─────────────────────────────────
        public class AcadCommandHandler : System.Windows.Input.ICommand
        {
            public event EventHandler CanExecuteChanged;
            public bool CanExecute(object parameter) { return true; }
            public void Execute(object parameter)
            {
                try
                {
                    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    if (doc != null) doc.SendStringToExecute((string)parameter, true, false, false);
                }
                catch { }
            }
        }

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        static extern bool DeleteObject(IntPtr hObject);

        static System.Windows.Media.ImageSource MakeIcon(int size, string text)
        {
            using (Bitmap bmp = new Bitmap(size, size))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    using (System.Drawing.Font f = new System.Drawing.Font("Segoe UI", size * 0.40f, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (SolidBrush br = new SolidBrush(System.Drawing.Color.FromArgb(30, 90, 160)))
                    using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        g.DrawString(text, f, br, new RectangleF(0, 0, size, size), sf);
                }
                IntPtr hbmp = bmp.GetHbitmap();
                try
                {
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                        hbmp, IntPtr.Zero, System.Windows.Int32Rect.Empty,
                        System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
                }
                finally { DeleteObject(hbmp); }
            }
        }

        static void AddRibbonTab()
        {
            try
            {
                Autodesk.Windows.RibbonControl rc = Autodesk.Windows.ComponentManager.Ribbon;
                if (rc == null) return;
                foreach (Autodesk.Windows.RibbonTab t in rc.Tabs)
                    if (t.Id == "XL2TAB_TAB") return;

                Autodesk.Windows.RibbonTab tab = new Autodesk.Windows.RibbonTab();
                tab.Title = "XL2TAB";
                tab.Id = "XL2TAB_TAB";

                Autodesk.Windows.RibbonPanel panel = new Autodesk.Windows.RibbonPanel();
                Autodesk.Windows.RibbonPanelSource src = new Autodesk.Windows.RibbonPanelSource();
                src.Title = "Excel в таблицы";
                panel.Source = src;

                Autodesk.Windows.RibbonButton btn = new Autodesk.Windows.RibbonButton();
                btn.Id = "XL2TAB_BTN";
                btn.Text = "XL2TAB\nПалитра";
                btn.ShowText = true;
                btn.ShowImage = true;
                btn.Size = Autodesk.Windows.RibbonItemSize.Large;
                btn.LargeImage = MakeIcon(32, "X2T");
                btn.Image = MakeIcon(16, "X2T");
                btn.CommandParameter = "XL2TAB ";
                btn.CommandHandler = new AcadCommandHandler();
                btn.ToolTip = "XL2TAB: Excel в связанные таблицы AutoCAD";
                src.Items.Add(btn);

                Autodesk.Windows.RibbonButton btn2 = new Autodesk.Windows.RibbonButton();
                btn2.Id = "XL2TAB_PROMPT_BTN";
                btn2.Text = "Пошагово";
                btn2.ShowText = true;
                btn2.ShowImage = true;
                btn2.Size = Autodesk.Windows.RibbonItemSize.Standard;
                btn2.Image = MakeIcon(16, ">");
                btn2.CommandParameter = "XL2TABPROMPT ";
                btn2.CommandHandler = new AcadCommandHandler();
                btn2.ToolTip = "XL2TABPROMPT: пошаговый режим";
                src.Items.Add(btn2);

                Autodesk.Windows.RibbonButton btn3 = new Autodesk.Windows.RibbonButton();
                btn3.Id = "XL2TAB_DIAG_BTN";
                btn3.Text = "Диагностика";
                btn3.ShowText = true;
                btn3.ShowImage = true;
                btn3.Size = Autodesk.Windows.RibbonItemSize.Standard;
                btn3.Image = MakeIcon(16, "?");
                btn3.CommandParameter = "XLTABDIAG2 ";
                btn3.CommandHandler = new AcadCommandHandler();
                btn3.ToolTip = "XLTABDIAG2: диагностика";
                src.Items.Add(btn3);

                tab.Panels.Add(panel);
                rc.Tabs.Add(tab);
            }
            catch (System.Exception ex) { Log("Лента: " + ex.Message); }
        }

        static void OnItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
        {
            try
            {
                if (Autodesk.Windows.ComponentManager.Ribbon != null)
                {
                    Autodesk.Windows.ComponentManager.ItemInitialized -= OnItemInitialized;
                    AddRibbonTab();
                }
            }
            catch { }
        }

        // ── классическая панель инструментов (для workspace без ленты) ──
        // kind: 0 - таблица, 1 - шаги, 2 - лупа
        static string EnsureIcon(string fileName, int size, int kind)
        {
            string path = Path.Combine(Path.GetTempPath(), fileName);
            try
            {
                // 24bpp без альфы: AutoCAD считает прозрачным цвет фона (маджента)
                using (Bitmap bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;
                        // маджента - прозрачный цвет в классических панелях AutoCAD
                        g.Clear(System.Drawing.Color.Magenta);
                        float k = size / 32f; // масштаб от макета 32px
                        if (kind == 0)
                        {
                            // таблица: рамка, синяя шапка, сетка
                            RectangleF body = new RectangleF(3 * k, 6 * k, 26 * k, 20 * k);
                            using (SolidBrush w = new SolidBrush(System.Drawing.Color.FromArgb(245, 245, 245)))
                                g.FillRectangle(w, body);
                            using (SolidBrush hdr = new SolidBrush(System.Drawing.Color.FromArgb(40, 110, 180)))
                                g.FillRectangle(hdr, new RectangleF(3 * k, 6 * k, 26 * k, 6.5f * k));
                            using (Pen p = new Pen(System.Drawing.Color.FromArgb(60, 60, 60), Math.Max(1f, 1.2f * k)))
                            {
                                g.DrawRectangle(p, body.X, body.Y, body.Width, body.Height);
                                g.DrawLine(p, 3 * k, 19 * k, 29 * k, 19 * k);
                                g.DrawLine(p, 12 * k, 6 * k, 12 * k, 26 * k);
                                g.DrawLine(p, 20 * k, 6 * k, 20 * k, 26 * k);
                            }
                        }
                        else if (kind == 1)
                        {
                            // шаги: три строки с маркерами
                            using (SolidBrush b = new SolidBrush(System.Drawing.Color.FromArgb(40, 110, 180)))
                            using (Pen p = new Pen(System.Drawing.Color.FromArgb(220, 220, 220), Math.Max(1f, 2f * k)))
                            {
                                float[] ys = { 8f, 16f, 24f };
                                foreach (float y in ys)
                                {
                                    g.FillRectangle(b, 4 * k, (y - 2.5f) * k, 5 * k, 5 * k);
                                    g.DrawLine(p, 12 * k, y * k, 28 * k, y * k);
                                }
                            }
                        }
                        else
                        {
                            // лупа
                            using (Pen p = new Pen(System.Drawing.Color.FromArgb(220, 220, 220), Math.Max(1.5f, 2.5f * k)))
                                g.DrawEllipse(p, 5 * k, 5 * k, 15 * k, 15 * k);
                            using (Pen h = new Pen(System.Drawing.Color.FromArgb(40, 110, 180), Math.Max(1.5f, 3f * k)))
                                g.DrawLine(h, 18 * k, 18 * k, 27 * k, 27 * k);
                        }
                    }
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
                }
            }
            catch { }
            return path;
        }

        // маска прозрачности (Windows AND-маска): белый = прозрачно, чёрный = непрозрачно
        static string EnsureMask(string srcPath, string maskPath)
        {
            try
            {
                using (Bitmap src = new Bitmap(srcPath))
                using (Bitmap mask = new Bitmap(src.Width, src.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
                {
                    System.Drawing.Color mag = System.Drawing.Color.Magenta;
                    for (int y = 0; y < src.Height; y++)
                    {
                        for (int x = 0; x < src.Width; x++)
                        {
                            System.Drawing.Color px = src.GetPixel(x, y);
                            bool transparent = (px.R == mag.R && px.G == mag.G && px.B == mag.B);
                            mask.SetPixel(x, y, transparent ? System.Drawing.Color.White : System.Drawing.Color.Black);
                        }
                    }
                    mask.Save(maskPath, System.Drawing.Imaging.ImageFormat.Bmp);
                }
            }
            catch { }
            return maskPath;
        }

        static void AddTbButton(object tb, int idx, string name, string help, string macro, int kind)
        {
            try
            {
                object btn = Call(tb, "AddToolbarButton", new object[] { idx, name, help, macro, false });
                try
                {
                    string bmp16 = EnsureIcon("xl2tab_tb" + (idx + 1) + "_16.bmp", 16, kind);
                    string mask = EnsureMask(bmp16, Path.Combine(Path.GetTempPath(), "xl2tab_tb" + (idx + 1) + "_mask.bmp"));
                    Call(btn, "SetBitmaps", new object[] { bmp16, mask });
                }
                catch (System.Exception ex) { Log("Тулбар-иконка: " + ex.Message); }
            }
            catch (System.Exception ex) { Log("Тулбар-кнопка: " + ex.Message); }
        }

        // создаёт (или обновляет) классическую панель XL2TAB через COM-API AutoCAD
        static void EnsureClassicToolbar(string where)
        {
            try
            {
                if (Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument == null)
                {
                    Log("Тулбар (" + where + "): нет открытого документа, создам после его открытия");
                    return;
                }
                object app = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
                if (app == null) { Log("Тулбар (" + where + "): AcadApplication недоступен"); return; }
                object groups = GetP(app, "MenuGroups");
                object group = GetPI(groups, "Item", new object[] { 0 });
                object toolbars = GetP(group, "Toolbars");
                object tb = null;
                try { tb = GetPI(toolbars, "Item", new object[] { "XL2TAB" }); } catch { }
                if (tb == null)
                {
                    tb = Call(toolbars, "Add", new object[] { "XL2TAB" });
                    AddTbButton(tb, 0, "XL2TAB",
                        "Excel в связанные таблицы AutoCAD (DataLink)\nОткрыть палитру настроек\nКоманда: XL2TAB",
                        "^C^C_XL2TAB ", 0);
                    AddTbButton(tb, 1, "Пошагово",
                        "Пошаговый режим без палитры (в командной строке)\nКоманда: XL2TABPROMPT",
                        "^C^C_XL2TABPROMPT ", 1);
                    AddTbButton(tb, 2, "Диагностика",
                        "Диагностика DataLink-связей\nЛог: %TEMP%\\xltabdiag_log.txt\nКоманда: XLTABDIAG2",
                        "^C^C_XLTABDIAG2 ", 2);
                    try { Call(tb, "Dock", new object[] { 1 }); } catch { } // 1 = сверху
                    Log("Тулбар XL2TAB создан (" + where + ")");
                }
                SetP(tb, "Visible", true);
                _toolbarReady = true;
            }
            catch (System.Exception ex)
            {
                string m = ex.Message;
                if (ex.InnerException != null) m += " / " + ex.InnerException.Message;
                Log("Тулбар (" + where + "): " + m);
            }
        }

        // ── палитра ─────────────────────────────────────────────────
        internal static PaletteSet Palette;
        internal static Xl2TabControl PaletteControl;

        internal static void ShowPalette()
        {
            try
            {
                if (Palette == null)
                {
                    Palette = new PaletteSet("XL2TAB");
                    Palette.Style = PaletteSetStyles.ShowAutoHideButton
                                  | PaletteSetStyles.ShowCloseButton
                                  | PaletteSetStyles.ShowPropertiesMenu;
                    PaletteControl = new Xl2TabControl();
                    Palette.Add("XL2TAB", PaletteControl);
                    Palette.Size = new Size(340, 620);
                    Palette.MinimumSize = new Size(300, 420);
                }
                Palette.Visible = true;
            }
            catch (System.Exception ex) { Log("Палитра: " + ex.Message); }
        }

        [CommandMethod("XL2TAB")]
        public void ShowXl2TabPalette()
        {
            EnsureClassicToolbar("команда");
            ShowPalette();
        }

        // вызывается из палитры через SendStringToExecute
        [CommandMethod("XL2TABRUN")]
        public void RunFromPalette()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            if (!Xl2TabControl.HasPending)
            {
                ed.WriteMessage("\nXL2TAB: сначала настройте параметры в палитре (команда XL2TAB).");
                return;
            }

            string path = Xl2TabControl.PendingPath;
            List<SheetInfo> picked = Xl2TabControl.PendingSheets;
            string rngIn = Xl2TabControl.PendingRange ?? "";
            double scale = Xl2TabControl.PendingScale;
            bool split = Xl2TabControl.PendingSplit;
            double maxH1 = Xl2TabControl.PendingH1;
            double maxH2 = Xl2TabControl.PendingH2;

            object app = null;
            object wsc = null;
            int[] excelPidsBefore = null;
            Dictionary<int, XlsxSheet> xlsx = null;
            try
            {
                Log("XL2TAB v11: файл: " + path);
                if (path.ToLowerInvariant().EndsWith(".xlsx"))
                {
                    try { xlsx = ReadXlsx(path); }
                    catch (System.Exception ez) { Log("ZIP-чтение: " + ez.Message); xlsx = null; }
                }
                if (xlsx == null)
                {
                    excelPidsBefore = GetExcelPids();
                    app = NewExcel();
                    SetP(app, "Visible", 0);
                    SetP(app, "DisplayAlerts", false);
                    try { SetP(app, "ScreenUpdating", false); } catch { }
                    try { SetP(app, "Calculation", -4135); } catch { }
                    object wbs = GetP(app, "Workbooks");
                    object wb = Call(wbs, "Open", new object[] { path, Miss, true, Miss, Miss, Miss, Miss,
                                                                Miss, Miss, Miss, Miss, Miss, Miss, Miss, Miss });
                    wsc = GetP(wb, "Worksheets");
                }

                Point3d pt = new Point3d(0, 0, 0);
                PromptPointOptions ppo = new PromptPointOptions("\nТочка вставки первой таблицы (верхний левый угол) <0,0>:");
                ppo.AllowNone = true;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.OK) pt = ppr.Value;

                int ok = CreateTables(db, ed, path, picked, xlsx, wsc, rngIn, scale, split, maxH1, maxH2, pt);
                ed.WriteMessage("\nXL2TAB: создано таблиц: " + ok + ".");
                if (PaletteControl != null) PaletteControl.SetStatus("Готово: таблиц " + ok + ".");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nXL2TAB ОШИБКА: " + ex.Message);
                Log("XL2TAB ОШИБКА: " + ex.Message);
                if (PaletteControl != null) PaletteControl.SetStatus("Ошибка: " + ex.Message);
            }
            finally
            {
                try { if (app != null) Call(app, "Quit", new object[0]); } catch { }
                if (excelPidsBefore != null) KillNewExcel(excelPidsBefore);
                GC.Collect();
            }
            ShowPalette();
        }

        // ── панель палитры ──────────────────────────────────────────
        public class Xl2TabControl : UserControl
        {
            public static bool HasPending;
            public static string PendingPath;
            public static List<SheetInfo> PendingSheets;
            public static string PendingRange = "";
            public static double PendingScale = 1.0;
            public static bool PendingSplit;
            public static double PendingH1 = 232.0;
            public static double PendingH2 = 247.0;

            string _path;
            Dictionary<int, XlsxSheet> _xlsx;
            List<SheetInfo> _sheets;

            Label _lblFile, _lblStatus;
            Button _btnFile, _btnAll, _btnNone, _btnCreate;
            CheckedListBox _clb;
            TextBox _txtRange;
            NumericUpDown _nudScale, _nudH1, _nudH2;
            CheckBox _chkSplit;

            public Xl2TabControl()
            {
                Width = 330;
                Height = 560;
                Font = new System.Drawing.Font("Segoe UI", 9f);

                _btnFile = new Button { Text = "Выбрать файл...", Left = 8, Top = 8, Width = 128, Height = 26 };
                _lblFile = new Label { Text = "файл не выбран", Left = 144, Top = 13, Width = 172, AutoEllipsis = true };
                Label lblSheets = new Label { Text = "Листы:", Left = 8, Top = 42, Width = 60 };
                _clb = new CheckedListBox { Left = 8, Top = 60, Width = 306, Height = 176, CheckOnClick = true };
                _btnAll = new Button { Text = "Все", Left = 8, Top = 242, Width = 70, Height = 24 };
                _btnNone = new Button { Text = "Ничего", Left = 84, Top = 242, Width = 70, Height = 24 };
                Label lblRange = new Label { Text = "Диапазон (пусто = весь лист):", Left = 8, Top = 274, Width = 250 };
                _txtRange = new TextBox { Left = 8, Top = 292, Width = 306 };
                Label lblScale = new Label { Text = "Масштаб размеров:", Left = 8, Top = 324, Width = 160 };
                _nudScale = new NumericUpDown
                {
                    Left = 8, Top = 342, Width = 70, DecimalPlaces = 2,
                    Minimum = 0.01M, Maximum = 100M, Value = 1M, Increment = 0.1M
                };
                _chkSplit = new CheckBox { Text = "Разбивать по высоте (под А3)", Left = 8, Top = 378, Width = 270, Height = 22 };
                Label lblH1 = new Label { Text = "Высота, мм (1-й лист):", Left = 8, Top = 404, Width = 150 };
                _nudH1 = new NumericUpDown
                {
                    Left = 8, Top = 422, Width = 78, DecimalPlaces = 1,
                    Minimum = 10M, Maximum = 3000M, Value = 232M, Enabled = false
                };
                Label lblH2 = new Label { Text = "последующие:", Left = 96, Top = 404, Width = 120 };
                _nudH2 = new NumericUpDown
                {
                    Left = 96, Top = 422, Width = 78, DecimalPlaces = 1,
                    Minimum = 10M, Maximum = 3000M, Value = 247M, Enabled = false
                };
                _btnCreate = new Button { Text = "Создать таблицы", Left = 8, Top = 460, Width = 160, Height = 34 };
                _lblStatus = new Label { Text = "", Left = 8, Top = 502, Width = 306, Height = 48 };

                Controls.AddRange(new Control[] {
                    _btnFile, _lblFile, lblSheets, _clb, _btnAll, _btnNone,
                    lblRange, _txtRange, lblScale, _nudScale, _chkSplit,
                    lblH1, _nudH1, lblH2, _nudH2, _btnCreate, _lblStatus
                });

                _btnFile.Click += (s, e) => PickFile();
                _btnAll.Click += (s, e) => CheckAll(true);
                _btnNone.Click += (s, e) => CheckAll(false);
                _chkSplit.CheckedChanged += (s, e) => { _nudH1.Enabled = _nudH2.Enabled = _chkSplit.Checked; };
                _btnCreate.Click += (s, e) => Create();
            }

            public void SetStatus(string text)
            {
                _lblStatus.Text = text;
            }

            void CheckAll(bool value)
            {
                for (int i = 0; i < _clb.Items.Count; i++)
                    _clb.SetItemChecked(i, value);
            }

            void PickFile()
            {
                using (System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog())
                {
                    ofd.Title = "Выберите файл Excel";
                    ofd.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*";
                    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    LoadFile(ofd.FileName);
                }
            }

            void LoadFile(string path)
            {
                try
                {
                    _xlsx = null;
                    _sheets = null;
                    List<SheetInfo> sheets = new List<SheetInfo>();

                    if (path.ToLowerInvariant().EndsWith(".xlsx"))
                    {
                        try
                        {
                            _xlsx = ReadXlsx(path);
                            foreach (KeyValuePair<int, XlsxSheet> kv in _xlsx)
                                if (!kv.Value.Hidden)
                                    sheets.Add(new SheetInfo { Index = kv.Key, Name = kv.Value.Name });
                        }
                        catch (System.Exception ez)
                        {
                            _xlsx = null;
                            SetStatus("ZIP не прочитан: " + ez.Message);
                        }
                    }

                    if (_xlsx == null)
                    {
                        // Excel COM (для .xls или если ZIP не прочитался)
                        int[] pids = GetExcelPids();
                        object app = null;
                        try
                        {
                            app = NewExcel();
                            SetP(app, "Visible", 0);
                            SetP(app, "DisplayAlerts", false);
                            object wbs = GetP(app, "Workbooks");
                            object wb = Call(wbs, "Open", new object[] { path, Miss, true, Miss, Miss, Miss, Miss,
                                                                        Miss, Miss, Miss, Miss, Miss, Miss, Miss, Miss });
                            object wsc = GetP(wb, "Worksheets");
                            int cnt = (int)GetP(wsc, "Count");
                            for (int i = 1; i <= cnt; i++)
                            {
                                object ws = GetPI(wsc, "Item", new object[] { i });
                                int vis = (int)GetP(ws, "Visible");
                                if (vis == -1) sheets.Add(new SheetInfo { Index = i, Name = (string)GetP(ws, "Name") });
                            }
                        }
                        finally
                        {
                            try { if (app != null) Call(app, "Quit", new object[0]); } catch { }
                            KillNewExcel(pids);
                        }
                    }

                    _sheets = sheets;
                    _path = path;
                    _lblFile.Text = Path.GetFileName(path);
                    _clb.Items.Clear();
                    foreach (SheetInfo si in sheets) _clb.Items.Add(si, true);
                    SetStatus("Листов: " + sheets.Count + ". Отметьте нужные и нажмите «Создать».");
                }
                catch (System.Exception ex)
                {
                    SetStatus("Ошибка чтения файла: " + ex.Message);
                }
            }

            void Create()
            {
                if (_path == null || _sheets == null || _sheets.Count == 0)
                {
                    SetStatus("Сначала выберите файл Excel.");
                    return;
                }
                List<SheetInfo> picked = new List<SheetInfo>();
                for (int i = 0; i < _clb.Items.Count; i++)
                    if (_clb.GetItemChecked(i)) picked.Add((SheetInfo)_clb.Items[i]);
                if (picked.Count == 0)
                {
                    SetStatus("Отметьте хотя бы один лист.");
                    return;
                }

                PendingPath = _path;
                PendingSheets = picked;
                PendingRange = (_txtRange.Text ?? "").Trim();
                PendingScale = (double)_nudScale.Value;
                PendingSplit = _chkSplit.Checked;
                PendingH1 = (double)_nudH1.Value;
                PendingH2 = (double)_nudH2.Value;
                HasPending = true;

                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    SetStatus("Нет активного чертежа.");
                    return;
                }

                if (Palette != null) Palette.Visible = false;
                SetStatus("Выберите точку вставки в чертеже...");
                doc.SendStringToExecute("XL2TABRUN\n", true, false, false);
            }
        }
    }
}
