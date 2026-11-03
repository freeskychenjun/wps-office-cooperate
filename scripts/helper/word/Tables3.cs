// v1.33 (2026-11-03) - table -Action plan|normalize, built from session-a305706b (14 tables, caption
// numbering + in-table captions + thick borders: 127 tool calls / 88.6 min, of which 43 steps were
// reconnaissance and rework that one command can do). Reference implementation: design-report-format
// (one exe call does scan -> format -> table, progress on stderr, no round trips).
//
//   plan       READ-ONLY structural reconnaissance: per table page span / caption side / borders /
//              table style + grid columns / risks, plus the caption-numbering proposal and the
//              in-text reference map. Reuses `captions` wholesale, so plan / captions / audit can
//              never disagree about a caption.
//   normalize  one command, per-table isolation: caption out of row 1 (three insertion methods, each
//              PROVEN by reading the page back), caption spacing, optional renumber, thick borders
//              with a grid-column snapshot/restore, header repeat, then a verification sweep.
//
// Why the column widths are snapshotted from WordOpenXML and not from COM (measured 2026-11-03):
//   writing borders into a table whose XML has no <w:tblBorders> makes WPS rebuild the table layout;
//   Column.Width then reports 9999999 (wdUndefined), so the width cannot be read back through COM at
//   all. The grid columns come from doc.Content.WordOpenXML instead (ONE COM call, ~0.16 s for a
//   4 MB document). The real trigger of the width blow-up is a table WITHOUT <w:tblBorders>, so a
//   table that already carries the wanted borders is skipped instead of being written again.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

internal static partial class Program
{
    // ------------------------------------------------------------ table XML facts (one COM fetch)

    /// <summary>What one &lt;w:tbl&gt; element says about itself, read from doc.Content.WordOpenXML.</summary>
    private sealed class TableXmlInfo
    {
        public int Index;
        public string Style = "";          // tblStyle w:val, "" = the table has no style (the fragile case)
        public string WidthType = "";      // pct / dxa / auto / ""
        public double Width;               // as written (pct = 50ths of a percent, dxa = twips)
        public string Layout = "";         // autofit / fixed
        public List<double> GridCols = new List<double>();   // twips, one per column
        public string FirstTcWType = "";
        public double FirstTcW;
        public bool HasBorders;            // <w:tblBorders> present - the width-blow-up trigger
        public int TopSz;                  // eighths of a point (OOXML w:sz)
        public int InsideHSz;
        public string Raw = "";            // the <w:tbl>...</w:tbl> chunk, kept only when asked for
        public bool Thick;                 // borders present, outer >= 1.5 pt, inner <= 0.5 pt
    }

    private static readonly Regex RxTblOpen = new Regex("<w:tbl(?=[ >])", RegexOptions.Compiled);
    private static readonly Regex RxTblPr = new Regex("<w:tblPr>(.*?)</w:tblPr>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex RxTblStyle = new Regex("<w:tblStyle w:val=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxTblW = new Regex("<w:tblW ([^/>]*)/>", RegexOptions.Compiled);
    private static readonly Regex RxLayout = new Regex("<w:tblLayout w:type=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxTblBorders = new Regex("<w:tblBorders>(.*?)</w:tblBorders>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex RxGrid = new Regex("<w:tblGrid>(.*?)</w:tblGrid>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex RxGridCol = new Regex("<w:gridCol w:w=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxTopSz = new Regex("<w:top [^/>]*w:sz=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxInsideHSz = new Regex("<w:insideH [^/>]*w:sz=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxFirstRow = new Regex("<w:tr(?=[ >])(.*?)</w:tr>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex RxTcW = new Regex("<w:tcW ([^/>]*)/>", RegexOptions.Compiled);
    private static readonly Regex RxAttrW = new Regex("w=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxAttrType = new Regex("type=\"([^\"]*)\"", RegexOptions.Compiled);
    private static readonly Regex RxAttrSz = new Regex("sz=\"([^\"]*)\"", RegexOptions.Compiled);

    /// <summary>The whole story as WordOpenXML (one COM call). null when the host refuses.</summary>
    private static string WholeOpenXml(dynamic doc)
    {
        try
        {
            string xml = Convert.ToString(doc.Content.WordOpenXML, CultureInfo.InvariantCulture);
            return (xml == null || xml.Length == 0) ? null : xml;
        }
        catch (Exception e)
        {
            Tr("wordopenxml failed: " + Clean(e.Message));
            return null;
        }
    }

    /// <summary>Every top-level &lt;w:tbl&gt; chunk in document order. Nested tables stay inside their parent.</summary>
    private static List<string> TableXmlChunks(string xml)
    {
        List<string> list = new List<string>();
        if (xml == null) return list;
        int pos = 0;
        while (pos < xml.Length)
        {
            Match m = RxTblOpen.Match(xml, pos);
            if (!m.Success) break;
            int start = m.Index;
            int depth = 0;
            int j = start;
            int close = -1;
            while (true)
            {
                Match mo = RxTblOpen.Match(xml, j + 1);
                int mc = xml.IndexOf("</w:tbl>", j + 1, StringComparison.Ordinal);
                if (mc < 0) break;
                if (mo.Success && mo.Index < mc) { depth++; j = mo.Index; }
                else if (depth > 0) { depth--; j = mc; }
                else { close = mc + 8; break; }
            }
            if (close < 0) break;
            list.Add(xml.Substring(start, close - start));
            pos = close;
        }
        return list;
    }

    private static string AttrOf(string attrs, Regex rx)
    {
        if (attrs == null) return "";
        Match m = rx.Match(attrs);
        return m.Success ? m.Groups[1].Value : "";
    }

    private static double XmlNum(string s, double fallback)
    {
        double v;
        if (s != null && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return v;
        return fallback;
    }

    /// <summary>The per-table facts of one WordOpenXML fetch. keepRaw keeps each table's XML chunk (rollback).</summary>
    private static List<TableXmlInfo> ParseTableXml(string xml, bool keepRaw)
    {
        List<TableXmlInfo> list = new List<TableXmlInfo>();
        List<string> chunks = TableXmlChunks(xml);
        for (int i = 0; i < chunks.Count; i++)
        {
            string chunk = chunks[i];
            TableXmlInfo info = new TableXmlInfo();
            info.Index = i + 1;
            if (keepRaw) info.Raw = chunk;
            Match pr = RxTblPr.Match(chunk);
            string props = pr.Success ? pr.Groups[1].Value : "";
            Match st = RxTblStyle.Match(props);
            info.Style = st.Success ? st.Groups[1].Value : "";
            Match tw = RxTblW.Match(props);
            if (tw.Success)
            {
                info.WidthType = AttrOf(tw.Groups[1].Value, RxAttrType);
                info.Width = XmlNum(AttrOf(tw.Groups[1].Value, RxAttrW), 0);
            }
            Match la = RxLayout.Match(props);
            info.Layout = la.Success ? la.Groups[1].Value : "";
            Match bd = RxTblBorders.Match(props);
            if (bd.Success)
            {
                info.HasBorders = true;
                string b = bd.Groups[1].Value;
                Match mt = RxTopSz.Match(b);
                if (mt.Success) info.TopSz = (int)XmlNum(AttrOf(mt.Groups[1].Value, RxAttrSz), 0);
                Match mh = RxInsideHSz.Match(b);
                if (mh.Success) info.InsideHSz = (int)XmlNum(AttrOf(mh.Groups[1].Value, RxAttrSz), 0);
            }
            Match grid = RxGrid.Match(chunk);
            if (grid.Success)
            {
                foreach (Match g in RxGridCol.Matches(grid.Groups[1].Value))
                    info.GridCols.Add(XmlNum(g.Groups[1].Value, 0));
            }
            Match tr = RxFirstRow.Match(chunk);
            if (tr.Success)
            {
                Match tcw = RxTcW.Match(tr.Groups[1].Value);
                if (tcw.Success)
                {
                    info.FirstTcWType = AttrOf(tcw.Groups[1].Value, RxAttrType);
                    info.FirstTcW = XmlNum(AttrOf(tcw.Groups[1].Value, RxAttrW), 0);
                }
            }
            info.Thick = info.HasBorders && info.TopSz >= 12 && info.InsideHSz >= 1 && info.InsideHSz <= 4;
            list.Add(info);
        }
        return list;
    }

    /// <summary>Grid columns in points (twips / 20), rounded, as a JSON-friendly list.</summary>
    private static List<object> GridColsPt(TableXmlInfo info)
    {
        List<object> list = new List<object>();
        if (info == null) return list;
        foreach (double tw in info.GridCols) list.Add(Math.Round(tw / 20.0, 1));
        return list;
    }

    private static double GridSumPt(TableXmlInfo info)
    {
        if (info == null) return 0;
        double sum = 0;
        foreach (double tw in info.GridCols) sum += tw;
        return sum / 20.0;
    }

    /// <summary>
    /// How far the grid columns drifted between two snapshots: the larger of "worst column relative
    /// change" and "sum relative change". 0 = identical, 1 = every column collapsed.
    /// </summary>
    private static double GridDrift(TableXmlInfo before, TableXmlInfo after)
    {
        if (before == null || after == null) return 0;
        if (before.GridCols.Count == 0 || before.GridCols.Count != after.GridCols.Count) return 0;
        double worst = 0, sumBefore = 0, sumAfter = 0;
        for (int i = 0; i < before.GridCols.Count; i++)
        {
            double b = before.GridCols[i], a = after.GridCols[i];
            sumBefore += b;
            sumAfter += a;
            if (b <= 0) continue;
            double rel = Math.Abs(a - b) / b;
            if (rel > worst) worst = rel;
        }
        if (sumBefore > 0)
        {
            double relSum = Math.Abs(sumAfter - sumBefore) / sumBefore;
            if (relSum > worst) worst = relSum;
        }
        return worst;
    }

    /// <summary>[outer top line width, inner horizontal line width] as the borders read back (0 when unset).</summary>
    private static int[] BorderPair(dynamic table)
    {
        int outer = 0, inner = 0;
        try
        {
            dynamic b = table.Borders.Item(BT_TOP);
            outer = Convert.ToInt32(TryGet(() => (object)(int)b.LineWidth, (object)0), CultureInfo.InvariantCulture);
            if (Convert.ToInt32(TryGet(() => (object)(int)b.LineStyle, (object)0), CultureInfo.InvariantCulture) == 0) outer = 0;
        }
        catch (Exception) { }
        try
        {
            dynamic b = table.Borders.Item(BT_HORZ);
            inner = Convert.ToInt32(TryGet(() => (object)(int)b.LineWidth, (object)0), CultureInfo.InvariantCulture);
            if (Convert.ToInt32(TryGet(() => (object)(int)b.LineStyle, (object)0), CultureInfo.InvariantCulture) == 0) inner = 0;
        }
        catch (Exception) { }
        // wdUndefined (9999999) is what WPS reports for "no width": it is not a width at all, and
        // letting it through made plan print borders:[0,9999999] for a table with no borders.
        if (outer > 100) outer = 0;
        if (inner > 100) inner = 0;
        return new int[] { outer, inner };
    }

    /// <summary>The scan predicate, kept identical to TabScan so plan / normalize / scan agree: thick = 1.5 pt frame, 0.5 pt rule.</summary>
    private static bool ThickBorders(int[] pair)
    {
        return pair[0] >= 12 && pair[1] <= 4;
    }

    private static int PageOfPara(dynamic doc, int para)
    {
        if (para < 1) return 0;
        try { return (int)doc.Paragraphs.Item(para).Range.Information(WD_PAGE); }
        catch (Exception) { return 0; }
    }

    private static int SectionOfPara(dynamic doc, int para)
    {
        if (para < 1) return 0;
        try { return (int)doc.Paragraphs.Item(para).Range.Sections.Item(1).Index; }
        catch (Exception) { return 0; }
    }

    /// <summary>
    /// DESIGN 3.1 step 1: is there a page break or a section break between the previous body
    /// paragraph and the table? Each answer is a different insertion method later on.
    /// </summary>
    private static void DetectBreakBefore(dynamic doc, dynamic table, int startPara,
        out bool tableAtSectionStart, out bool sectPrBefore, out bool pageBreakChar, out bool pageBreakBefore)
    {
        tableAtSectionStart = false;
        sectPrBefore = false;
        pageBreakChar = false;
        pageBreakBefore = false;

        int tableStart = Convert.ToInt32(TryGet(() => (object)(int)table.Range.Start, (object)-1), CultureInfo.InvariantCulture);
        int secIndex = Convert.ToInt32(TryGet(() => (object)(int)table.Range.Sections.Item(1).Index, (object)0), CultureInfo.InvariantCulture);
        if (secIndex > 0)
        {
            int secStart = Convert.ToInt32(TryGet(() => (object)(int)doc.Sections.Item(secIndex).Range.Start, (object)-1), CultureInfo.InvariantCulture);
            if (secStart >= 0 && tableStart >= 0 && tableStart <= secStart + 1) tableAtSectionStart = true;
        }
        if (startPara <= 1) return;

        dynamic prev = null;
        try { prev = doc.Paragraphs.Item(startPara - 1); }
        catch (Exception) { return; }
        string text = CleanPara(Convert.ToString(TryGet(() => (object)prev.Range.Text, ""), CultureInfo.InvariantCulture));
        if (text.IndexOf('\f') >= 0) pageBreakChar = true;
        pageBreakBefore = Convert.ToInt32(TryGet(() => (object)(int)prev.Range.ParagraphFormat.PageBreakBefore, (object)0), CultureInfo.InvariantCulture) != 0;
        // NOTE (hang post-mortem 2026-11-03): do NOT look for <w:sectPr in prev.Range.WordOpenXML.
        // Range.WordOpenXML embeds the section properties of the RANGE, so a perfectly ordinary
        // paragraph in a single-section document matched it -> every table looked like
        // "section break in front" -> the cut-paste method ran first -> WPS Writer spun at 100% CPU
        // forever on the first caption move. Section identity is decided by the section index below.
        if (!sectPrBefore && secIndex > 0)
        {
            int prevSec = Convert.ToInt32(TryGet(() => (object)(int)prev.Range.Sections.Item(1).Index, (object)0), CultureInfo.InvariantCulture);
            if (prevSec > 0 && prevSec != secIndex) sectPrBefore = true;
        }
    }

    // ============================================================ table -Action plan

    /// <summary>
    /// table -Action plan: everything an agent has to know BEFORE touching the tables, in one call:
    /// per table page span / rows / columns / caption side / borders / table style / grid columns /
    /// risks, the caption list with its numbering issues, the in-text reference map, and a purely
    /// mechanical renumbering proposal (never applied without an approved plan file).
    /// Read-only: no undo record, no write, works on a read-only document.
    /// </summary>
    private static object TabPlan(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (total == 0) throw new Fail("BAD_ARGS", "the document has no table; nothing to plan.");
        int only = a.Int("table", 0);
        if (only > 0 && only > total) throw new Fail("BAD_ARGS", "table " + only + " does not exist (the document has " + total + ").");

        dynamic app = App(true, false);
        int selStart = -1, selEnd = -1;
        try { selStart = (int)app.Selection.Start; selEnd = (int)app.Selection.End; } catch (Exception) { }
        int paraTotal = 0;
        try
        {
            try { doc.Repaginate(); } catch (Exception) { }
            paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            double textWidth = TextWidth(doc);

            string xml = WholeOpenXml(doc);
            List<TableXmlInfo> infos = xml == null ? new List<TableXmlInfo>() : ParseTableXml(xml, false);

            // The caption list comes from the captions command itself - one parser, one answer.
            Dictionary<string, object> caps = null;
            try { caps = CaptionsDoc(a) as Dictionary<string, object>; }
            catch (Exception e) { Tr("plan: captions failed: " + Clean(e.Message)); }
            Dictionary<int, Dictionary<string, object>> capsByTable = new Dictionary<int, Dictionary<string, object>>();
            List<object> captionRows = new List<object>();
            List<object> references = new List<object>();
            List<object> issues = new List<object>();
            if (caps != null)
            {
                List<object> rowsRaw = caps["tables"] as List<object>;
                if (rowsRaw != null)
                {
                    foreach (object o in rowsRaw)
                    {
                        Dictionary<string, object> row = o as Dictionary<string, object>;
                        if (row == null) continue;
                        int tn = Convert.ToInt32(row["table"], CultureInfo.InvariantCulture);
                        Dictionary<string, object> cap = row["caption"] as Dictionary<string, object>;
                        if (cap != null) capsByTable[tn] = cap;
                        captionRows.Add(row);
                    }
                }
                List<object> refRaw = caps["references"] as List<object>;
                if (refRaw != null) references = refRaw;
                List<object> issRaw = caps["issues"] as List<object>;
                if (issRaw != null) issues = issRaw;
            }

            List<object> tables = new List<object>();
            List<object> risks = new List<object>();
            int capRows = 0, noThick = 0, breaks = 0, noStyle = 0, noBorders = 0, pctWidths = 0;
            int multipage = 0, headerMissing = 0, gridOver = 0, widthsUnknown = 0;
            int moveRisks = 0, layoutRisks = 0;

            for (int n = 1; n <= total; n++)
            {
                if (only > 0 && n != only) continue;
                Tr("plan table " + n);
                dynamic table = doc.Tables.Item(n);
                TableXmlInfo info = (n - 1) < infos.Count ? infos[n - 1] : null;
                int startPara = TableStartPara(doc, table, paraTotal);
                int[] span = PageSpanOf(doc, table);
                bool split = span[1] > span[0] && span[0] > 0;
                int heading = ReadHeadingFormat(app, table);
                int[] borders = BorderPair(table);
                bool thick = ThickBorders(borders);
                string firstCell = CleanPara(Convert.ToString(TryGet(() => (object)table.Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                bool capInRow = IsCaptionText(firstCell);

                bool atSectionStart, sectPrBefore, pageBreakChar, pageBreakBefore;
                DetectBreakBefore(doc, table, startPara, out atSectionStart, out sectPrBefore, out pageBreakChar, out pageBreakBefore);
                bool before = atSectionStart || sectPrBefore || pageBreakChar || pageBreakBefore;

                Dictionary<string, object> capOut = null;
                Dictionary<string, object> capEntry = null;
                capsByTable.TryGetValue(n, out capEntry);
                if (capEntry != null)
                {
                    capOut = new Dictionary<string, object>();
                    capOut["pos"] = capEntry["pos"];
                    capOut["para"] = capEntry["para"];
                    int capPara = Convert.ToInt32(capEntry["para"], CultureInfo.InvariantCulture);
                    capOut["page"] = PageOfPara(doc, capPara);
                    string number = Convert.ToString(capEntry["number"], CultureInfo.InvariantCulture);
                    string title = Convert.ToString(capEntry["title"], CultureInfo.InvariantCulture);
                    string label = Convert.ToString(capEntry["label"], CultureInfo.InvariantCulture);
                    capOut["text"] = Head((label + number + " " + title).Trim(), 60);
                    if (label.Length > 0) capOut["label"] = label;
                }

                Dictionary<string, object> row = new Dictionary<string, object>();
                row["table"] = n;
                row["para"] = startPara;
                row["pages"] = span[0] == span[1]
                    ? (object)span[0]
                    : (object)(span[0].ToString(CultureInfo.InvariantCulture) + "-" + span[1].ToString(CultureInfo.InvariantCulture));
                row["rows"] = TryGet(() => (object)(int)table.Rows.Count, 0);
                row["columns"] = TryGet(() => (object)(int)table.Columns.Count, 0);
                if (capOut != null) row["caption"] = capOut; else row["caption"] = null;
                row["split"] = split;
                row["headerRepeated"] = heading == -1;
                row["borders"] = new List<object>(new object[] { borders[0], borders[1] });
                row["tblStyle"] = info == null ? "?" : info.Style;
                row["tblW"] = info == null ? "?" : (info.WidthType + "/" + Convert.ToString(Math.Round(info.Width, 1), CultureInfo.InvariantCulture));
                row["gridCols"] = info == null ? null : (object)GridColsPt(info);
                double fontSize = Num((object)TryGet(() => (object)table.Range.Font.Size, (object)0), 0);
                if (fontSize > 100) fontSize = 0;          // 9999999 = wdUndefined
                row["fontSize"] = Math.Round(fontSize, 1);

                List<string> risk = new List<string>();
                if (capInRow) { risk.Add("captionInFirstRow"); capRows++; }
                if (!thick) { risk.Add("bordersNotThick"); noThick++; }
                if (before) { risk.Add("breakBefore"); breaks++; }
                if (info != null)
                {
                    if (info.Style.Length == 0) { risk.Add("noStyle"); noStyle++; }
                    if (!info.HasBorders) { risk.Add("noBorders"); noBorders++; }
                    if (info.WidthType == "pct") { risk.Add("pctWidths"); pctWidths++; }
                    double gridPt = GridSumPt(info);
                    if (gridPt > 0 && gridPt > textWidth * 1.05) { risk.Add("gridOverflow"); gridOver++; }
                }
                else widthsUnknown++;
                if (split) { multipage++; if (heading != -1) { risk.Add("multipageWithoutHeaderRepeat"); headerMissing++; } }
                // A caption in row 1 that sits behind a page/section break is the expensive case:
                // inserting the caption the naive way leaves it on the PREVIOUS page (session bug 1).
                if (capInRow && before)
                {
                    string cr = (sectPrBefore || atSectionStart) ? "section-break" : "page-break";
                    row["captionMoveRisk"] = cr;
                    // user decision 2026-11-03: a caption behind a section break is NOT moved on WPS
                    if (cr == "section-break") row["moveOut"] = "unsupported-on-wps";
                    risk.Add("captionMoveRisk");
                    moveRisks++;
                }
                // LAYOUT_RISK: writing borders into a table with no style or with a grid wider than the
                // text column is where "fold every line / shrink the font" decisions get made by hand.
                if (info != null && (info.Style.Length == 0 || GridSumPt(info) > textWidth * 1.05))
                {
                    row["layoutRisk"] = "LAYOUT_RISK";
                    risk.Add("layoutRisk");
                    layoutRisks++;
                }
                if (risk.Count > 0)
                {
                    Dictionary<string, object> r = new Dictionary<string, object>();
                    r["table"] = n;
                    r["flags"] = new List<object>(risk.ToArray());
                    risks.Add(r);
                }
                row["risks"] = new List<object>(risk.ToArray());
                tables.Add(row);
            }

            Dictionary<string, object> summary = new Dictionary<string, object>();
            summary["tables"] = only > 0 ? 1 : total;
            summary["captions"] = captionRows.Count;
            summary["captionInFirstRow"] = capRows;
            summary["bordersNotThick"] = noThick;
            summary["noStyle"] = noStyle;
            summary["noBorders"] = noBorders;
            summary["pctWidths"] = pctWidths;
            summary["breakBefore"] = breaks;
            summary["multipage"] = multipage;
            summary["multipageWithoutHeaderRepeat"] = headerMissing;
            summary["gridOverflow"] = gridOver;
            summary["captionMoveRisk"] = moveRisks;
            summary["layoutRisk"] = layoutRisks;
            summary["textWidth"] = Math.Round(textWidth, 1);
            if (widthsUnknown > 0) summary["gridColsUnknown"] = widthsUnknown;

            Dictionary<string, object> r2 = new Dictionary<string, object>();
            r2["action"] = "plan";
            r2["tables"] = tables;
            r2["risks"] = risks;
            r2["captions"] = captionRows;
            r2["references"] = references;
            r2["issues"] = issues;
            r2["renumber"] = RenumberFromCaptions(captionRows, references);
            r2["summary"] = summary;
            r2["paragraphs"] = paraTotal;
            r2["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
            r2["saved"] = false;
            r2["note"] = "read-only (no undo record, nothing written). gridCols come from ONE WordOpenXML "
                + "fetch: a table flagged noBorders is exactly the one whose column widths WPS rebuilds "
                + "when borders are first written - normalize snapshots those and restores them. "
                + "renumber is only a proposal: copy it into a plan file, check it, set approved:true, "
                + "then run table -Action normalize -Plan <file>.";
            return r2;
        }
        finally { RestoreSelection(doc, selStart, selEnd); }
    }

    /// <summary>
    /// DESIGN 3.3: a MECHANICAL renumbering proposal. Within one number prefix the captions get
    /// prefix.1, prefix.2, ... in document order; a number used twice is marked unique:false and
    /// blocks automatic in-text reference rewriting (ambiguity is never guessed away).
    /// </summary>
    private static Dictionary<string, object> RenumberFromCaptions(List<object> captionRows, List<object> references)
    {
        List<Dictionary<string, object>> caps = new List<Dictionary<string, object>>();
        foreach (object o in captionRows)
        {
            Dictionary<string, object> row = o as Dictionary<string, object>;
            if (row == null) continue;
            Dictionary<string, object> cap = row["caption"] as Dictionary<string, object>;
            if (cap == null) continue;
            caps.Add(row);
        }

        // How often the body text points at each number - the second half of the "basis" a human
        // needs when confirming the proposal (as the LEAD asked for: judgement, not just numbers).
        Dictionary<string, int> refCounts = new Dictionary<string, int>();
        if (references != null)
        {
            foreach (object o in references)
            {
                Dictionary<string, object> refRow = o as Dictionary<string, object>;
                if (refRow == null) continue;
                string num = Convert.ToString(refRow["number"], CultureInfo.InvariantCulture);
                if (num.Length == 0) continue;
                refCounts[num] = Convert.ToInt32(refRow["count"], CultureInfo.InvariantCulture);
            }
        }

        Dictionary<string, int> numberCount = new Dictionary<string, int>();
        foreach (Dictionary<string, object> row in caps)
        {
            Dictionary<string, object> cap = (Dictionary<string, object>)row["caption"];
            string number = Convert.ToString(cap["number"], CultureInfo.InvariantCulture);
            int c;
            numberCount.TryGetValue(number, out c);
            numberCount[number] = c + 1;
        }

        List<string> order = new List<string>();
        Dictionary<string, List<Dictionary<string, object>>> groups = new Dictionary<string, List<Dictionary<string, object>>>();
        foreach (Dictionary<string, object> row in caps)
        {
            Dictionary<string, object> cap = (Dictionary<string, object>)row["caption"];
            string number = Convert.ToString(cap["number"], CultureInfo.InvariantCulture);
            string prefix = NumberPrefix(number);
            List<Dictionary<string, object>> list;
            if (!groups.TryGetValue(prefix, out list)) { list = new List<Dictionary<string, object>>(); groups[prefix] = list; order.Add(prefix); }
            list.Add(row);
        }

        List<object> items = new List<object>();
        Dictionary<string, string> fromTo = new Dictionary<string, string>();
        Dictionary<string, bool> fromUnique = new Dictionary<string, bool>();
        Dictionary<string, List<string>> proposedTo = new Dictionary<string, List<string>>();
        int changed = 0, duplicates = 0, ambiguous = 0;
        foreach (string prefix in order)
        {
            List<Dictionary<string, object>> list = groups[prefix];
            for (int i = 0; i < list.Count; i++)
            {
                Dictionary<string, object> cap = (Dictionary<string, object>)list[i]["caption"];
                string number = Convert.ToString(cap["number"], CultureInfo.InvariantCulture);
                string format = Convert.ToString(cap["format"], CultureInfo.InvariantCulture);
                string sep = format == "hyphen" ? "-" : (format == "plain" ? "" : ".");
                string to = prefix.Length == 0 ? (i + 1).ToString(CultureInfo.InvariantCulture) : (prefix + sep + (i + 1).ToString(CultureInfo.InvariantCulture));
                bool unique = numberCount.ContainsKey(number) && numberCount[number] == 1;
                if (unique && !fromTo.ContainsKey(number)) { fromTo[number] = to; fromUnique[number] = true; }
                else if (fromTo.ContainsKey(number)) { fromUnique[number] = false; }
                if (!unique) duplicates++;
                if (to != number) changed++;

                Dictionary<string, object> item = new Dictionary<string, object>();
                item["table"] = list[i]["table"];
                item["para"] = list[i]["para"];
                item["from"] = number;
                item["to"] = to;
                item["unique"] = unique;
                int refCount = 0;
                refCounts.TryGetValue(number, out refCount);
                item["basis"] = refCount > 0 ? "document-order+refs" : "document-order";
                if (refCount > 0) item["refs"] = refCount;
                List<string> targets;
                if (!proposedTo.TryGetValue(number, out targets)) { targets = new List<string>(); proposedTo[number] = targets; }
                targets.Add(to);
                items.Add(item);
            }
        }

        List<object> refs = new List<object>();
        if (references != null)
        {
            foreach (object o in references)
            {
                Dictionary<string, object> refRow = o as Dictionary<string, object>;
                if (refRow == null) continue;
                string num = Convert.ToString(refRow["number"], CultureInfo.InvariantCulture);
                bool unique;
                fromUnique.TryGetValue(num, out unique);
                Dictionary<string, object> refOut = new Dictionary<string, object>();
                refOut["number"] = num;
                refOut["count"] = refRow["count"];
                if (unique && fromTo.ContainsKey(num))
                {
                    refOut["to"] = fromTo[num];
                    refOut["ambiguous"] = false;
                }
                else
                {
                    refOut["ambiguous"] = true;
                    List<string> targets;
                    if (proposedTo.TryGetValue(num, out targets))
                        refOut["candidates"] = new List<object>(targets.ToArray());
                    ambiguous++;
                }
                refs.Add(refOut);
            }
        }

        Dictionary<string, object> renumber = new Dictionary<string, object>();
        renumber["approved"] = false;
        renumber["items"] = items;
        renumber["refs"] = refs;
        renumber["changed"] = changed;
        renumber["duplicates"] = duplicates;
        renumber["ambiguous"] = ambiguous;
        renumber["note"] = "proposal only. Set approved:true (and edit items/refs if the numbering is "
            + "not just 'renumber in document order per prefix') and pass it back with -Plan. A number "
            + "used twice stays ambiguous and is never rewritten automatically.";
        return renumber;
    }

    // ============================================================ table -Action normalize
    //
    // One command, per-table isolation. Order inside the command (DESIGN 2.2): caption out of row 1
    // -> caption spacing -> (optional) renumber -> thick borders with a grid snapshot/restore ->
    // header repeat -> verification sweep. Everything happens inside ONE custom undo record, so one
    // Ctrl+Z takes the whole command back.

    private const string WORD_NS =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
        + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" "
        + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
        + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" "
        + "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ";

    /// <summary>A Flat OPC package holding one &lt;w:tbl&gt; chunk, for InsertXML (WPS only takes tables back as a package).</summary>
    private static string FlatTablePackage(string tableXml)
    {
        return "<?xml version=\"1.0\" standalone=\"yes\"?>"
            + "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
            + "<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\">"
            + "<pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships></pkg:xmlData></pkg:part>"
            + "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\">"
            + "<pkg:xmlData><w:document " + WORD_NS + "><w:body>" + tableXml + "<w:p/></w:body></w:document></pkg:xmlData></pkg:part>"
            + "</pkg:package>";
    }

    /// <summary>
    /// True when the COM host is WPS (it registers the Word.Application ProgID and reports
    /// Name="Microsoft Word" with Version 12.0 / a Kingsoft path - Name is useless as a test).
    /// </summary>
    private static bool IsWpsApp(dynamic app)
    {
        string path = Convert.ToString(TryGet(() => (object)app.Path, ""), CultureInfo.InvariantCulture);
        if (path != null && path.IndexOf("Kingsoft", StringComparison.OrdinalIgnoreCase) >= 0) return true;
        string version = Convert.ToString(TryGet(() => (object)app.Version, ""), CultureInfo.InvariantCulture);
        if (version == null) return false;
        return version.StartsWith("12.0", StringComparison.Ordinal) || version.StartsWith("11.", StringComparison.Ordinal);
    }

    private sealed class TableWork
    {
        public int Index;
        public int AnchorEnd;      // char position where the table started (anchor paragraph's end)
        public string SavedXml = "";   // the original <w:tbl> chunk, for the XML rollback
        public int InsertedPara;       // the caption paragraph we created, if any
        public int RowsBefore;         // row count when the table work started (damage detection)
        public int ParasBefore;        // paragraph count when the table work started
        public int TablesBefore;       // table count when the table work started
    }

    private sealed class CaptionMoveResult
    {
        public bool Ok;
        public int NewPara;
        public int RowsBefore;
        public int RowsAfter;
        public int Page;
        public string Method = "";
        public string TextBefore = "";
        public string TextAfter = "";
        public string Error = "";
        public bool NeedsHuman;        // a permanent host limitation, not a bug (caption left inside)
        public bool Incomplete;        // the failed attempt could not be undone: stop the batch
    }

    private static void RemoveParagraph(dynamic doc, int para)
    {
        try
        {
            int total = (int)doc.Paragraphs.Count;
            if (para < 1 || para > total) return;
            // NEVER delete a paragraph that lives inside a table: the WPS paragraph numbering puts
            // table rows into Paragraphs, so a mis-addressed "caption paragraph" index here would
            // tear a table apart (measured 2026-11-03: it cost 3 paragraphs and 2 pages on table 3).
            if (TabOfParagraph(doc, para) > 0) { Tr("refusing to remove paragraph " + para + ": it is inside a table"); return; }
            dynamic p = doc.Paragraphs.Item(para);
            int s = (int)p.Range.Start;
            int e = (int)p.Range.End;
            if (e > s) doc.Range(s, e).Delete();
        }
        catch (Exception e) { Tr("remove paragraph " + para + " failed: " + Clean(e.Message)); }
    }

    /// <summary>
    /// Find the paragraph that now holds the caption: it must be OUTSIDE the table and start with the
    /// caption text. Looking up the number the insert reported is not enough - after an insert at the
    /// table's start that index resolves to a paragraph INSIDE the table.
    /// </summary>
    private static int LocateCaptionParagraph(dynamic doc, int n, string caption)
    {
        int probe = Math.Min(8, caption.Trim().Length);
        if (probe < 2) return 0;
        string head = caption.Trim().Substring(0, probe);
        int paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        int startPara = TableStartPara(doc, doc.Tables.Item(n), paraTotal);
        for (int k = startPara - 1; k >= 1 && k >= startPara - 5; k--)
        {
            string text = ParaTextAt(doc, k).Trim();
            if (text.Length < 2) continue;
            if (!text.StartsWith(head, StringComparison.Ordinal)) continue;
            if (TabOfParagraph(doc, k) > 0) continue;
            return k;
        }
        return 0;
    }

    /// <summary>Give the new caption paragraph the look of the caption cell it came from; keep it with the table.</summary>
    private static void StyleCaptionParagraph(dynamic doc, int para, string fontName, string fontFe, double fontSize, int align)
    {
        try
        {
            dynamic np = doc.Paragraphs.Item(para);
            if (fontName != null && fontName.Length > 0)
            {
                try { np.Range.Font.Name = fontName; } catch (Exception) { }
                try { if (fontFe != null && fontFe.Length > 0) np.Range.Font.NameFarEast = fontFe; } catch (Exception) { }
            }
            if (fontSize > 0) { try { np.Range.Font.Size = (float)fontSize; } catch (Exception) { } }
            try
            {
                np.Format.Alignment = align;
                np.Format.FirstLineIndent = 0;
                np.Format.LeftIndent = 0;
                np.Format.RightIndent = 0;
                np.Format.KeepWithNext = -1;      // a caption must never be split from its table
                np.Format.PageBreakBefore = 0;
            }
            catch (Exception) { }
        }
        catch (Exception e) { Tr("style caption failed: " + Clean(e.Message)); }
    }

    /// <summary>
    /// A caption is only accepted when the page number reads back equal to the table's first page
    /// (DESIGN 3.1 step 5). This is what the naive implementation of 2026-11-03 never checked.
    /// </summary>
    private static bool VerifyCaptionParagraph(dynamic doc, int n, int newPara, string caption, out string error)
    {
        error = null;
        if (newPara < 1) { error = "no caption paragraph was created"; return false; }
        int inside = TabOfParagraph(doc, newPara);
        if (inside > 0) { error = "the caption paragraph landed inside table " + inside; return false; }
        string text = ParaTextAt(doc, newPara).Trim();
        int probe = Math.Min(6, caption.Length);
        if (probe > 0 && text.IndexOf(caption.Substring(0, probe), StringComparison.Ordinal) < 0)
        {
            error = "the caption paragraph does not hold the caption text (got \"" + Head(text, 30) + "\")";
            return false;
        }
        try { doc.Repaginate(); } catch (Exception) { }
        int[] span = PageSpanOf(doc, doc.Tables.Item(n));
        int capPage = PageOfPara(doc, newPara);
        if (span[0] > 0 && capPage > 0 && capPage != span[0])
        {
            error = "the caption sits on page " + capPage + " but the table starts on page " + span[0];
            // Diagnostics: what the neighbouring paragraphs look like tells whether the caption went
            // into the PREVIOUS paragraph/section or just did not fit on the table's page.
            Tr("  verify fail: caption para " + newPara + " page " + capPage + " text=\"" + Head(text, 16) + "\"");
            int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            for (int k = newPara - 2; k <= newPara + 1; k++)
            {
                if (k < 1 || k > total) continue;
                int sec = Convert.ToInt32(TryGet(() => (object)(int)doc.Paragraphs.Item(k).Range.Sections.Item(1).Index, (object)0), CultureInfo.InvariantCulture);
                Tr("  para " + k + ": page " + PageOfPara(doc, k) + " section " + sec + " text=\"" + Head(ParaTextAt(doc, k).Trim(), 16) + "\"");
            }
            Tr("  table " + n + " start page " + span[0] + " end page " + span[1]);
            return false;
        }
        return true;
    }

    // ---------------------------------------------------------- the three insertion methods

    /// <summary>
    /// Method A (DESIGN 3.1 step 2): InsertParagraphBefore at (table start - 1), then write the text.
    /// The page/section break stays IN FRONT of the new paragraph, so the caption belongs to the
    /// table's page instead of the previous one.
    /// </summary>
    private static bool InsertCaptionTableStart(dynamic doc, int n, string caption, out int newPara, out string error)
    {
        newPara = 0;
        error = null;
        int start = Convert.ToInt32(TryGet(() => (object)(int)doc.Tables.Item(n).Range.Start, (object)-1), CultureInfo.InvariantCulture);
        if (start < 1) { error = "table start unknown"; return false; }
        doc.Range(start - 1, start - 1).InsertParagraphBefore();
        int start2 = Convert.ToInt32(TryGet(() => (object)(int)doc.Tables.Item(n).Range.Start, (object)-1), CultureInfo.InvariantCulture);
        if (start2 < 2) { error = "table start unknown after the paragraph insert"; return false; }
        dynamic np = doc.Range(start2 - 1, start2 - 1).Paragraphs.Item(1);
        np.Range.InsertBefore(caption);
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        newPara = ParaAtOffset(doc, start2 - 2, total);
        if (newPara < 1) { error = "could not address the new caption paragraph"; return false; }
        return true;
    }

    /// <summary>Method B (pre-v1.33 behaviour, kept as the fallback): insert at the previous paragraph's mark.</summary>
    private static bool InsertCaptionBeforePara(dynamic doc, int n, int startPara, string caption, out int newPara, out string error)
    {
        newPara = 0;
        error = null;
        if (startPara <= 1) { error = "the table has no paragraph in front of it"; return false; }
        int tableStart = (int)TryGet(() => (object)(int)doc.Tables.Item(n).Range.Start, 0);
        if (tableStart < 2) { error = "table start unknown"; return false; }
        int prevCount = (int)TryGet(() => (object)(int)doc.Range(0, tableStart - 1).Paragraphs.Count, 0);
        if (prevCount < 1) { error = "no paragraph before the table"; return false; }
        int pos = (int)doc.Paragraphs.Item(prevCount).Range.End - 1;
        doc.Range(pos, pos).InsertAfter("\r" + caption);
        newPara = prevCount + 1;
        return true;
    }

    /// <summary>
    /// Method A2 (DESIGN 3.1 step 2 verbatim): Table.Range.InsertParagraphBefore(). Word knows that a
    /// paragraph inserted before a table range belongs to the table's section, so this is the
    /// primitive to try FIRST when the table starts a section. The v1.33 first attempt used
    /// Range(start-1).InsertParagraphBefore() (the reference implementation's spelling) which splits
    /// the PREVIOUS paragraph - and a sectPr stays with the last paragraph of its section, so the new
    /// paragraph ended up in the PREVIOUS section (measured: caption page 1, table page 2).
    /// </summary>
    private static bool InsertCaptionTablePara(dynamic doc, int n, string caption, out int newPara, out string error)
    {
        newPara = 0;
        error = null;
        dynamic table = null;
        try { table = doc.Tables.Item(n); }
        catch (Exception e) { error = "table not reachable: " + Clean(e.Message); return false; }
        int start = Convert.ToInt32(TryGet(() => (object)(int)table.Range.Start, (object)-1), CultureInfo.InvariantCulture);
        if (start < 1) { error = "table start unknown"; return false; }
        try { table.Range.InsertParagraphBefore(); }
        catch (Exception e) { error = "InsertParagraphBefore refused: " + Clean(e.Message); return false; }
        int start2 = Convert.ToInt32(TryGet(() => (object)(int)doc.Tables.Item(n).Range.Start, (object)-1), CultureInfo.InvariantCulture);
        if (start2 < 2) { error = "table start unknown after the insert"; return false; }
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        newPara = ParaAtOffset(doc, start2 - 2, total);
        if (newPara < 1) { error = "could not address the new caption paragraph"; return false; }
        dynamic np = doc.Paragraphs.Item(newPara);
        np.Range.InsertBefore(caption);
        return true;
    }



    /// <summary>
    /// Method C (DESIGN 3.1 step 4): write "caption + paragraph mark" AT the table's start position.
    /// Intended for a table that BEGINS ITS SECTION - the text lands after the section break (which
    /// belongs to the previous paragraph), so the caption belongs to the table's section and page.
    /// Does so WITHOUT the reference implementation's save/delete/InsertXML dance, which hangs WPS
    /// Writer 12.1 (measured 2026-11-03, see MoveCaptionOutOfTable).
    /// </summary>
    private static bool InsertCaptionRangeInsert(dynamic doc, int n, string caption, out int newPara, out string error)
    {
        newPara = 0;
        error = null;
        int tableStart = Convert.ToInt32(TryGet(() => (object)(int)doc.Tables.Item(n).Range.Start, (object)-1), CultureInfo.InvariantCulture);
        if (tableStart < 1) { error = "table start unknown"; return false; }
        doc.Range(tableStart, tableStart).InsertBefore(caption + "\r");
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        newPara = ParaAtOffset(doc, tableStart, total);
        if (newPara < 1) { error = "could not address the new caption paragraph"; return false; }
        return true;
    }

    /// <summary>
    /// Method D (DESIGN 3.1 / reference TableTitleExtractor MoveTitleViaCutPaste): the table starts
    /// its section, so save it as Flat OPC, delete it, write "caption + paragraph mark" at the
    /// section start, put the table back with InsertXML.
    /// DISABLED BY DEFAULT: on WPS 12.1 the first caption move through this path froze the Writer
    /// process (trace: "caption move table 1 via cut-paste" and then nothing for 2.5 minutes).
    /// Only tried when the caller passes -CutPaste or when the host is Microsoft Word.
    /// </summary>
    private static bool InsertCaptionCutPaste(dynamic doc, int n, string caption, out int newPara, out string error)
    {
        newPara = 0;
        error = null;
        int secIndex = 0, rowsBefore = 0, tableStart = 0;
        string xml = null;
        try
        {
            dynamic table = doc.Tables.Item(n);
            secIndex = (int)table.Range.Sections.Item(1).Index;
            rowsBefore = (int)table.Rows.Count;
            tableStart = (int)table.Range.Start;
            xml = Convert.ToString(table.Range.WordOpenXML, CultureInfo.InvariantCulture);
        }
        catch (Exception e) { error = "WordOpenXML refused: " + Clean(e.Message); return false; }
        if (xml == null || xml.Length == 0) { error = "WordOpenXML came back empty"; return false; }

        try { doc.Tables.Item(n).Delete(); }
        catch (Exception e) { error = "WPS refused to delete the table: " + Clean(e.Message); return false; }

        int secStart = -1;
        try { secStart = (int)doc.Sections.Item(secIndex).Range.Start; }
        catch (Exception e)
        {
            error = "the section disappeared when the table was deleted: " + Clean(e.Message);
            RecoverTableXml(doc, tableStart, xml);
            return false;
        }

        try
        {
            doc.Range(secStart, secStart).InsertBefore(caption + "\r");
            int at = secStart + caption.Length + 1;
            doc.Range(at, at).InsertXML(xml);
        }
        catch (Exception e)
        {
            error = "writing the caption / putting the table back failed: " + Clean(e.Message);
            try { doc.Range(secStart, secStart + caption.Length + 1).Delete(); } catch (Exception) { }
            RecoverTableXml(doc, secStart, xml);
            return false;
        }

        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (total < n)
        {
            error = "the table did not come back after InsertXML";
            try { doc.Range(secStart, secStart + caption.Length + 1).Delete(); } catch (Exception) { }
            RecoverTableXml(doc, secStart, xml);
            return false;
        }
        int rowsBack = Convert.ToInt32(TryGet(() => (object)(int)doc.Tables.Item(n).Rows.Count, (object)0), CultureInfo.InvariantCulture);
        if (rowsBack != rowsBefore)
        {
            error = "the table came back with " + rowsBack + " rows instead of " + rowsBefore;
        }
        int paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        newPara = ParaAtOffset(doc, secStart, paraTotal);
        if (newPara < 1) { error = "could not address the caption paragraph"; return false; }
        return error == null;
    }

    /// <summary>Last-ditch restore of the table XML (used when the cut-paste path lost the table).</summary>
    private static void RecoverTableXml(dynamic doc, int at, string xml)
    {
        if (xml == null || xml.Length == 0) return;
        try
        {
            if (at < 0) at = 0;
            doc.Range(at, at).InsertXML(FlatTablePackage(xml));
            Tr("recovered table xml at " + at);
        }
        catch (Exception e) { Tr("recover table xml failed: " + Clean(e.Message)); }
    }

    /// <summary>
    /// table -Action normalize's caption step: move the caption typed into row 1 out of the table,
    /// trying the three methods in the order the document's structure suggests and keeping the first
    /// one whose page number reads back correct. The row that held the caption is deleted last.
    /// </summary>
    private static CaptionMoveResult MoveCaptionOutOfTable(dynamic doc, dynamic app, int n, bool allowCutPaste)
    {
        CaptionMoveResult res = new CaptionMoveResult();
        dynamic table = null;
        try { table = doc.Tables.Item(n); }
        catch (Exception e) { res.Error = "table " + n + " is not reachable: " + Clean(e.Message); return res; }

        res.RowsBefore = Convert.ToInt32(TryGet(() => (object)(int)table.Rows.Count, (object)0), CultureInfo.InvariantCulture);
        if (res.RowsBefore < 2) { res.Error = "table " + n + " has only " + res.RowsBefore + " row(s)"; return res; }

        string fontName = "", fontFe = "";
        double fontSize = 0;
        int align = 1;
        res.TextBefore = "";
        try
        {
            dynamic firstCell = table.Cell(1, 1);
            res.TextBefore = CleanPara(Convert.ToString(firstCell.Range.Paragraphs.Item(1).Range.Text, CultureInfo.InvariantCulture)).Trim();
            fontName = Convert.ToString(TryGet(() => (object)firstCell.Range.Font.Name, ""), CultureInfo.InvariantCulture);
            fontFe = Convert.ToString(TryGet(() => (object)firstCell.Range.Font.NameFarEast, ""), CultureInfo.InvariantCulture);
            fontSize = Num(TryGet(() => (object)firstCell.Range.Font.Size, (object)0), 0);
            align = Convert.ToInt32(TryGet(() => (object)(int)firstCell.Range.ParagraphFormat.Alignment, (object)1), CultureInfo.InvariantCulture);
        }
        catch (Exception e) { res.Error = "the first cell is not readable: " + Clean(e.Message); return res; }
        string caption = res.TextBefore;
        if (caption.Length == 0) { res.Error = "the first cell of table " + n + " is empty"; return res; }

        int paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        int startPara = TableStartPara(doc, table, paraTotal);
        bool atSectionStart, sectPrBefore, pageBreakChar, pageBreakBefore;
        DetectBreakBefore(doc, table, startPara, out atSectionStart, out sectPrBefore, out pageBreakChar, out pageBreakBefore);
        bool sectionInFront = sectPrBefore || atSectionStart;
        bool breakInFront = sectionInFront || pageBreakChar || pageBreakBefore;
        Tr("caption move table " + n + ": break in front = " + (sectionInFront ? "section" : (breakInFront ? "page" : "none")));

        // Method order (hang post-mortem 2026-11-03): the cut-paste method (delete the table, put it
        // back with InsertXML) made WPS Writer spin at 100% CPU forever on the very first try, so it
        // is NOT attempted on WPS unless the caller explicitly asks (-CutPaste). A caption we cannot
        // move is reported as needsHuman; a frozen WPS is a much worse outcome.
        // Method order (measured on WPS 12.1, 2026-11-03):
        //   table-para  Table.Range.InsertParagraphBefore() - DESIGN 3.1 method A verbatim. It is the
        //               ONLY spelling that works when the table begins a section: Word puts the new
        //               paragraph in the TABLE's section (the sectPr stays with the previous
        //               paragraph), so the caption and the table share a page. Measured to pass the
        //               page check on a section-break fixture in one try.
        //   range-insert / table-start / before-para are kept as fallbacks for shapes where the table
        //               range is not addressable or the host refuses the insert.
        //   cut-paste   only with -CutPaste / a Microsoft Word host: it freezes WPS Writer.
        // Method order (measured on WPS 12.1, 2026-11-03):
        //   a table that BEGINS ITS SECTION (the section-break mark sits in the paragraph right in
        //   front of it) cannot get a NEW paragraph in its own section: every insert lands in the
        //   previous section (trace: caption page 8 / section 1, table page 9 / section 2). Per the
        //   user's decision of 2026-11-03 this is NOT attempted at all any more - the table is
        //   reported as needsHuman, so the document can never be damaged and WPS can never be
        //   driven into a state we cannot undo.
        List<string> methods = new List<string>();
        if (sectionInFront)
        {
            res.NeedsHuman = true;
            res.Error = "wps-section-start: the section-break mark sits in the paragraph right in front of "
                + "this table, so WPS puts every inserted paragraph into the PREVIOUS section";
            Tr("caption move table " + n + ": section-start table -> needsHuman (not attempted)");
            return res;
        }
        // Only the spelling the field measurements proved safe on WPS: insert at the previous
        // paragraph's mark (lead's probe 2026-11-03: plain 2.0 s, manual page break 1.16 s, both OK).
        // table-para / range-insert / table-start were measured to leave the caption inside the table
        // or in the previous section, and cleaning that up damaged the document - so they are gone.
        methods.Add("before-para");
        if (allowCutPaste) methods.Add("cut-paste");

        string lastError = null;
        foreach (string method in methods)
        {
            Tr("caption move table " + n + " via " + method);
            int newPara = 0;
            string err = null;
            bool inserted = false;
            int parasBefore = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            try
            {
                if (method == "cut-paste") inserted = InsertCaptionCutPaste(doc, n, caption, out newPara, out err);
                else if (method == "table-start") inserted = InsertCaptionTableStart(doc, n, caption, out newPara, out err);
                else if (method == "table-para") inserted = InsertCaptionTablePara(doc, n, caption, out newPara, out err);
                else if (method == "range-insert") inserted = InsertCaptionRangeInsert(doc, n, caption, out newPara, out err);
                else inserted = InsertCaptionBeforePara(doc, n, startPara, caption, out newPara, out err);
            }
            catch (Exception e) { err = Clean(e.Message); inserted = false; }
            if (!inserted)
            {
                lastError = method + ": " + (err == null ? "failed" : err);
                Tr("caption move table " + n + " " + method + " failed: " + lastError);
                continue;
            }
            // Address the caption paragraph by TEXT, outside the table: the paragraph number the
            // insert reported may resolve to a cell paragraph (the WPS paragraph numbering counts
            // table rows), and removing THAT would tear the table apart.
            int capPara = newPara > 0 && TabOfParagraph(doc, newPara) == 0 && ParaTextAt(doc, newPara).Trim().StartsWith(Head(caption.Trim(), 8), StringComparison.Ordinal)
                ? newPara : LocateCaptionParagraph(doc, n, caption);
            if (capPara < 1)
            {
                lastError = method + ": the caption paragraph was not found outside the table";
                Tr("caption move table " + n + " " + method + " failed: " + lastError);
                if (!ParagraphCountStill(doc, parasBefore)) { res.Incomplete = true; res.Error = lastError; return res; }
                continue;
            }
            Tr("  " + method + ": capPara=" + capPara + " page=" + PageOfPara(doc, capPara)
                + " section=" + SectionOfPara(doc, capPara) + " tableSection=" + SectionOfPara(doc, startPara)
                + " text=\"" + Head(ParaTextAt(doc, capPara).Trim(), 14) + "\"");
            string verr = null;
            bool ok = false;
            // Style FIRST: KeepWithNext keeps the caption on the table's page when the table would
            // otherwise start a new page, and the page check below must see that final state.
            StyleCaptionParagraph(doc, capPara, fontName, fontFe, fontSize, align);
            try { ok = VerifyCaptionParagraph(doc, n, capPara, caption, out verr); }
            catch (Exception e) { verr = Clean(e.Message); ok = false; }
            if (!ok)
            {
                lastError = method + ": " + (verr == null ? "page check failed" : verr);
                Tr("caption move table " + n + " " + method + " verify failed: " + lastError);
                if (!RemoveCaptionParagraph(doc, capPara, parasBefore))
                {
                    res.Incomplete = true;
                    res.Error = "ROLLBACK_INCOMPLETE after " + lastError;
                    return res;
                }
                continue;
            }
            // Accepted. Drop the row that held the caption.
            res.NewPara = capPara;
            res.Method = method;
            res.TextAfter = ParaTextAt(doc, capPara).Trim();
            int rowsAfter = res.RowsBefore;
            try { rowsAfter = DeleteRowViaSelection(app, doc.Tables.Item(n), 1); }
            catch (Exception e) { res.Error = "the caption moved but WPS refused to delete row 1: " + Clean(e.Message); }
            res.RowsAfter = rowsAfter;
            if (rowsAfter != res.RowsBefore - 1)
            {
                if (res.Error == null) res.Error = "WPS did not delete row 1 (rows " + res.RowsBefore + " -> " + rowsAfter + ")";
                RemoveCaptionParagraph(doc, res.NewPara, parasBefore);
                res.NewPara = 0;
                res.Method = "";
                return res;
            }
            try { doc.Repaginate(); } catch (Exception) { }
            res.Page = PageOfPara(doc, res.NewPara);
            res.Ok = true;
            return res;
        }
        res.Error = (lastError == null ? "all insertion methods failed" : lastError);
        res.NeedsHuman = sectionInFront;
        return res;
    }

    private static bool ParagraphCountStill(dynamic doc, int before)
    {
        int now = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, -1);
        return now == before;
    }

    /// <summary>Remove the caption paragraph we inserted and prove the paragraph count came back.</summary>
    private static bool RemoveCaptionParagraph(dynamic doc, int para, int parasBefore)
    {
        try
        {
            int total = (int)doc.Paragraphs.Count;
            if (para >= 1 && para <= total && TabOfParagraph(doc, para) == 0)
            {
                dynamic p = doc.Paragraphs.Item(para);
                int s = (int)p.Range.Start;
                int e = (int)p.Range.End;
                if (e > s) doc.Range(s, e).Delete();
            }
        }
        catch (Exception e) { Tr("remove caption paragraph failed: " + Clean(e.Message)); }
        if (ParagraphCountStill(doc, parasBefore)) return true;
        Tr("paragraph count " + (int)TryGet(() => (object)(int)doc.Paragraphs.Count, -1)
            + " != " + parasBefore + " after removing the caption paragraph");
        return false;
    }

    // ---------------------------------------------------------- caption spacing (never per character)

    /// <summary>
    /// "表2.1.9虎山站..." -> "表 2.1.9  title" (label, one space, number, -Gap spaces, title) with ONE
    /// text write per caption paragraph. Per-character cleaning is what made the reference
    /// implementation 79x slower in its table phase, so it is deliberately not done here.
    /// </summary>
    private static bool NormalizeCaptionGap(dynamic doc, int para, int gap, out string before, out string after, out string skip)
    {
        before = ParaTextAt(doc, para).Trim();
        after = before;
        skip = null;
        CaptionInfo ci = ParseCaption(before);
        if (ci == null || ci.Number.Length == 0) { skip = "not a caption"; return false; }
        int fields = Convert.ToInt32(TryGet(() => (object)(int)doc.Paragraphs.Item(para).Range.Fields.Count, (object)0), CultureInfo.InvariantCulture);
        if (fields > 0) { skip = "caption-has-field"; return false; }
        string title = ci.Title.Trim();
        // Keep the label/number spacing the document already uses: "表2.1.1  xx" must NOT become
        // "表 2.1.1  xx" (2026-11-03: the first version did exactly that and rewrote the convention of
        // 12 captions; only the gap between the number and the title is what -Gap normalizes).
        int labelAt = before.IndexOf(ci.Label, StringComparison.Ordinal);
        bool labelSpace = labelAt >= 0 && labelAt + ci.Label.Length < before.Length
            && char.IsWhiteSpace(before[labelAt + ci.Label.Length]);
        string want = ci.Label + (labelSpace ? " " : "") + ci.Number;
        if (title.Length > 0) want = want + (gap > 0 ? new string(' ', gap) : " ") + title;
        after = want;
        if (want == before) return false;
        try
        {
            dynamic p = doc.Paragraphs.Item(para);
            int start = (int)p.Range.Start;
            string raw = Convert.ToString(p.Range.Text, CultureInfo.InvariantCulture);
            int len = CleanPara(raw).Length;
            if (len <= 0) { skip = "empty"; return false; }
            doc.Range(start, start + len).Text = want;
            return true;
        }
        catch (Exception e)
        {
            skip = Clean(e.Message);
            after = before;
            return false;
        }
    }

    // ---------------------------------------------------------- borders + grid restore

    /// <summary>
    /// DESIGN 3.2 (rewritten after the 2026-11-03 audit): write the borders only when the table is
    /// not already thick, then report before/after. The column widths are compared and restored
    /// from WordOpenXML snapshots by the caller - Column.Width reads back as 9999999 (wdUndefined)
    /// after WPS rebuilds a table layout, so COM is useless for that.
    /// </summary>
    private static void ApplyBordersKeepLayout(dynamic doc, dynamic table, string mode, TableXmlInfo info,
        Dictionary<string, object> outBorders, List<string> warnings, ref int applied, ref int skipped)
    {
        int[] before = BorderPair(table);
        Dictionary<string, object> b = new Dictionary<string, object>();
        b["before"] = new List<object>(new object[] { before[0], before[1] });
        // The short circuit (audit 2026-11-03): a table that already carries a 1.5 pt frame, 0.5 pt
        // rules AND an explicit <w:tblBorders> needs no write at all - and every write into a table
        // without that element is what makes WPS rebuild the layout (the width blow-up). The sizes
        // are read through COM because WPS's in-memory XML may keep them in the table STYLE instead
        // of spelling them out per table.
        bool already = ThickBorders(before) && (info == null || info.HasBorders);
        if (already)
        {
            b["applied"] = false;
            b["skipped"] = "already-thick";
            b["ok"] = ThickBorders(before);
            b["after"] = new List<object>(new object[] { before[0], before[1] });
            skipped++;
            outBorders["borders"] = b;
            return;
        }
        Tr("borders table");
        TabBorders(table, mode);
        int[] after = BorderPair(table);
        applied++;
        b["applied"] = true;
        b["after"] = new List<object>(new object[] { after[0], after[1] });
        b["ok"] = ThickBorders(after);
        if (!ThickBorders(after))
            warnings.Add("borders did not read back as thick (outer " + after[0] + ", inner " + after[1] + ")");
        outBorders["borders"] = b;
    }

    /// <summary>
    /// Put the grid columns back to what the snapshot says, in points, then let the caller re-read
    /// the XML to confirm. The table's preferred width is written back with the type it had.
    /// </summary>
    private static bool RestoreGridWidths(dynamic doc, dynamic table, TableXmlInfo before)
    {
        if (before == null || before.GridCols.Count == 0) return false;
        bool any = false;
        try
        {
            if (before.WidthType == "pct")
            {
                // OOXML pct is 50ths of a percent; COM wdPreferredWidthPercent wants 0..100.
                double pct = before.Width / 50.0;
                if (pct < 0.1) pct = 0.1;
                if (pct > 100) pct = 100;
                table.PreferredWidthType = PREF_PERCENT;
                table.PreferredWidth = (float)Math.Round(pct, 2);
                any = true;
            }
            else if (before.WidthType == "dxa") { table.PreferredWidthType = PREF_POINTS; table.PreferredWidth = (float)Math.Round(before.Width / 20.0, 1); any = true; }
            else { table.PreferredWidthType = PREF_POINTS; table.PreferredWidth = (float)Math.Round(GridSumPt(before), 1); any = true; }
        }
        catch (Exception e) { Tr("preferred width restore failed: " + Clean(e.Message)); }
        // WPS may report fewer columns than the XML snapshot (merged cells / re-derived grid), so
        // never ask for a column index the live table does not have.
        int live = Convert.ToInt32(TryGet(() => (object)(int)table.Columns.Count, (object)0), CultureInfo.InvariantCulture);
        int limit = before.GridCols.Count;
        if (live > 0 && live < limit) limit = live;
        for (int c = 1; c <= limit; c++)
        {
            double pt = Math.Round(before.GridCols[c - 1] / 20.0, 1);
            try
            {
                dynamic col = table.Columns.Item(c);
                col.PreferredWidthType = PREF_POINTS;
                col.PreferredWidth = (float)pt;
                col.Width = (float)pt;
                any = true;
            }
            catch (Exception e) { Tr("column " + c + " width restore failed: " + Clean(e.Message)); }
        }
        return any;
    }

    // ---------------------------------------------------------- optional renumber (approved plan only)

    private static Dictionary<string, object> LoadRenumber(string path)
    {
        if (!File.Exists(path)) throw new Fail("BAD_ARGS", "the plan file does not exist: " + path);
        object parsed = Json.Parse(File.ReadAllText(path, Encoding.UTF8));
        Dictionary<string, object> root = parsed as Dictionary<string, object>;
        if (root == null) throw new Fail("BAD_ARGS", "the plan file must hold a JSON object");
        Dictionary<string, object> rn = root.ContainsKey("renumber") ? root["renumber"] as Dictionary<string, object> : root;
        if (rn == null) return null;
        object approved = null;
        if (rn.ContainsKey("approved")) approved = rn["approved"];
        if (!(approved is bool) || !(bool)approved) return null;   // not approved = propose only, change nothing
        return rn;
    }

    /// <summary>Replace the number token of one caption paragraph, leaving label / gap / title alone.</summary>
    private static bool ReplaceCaptionNumber(dynamic doc, int para, string from, string to)
    {
        try
        {
            dynamic p = doc.Paragraphs.Item(para);
            int fields = Convert.ToInt32(TryGet(() => (object)(int)p.Range.Fields.Count, (object)0), CultureInfo.InvariantCulture);
            if (fields > 0) return false;   // a SEQ field owns the number: rewriting it would kill the field
            string text = ParaTextAt(doc, para);
            CaptionInfo ci = ParseCaption(text);
            if (ci == null || ci.Number != from) return false;
            int at = text.IndexOf(from, ci.Label.Length, StringComparison.Ordinal);
            if (at < 0) return false;
            int start = (int)p.Range.Start;
            doc.Range(start + at, start + at + from.Length).Text = to;
            return true;
        }
        catch (Exception e) { Tr("replace caption number failed: " + Clean(e.Message)); return false; }
    }

    /// <summary>
    /// DESIGN 3.3: apply an APPROVED renumbering plan. Caption items are re-located (paragraph
    /// numbers move when a caption is inserted or a row deleted), then the in-text references that
    /// the plan maps explicitly are rewritten from the end of the document backwards.
    /// </summary>
    private static void ApplyRenumber(dynamic doc, Dictionary<string, object> rn,
        ref int renumbered, ref int referencesUpdated, List<string> warnings)
    {
        List<object> items = rn.ContainsKey("items") ? rn["items"] as List<object> : null;
        Dictionary<string, string> map = new Dictionary<string, string>();
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        int paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        if (items != null)
        {
            foreach (object o in items)
            {
                Dictionary<string, object> it = o as Dictionary<string, object>;
                if (it == null) continue;
                string from = Convert.ToString(it["from"], CultureInfo.InvariantCulture);
                string to = Convert.ToString(it["to"], CultureInfo.InvariantCulture);
                if (from.Length == 0 || to.Length == 0) continue;
                if (from == to) { map[from] = to; continue; }
                int target = 0;
                int para = it.ContainsKey("para") ? Convert.ToInt32(it["para"], CultureInfo.InvariantCulture) : 0;
                if (para >= 1 && para <= paraTotal)
                {
                    CaptionInfo ci = ParseCaption(ParaTextAt(doc, para));
                    if (ci != null && ci.Number == from) target = para;
                }
                if (target == 0 && it.ContainsKey("table"))
                {
                    int tn = Convert.ToInt32(it["table"], CultureInfo.InvariantCulture);
                    if (tn >= 1 && tn <= tables)
                    {
                        int sp = TableStartPara(doc, doc.Tables.Item(tn), paraTotal);
                        CaptionRec rec = CaptionAbove(doc, sp, new HashSet<int>());
                        if (rec != null && rec.Info.Number == from) target = rec.Para;
                    }
                }
                if (target == 0)
                {
                    warnings.Add("renumber: caption " + from + " could not be re-located (paragraphs moved); left alone");
                    continue;
                }
                if (ReplaceCaptionNumber(doc, target, from, to)) { renumbered++; map[from] = to; }
                else warnings.Add("renumber: caption " + from + " has a field or unexpected text; left alone");
            }
        }
        List<object> refsPlan = rn.ContainsKey("refs") ? rn["refs"] as List<object> : null;
        if (refsPlan != null)
        {
            foreach (object o in refsPlan)
            {
                Dictionary<string, object> it = o as Dictionary<string, object>;
                if (it == null) continue;
                string from = Convert.ToString(it["number"], CultureInfo.InvariantCulture);
                string to = Convert.ToString(it["to"], CultureInfo.InvariantCulture);
                if (from.Length == 0 || to.Length == 0) continue;
                if (it.ContainsKey("ambiguous") && it["ambiguous"] is bool && (bool)it["ambiguous"])
                {
                    warnings.Add("renumber: reference " + from + " is ambiguous; not rewritten");
                    continue;
                }
                map[from] = to;
            }
        }
        if (map.Count > 0) referencesUpdated = RewriteInTextReferences(doc, map, warnings);
    }

    /// <summary>
    /// Rewrite "表N" mentions in body text according to map, from the END backwards so the character
    /// offsets stay valid. Caption lines themselves are left alone (that is the caption's own number).
    /// </summary>
    private static int RewriteInTextReferences(dynamic doc, Dictionary<string, string> map, List<string> warnings)
    {
        string all = WholeText(doc);
        if (all == null) { warnings.Add("references: the host refused to give the whole text; nothing rewritten"); return 0; }
        int baseOffset = Convert.ToInt32(TryGet(() => (object)(int)doc.Content.Start, (object)0), CultureInfo.InvariantCulture);
        List<int[]> hits = new List<int[]>();
        List<string> news = new List<string>();
        string[] lines = all.Split('\r');
        int offset = 0;
        foreach (string rawLine in lines)
        {
            string line = rawLine;
            if (line != null && line.Length > 0)
            {
                string trimmed = line.Trim();
                CaptionInfo self = ParseCaption(trimmed);
                int lead = 0;
                while (lead < line.Length && (line[lead] == ' ' || line[lead] == '\u3000' || line[lead] == '\t')) lead++;
                foreach (Match m in RefRx.Matches(line))
                {
                    string num = m.Groups["num"].Value;
                    string to;
                    if (!map.TryGetValue(num, out to)) continue;
                    if (self != null && m.Index == lead && m.Groups["label"].Value == self.Label && num == self.Number) continue;
                    hits.Add(new int[] { baseOffset + offset + m.Groups["num"].Index, num.Length });
                    news.Add(to);
                }
            }
            offset += (line == null ? 0 : line.Length) + 1;
        }
        int done = 0;
        for (int i = hits.Count - 1; i >= 0; i--)
        {
            try
            {
                doc.Range(hits[i][0], hits[i][0] + hits[i][1]).Text = news[i];
                done++;
            }
            catch (Exception e) { Tr("reference rewrite failed: " + Clean(e.Message)); }
        }
        return done;
    }

    // ---------------------------------------------------------- per-table rollback

    /// <summary>
    /// A failed table is rolled back on its own: the caption paragraph is removed and - only if the
    /// table really was damaged (a row deleted, the table count changed, paragraphs lost) - the
    /// ORIGINAL table XML is put back where the table stood. Other tables are untouched (DESIGN 2.2).
    /// Returns how the rollback was done ("none" = nothing to undo, "xml", "undo").
    /// </summary>
    private static string RollbackTable(dynamic doc, TableWork w)
    {
        if (w.InsertedPara > 0) RemoveParagraph(doc, w.InsertedPara);
        // Has anything actually been damaged? A failed caption move never deletes a row, so the
        // destructive delete+InsertXML restore must NOT run for it (measured 2026-11-03: re-inserting
        // an 18-column landscape table this way added 6 pages to the document).
        bool damaged = false;
        try
        {
            if ((int)doc.Tables.Count != w.TablesBefore) damaged = true;
            else if (w.RowsBefore > 0 && (int)doc.Tables.Item(w.Index).Rows.Count != w.RowsBefore) damaged = true;
            else if (w.ParasBefore > 0 && (int)doc.Paragraphs.Count != w.ParasBefore) damaged = true;
        }
        catch (Exception) { damaged = true; }
        if (!damaged) return "none";
        if (w.SavedXml.Length > 0)
        {
            try
            {
                int count = (int)doc.Tables.Count;
                if (w.Index <= count) { try { doc.Tables.Item(w.Index).Delete(); } catch (Exception) { } }
                int at = Math.Max(0, w.AnchorEnd);
                doc.Range(at, at).InsertXML(FlatTablePackage(w.SavedXml));
                int after = (int)doc.Tables.Count;
                if (after == count) return "xml";
            }
            catch (Exception e) { Tr("xml rollback failed: " + Clean(e.Message)); }
        }
        try { if ((bool)doc.Undo()) return "undo"; }
        catch (Exception) { }
        return "none";
    }

    // ---------------------------------------------------------- the command itself

    private static object TabNormalize(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (total == 0) throw new Fail("BAD_ARGS", "the document has no table; nothing to normalize.");
        int only = a.Int("table", 0);
        if (only < 0) throw new Fail("BAD_ARGS", "\"table\" must be 1-based");
        if (only > total) throw new Fail("BAD_ARGS", "table " + only + " does not exist (the document has " + total + ").");
        bool dryRun = a.Flag("dryRun", false);
        if (!dryRun && IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; nothing was changed (-DryRun still works).");

        bool doCaptions = !a.Flag("noCaptions", false);
        bool doBorders = !a.Flag("noBorders", false);
        bool doRepeat = !a.Flag("noHeaderRepeat", false);
        int gap = a.Int("gap", 2);
        if (gap < 0) gap = 0;
        if (gap > 6) gap = 6;
        string borderMode = TabBorderMode(a.Str("borders", "thick"));
        if (borderMode.Length == 0) borderMode = "thick";
        string fit = a.Str("fit", "keep").Trim().ToLowerInvariant();
        if (fit != "keep" && fit != "window" && fit != "content")
            throw new Fail("BAD_ARGS", "unknown fit \"" + a.Str("fit", "") + "\" (keep | window | content)");
        int failTable = a.Int("failTable", 0);   // test hook for DESIGN A8 (isolate one broken table)
        int from = only > 0 ? only : 1;
        int to = only > 0 ? only : total;

        dynamic app = App(true, false);
        // The cut-paste caption method is only safe on real Microsoft Word (post-mortem 2026-11-03:
        // WPS Writer spins forever when a table is deleted and re-inserted through InsertXML), so it
        // needs an explicit -CutPaste on WPS.
        bool allowCutPaste = a.Flag("cutPaste", false) || !IsWpsApp(app);
        int selStart = -1, selEnd = -1;
        try { selStart = (int)app.Selection.Start; selEnd = (int)app.Selection.End; } catch (Exception) { }
        bool screenWas = true, screenSet = false;
        try { screenWas = (bool)app.ScreenUpdating; app.ScreenUpdating = false; screenSet = true; } catch (Exception) { }
        try
        {
            try { doc.Repaginate(); } catch (Exception) { }
            int paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            int pagesBefore = Convert.ToInt32(TryGet(() => (object)(int)doc.ComputeStatistics(2), (object)0), CultureInfo.InvariantCulture);
            int revBefore = Convert.ToInt32(TryGet(() => (object)(int)doc.Revisions.Count, (object)0), CultureInfo.InvariantCulture);
            int comBefore = Convert.ToInt32(TryGet(() => (object)(int)doc.Comments.Count, (object)0), CultureInfo.InvariantCulture);
            double textWidth = TextWidth(doc);
            string xmlBefore = WholeOpenXml(doc);
            List<TableXmlInfo> infosBefore = xmlBefore == null ? new List<TableXmlInfo>() : ParseTableXml(xmlBefore, true);

            if (dryRun)
                return NormalizeDryRun(doc, app, total, from, to, infosBefore, textWidth, doCaptions, doBorders, doRepeat, gap, borderMode, fit);

            Dictionary<string, object> renumberPlan = null;
            string planPath = a.Str("plan", "");
            if (planPath.Length > 0) renumberPlan = LoadRenumber(planPath);

            int moved = 0, moveFailed = 0, spaces = 0, bordersApplied = 0, bordersSkipped = 0;
            int widthsRestored = 0, renumbered = 0, refsUpdated = 0, headerChanged = 0;
            List<object> tableResults = new List<object>();
            List<object> failures = new List<object>();
            List<object> needsHuman = new List<object>();
            List<string> globalWarnings = new List<string>();
            Dictionary<int, Dictionary<string, object>> entries = new Dictionary<int, Dictionary<string, object>>();
            Dictionary<int, bool> restoredAt = new Dictionary<int, bool>();
            bool abortBatch = false;

            using (P2Edit edit = new P2Edit(doc, false, "wps-word-cooperate table normalize"))
            {
                for (int n = from; n <= to; n++)
                {
                    Tr("normalize table " + n);
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["table"] = n;
                    List<string> warnings = new List<string>();
                    TableWork work = new TableWork();
                    work.Index = n;
                    try
                    {
                        dynamic table = doc.Tables.Item(n);
                        int startPara = TableStartPara(doc, table, paraTotal);
                        work.TablesBefore = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
                        work.RowsBefore = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
                        work.ParasBefore = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
                        if (startPara > 1)
                            work.AnchorEnd = Convert.ToInt32(TryGet(() => (object)(int)doc.Paragraphs.Item(startPara - 1).Range.End, (object)0), CultureInfo.InvariantCulture);
                        TableXmlInfo infoBefore = (n - 1) < infosBefore.Count ? infosBefore[n - 1] : null;
                        if (infoBefore != null) work.SavedXml = infoBefore.Raw;

                        // ---- 1) the caption out of row 1
                        if (doCaptions)
                        {
                            string firstCell = CleanPara(Convert.ToString(TryGet(() => (object)doc.Tables.Item(n).Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                            if (IsCaptionText(firstCell))
                            {
                                CaptionMoveResult mv = MoveCaptionOutOfTable(doc, app, n, allowCutPaste);
                                if (mv.Ok)
                                {
                                    moved++;
                                    work.InsertedPara = mv.NewPara;
                                    Dictionary<string, object> cap = new Dictionary<string, object>();
                                    cap["before"] = Head(mv.TextBefore, 80);
                                    cap["after"] = Head(mv.TextAfter, 80);
                                    cap["para"] = mv.NewPara;
                                    cap["page"] = mv.Page;
                                    cap["method"] = mv.Method;
                                    one["caption"] = cap;
                                }
                                else if (mv.Incomplete)
                                {
                                    moveFailed++;
                                    throw new Fail("ROLLBACK_INCOMPLETE", mv.Error);
                                }
                                else if (mv.NeedsHuman)
                                {
                                    // Not a bug: WPS cannot put a body paragraph in front of a table
                                    // that begins its section. Report it, leave the table untouched
                                    // and keep going with the rest of the table's work.
                                    moveFailed++;
                                    Dictionary<string, object> nh = new Dictionary<string, object>();
                                    nh["captionMoved"] = "needsHuman:wps-section-start";
                                    nh["detail"] = "the section-break mark sits in the paragraph right in "
                                        + "front of this table, so WPS puts every inserted paragraph into "
                                        + "the PREVIOUS section (caption page < table page). Move the "
                                        + "section break behind the caption by hand, or run this table in "
                                        + "Microsoft Word with -CutPaste.";
                                    one["needsHuman"] = nh;
                                    Dictionary<string, object> nhTop = new Dictionary<string, object>();
                                    nhTop["table"] = n;
                                    nhTop["reason"] = "wps-section-start";
                                    nhTop["hint"] = "WPS cannot create a body paragraph in front of a table that "
                                        + "begins its section. In WPS, cut the section break behind the caption by "
                                        + "hand; in Microsoft Word, run this table with -CutPaste.";
                                    needsHuman.Add(nhTop);
                                    warnings.Add("caption left inside the table: WPS section-start limitation");
                                }
                                else
                                {
                                    moveFailed++;
                                    Dictionary<string, object> fail = new Dictionary<string, object>();
                                    fail["table"] = n;
                                    fail["code"] = "CAPTION_MOVE_FAILED";
                                    fail["detail"] = Head(mv.Error, 200);
                                    fail["rollback"] = "none";
                                    failures.Add(fail);
                                    warnings.Add("caption move failed: " + Head(mv.Error, 120));
                                }
                            }
                            else one["captionMoved"] = false;
                        }

                        // test hook: break ONE table on purpose to prove the isolation (DESIGN A8)
                        if (failTable == n) throw new Fail("INJECTED_FAILURE", "test hook -FailTable " + n + " was requested");

                        // ---- 2) caption spacing (whole paragraph text, one write per caption)
                        if (gap > 0)
                        {
                            dynamic tableNow = doc.Tables.Item(n);
                            int startParaNow = TableStartPara(doc, tableNow, (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0));
                            CaptionRec above = CaptionAbove(doc, startParaNow, new HashSet<int>());
                            int capPara = above == null ? 0 : above.Para;
                            if (capPara > 0)
                            {
                                string sb, sa, skip;
                                bool changed = NormalizeCaptionGap(doc, capPara, gap, out sb, out sa, out skip);
                                if (changed) spaces++;
                                if (skip != null && skip != "not a caption")
                                {
                                    warnings.Add("caption " + capPara + ": spacing left alone (" + skip + ")");
                                }
                            }
                        }

                        // ---- 3) thick borders (skipped when the table already carries them)
                        if (doBorders)
                        {
                            Dictionary<string, object> borderOut = new Dictionary<string, object>();
                            ApplyBordersKeepLayout(doc, doc.Tables.Item(n), borderMode, infoBefore, borderOut, warnings,
                                ref bordersApplied, ref bordersSkipped);
                            one["borders"] = borderOut["borders"];
                        }
                    }
                    catch (Exception e)
                    {
                        Fail f = e as Fail;
                        string code = f != null ? f.Code : "COM_ERROR";
                        string detail = Clean(e.Message);
                        object extra = f != null ? f.Detail : null;
                        string rollback = RollbackTable(doc, work);
                        Dictionary<string, object> fail = new Dictionary<string, object>();
                        fail["table"] = n;
                        fail["code"] = code;
                        fail["detail"] = Head(detail, 200);
                        fail["rollback"] = rollback;
                        if (extra != null) fail["extra"] = extra;
                        failures.Add(fail);
                        warnings.Add(code + ": " + Head(detail, 120));
                        one["failed"] = true;
                        // An incomplete rollback leaves the document dirty: stop the whole batch
                        // rather than pile more edits on top of it (lead's rule 2026-11-03).
                        if (code == "ROLLBACK_INCOMPLETE") abortBatch = true;
                    }
                    one["warnings"] = new List<object>(warnings.ToArray());
                    tableResults.Add(one);
                    entries[n] = one;
                    if (abortBatch) break;
                }

                // ---- 4) grid columns: compare with the snapshot and put drifted tables back
                string xmlAfter = WholeOpenXml(doc);
                List<TableXmlInfo> infosAfter = xmlAfter == null ? new List<TableXmlInfo>() : ParseTableXml(xmlAfter, false);
                Dictionary<int, double> drift = new Dictionary<int, double>();
                for (int n = from; n <= to; n++)
                {
                    double d = (n - 1) < infosAfter.Count && (n - 1) < infosBefore.Count
                        ? GridDrift(infosBefore[n - 1], infosAfter[n - 1]) : 0;
                    drift[n] = d;
                    if (d <= 0.02) continue;
                    Tr("table " + n + " grid drifted by " + Convert.ToString(Math.Round(d * 100, 1), CultureInfo.InvariantCulture) + "%");
                    bool fixedOk = false;
                    try { fixedOk = RestoreGridWidths(doc, doc.Tables.Item(n), infosBefore[n - 1]); }
                    catch (Exception e) { Tr("grid restore failed for " + n + ": " + Clean(e.Message)); }
                    if (fixedOk) { widthsRestored++; restoredAt[n] = true; }
                }

                // ---- 5) repeat header rows on every page-straddling table
                if (doRepeat)
                {
                    Dictionary<string, object> rep = RepeatTableHeaders(doc, app, from, to, 3, 0, only > 0, total);
                    headerChanged = Convert.ToInt32(rep["changed"], CultureInfo.InvariantCulture);
                }

                // ---- 6) optional renumbering (approved plan file only)
                if (renumberPlan != null) ApplyRenumber(doc, renumberPlan, ref renumbered, ref refsUpdated, globalWarnings);
            }

            // ---- verification sweep (fresh page numbers, fresh XML)
            try { doc.Repaginate(); } catch (Exception) { }
            string xmlFinal = WholeOpenXml(doc);
            List<TableXmlInfo> infosFinal = xmlFinal == null ? new List<TableXmlInfo>() : ParseTableXml(xmlFinal, false);
            int widthsChanged = 0;
            for (int n = from; n <= to; n++)
            {
                double d = (n - 1) < infosFinal.Count && (n - 1) < infosBefore.Count
                    ? GridDrift(infosBefore[n - 1], infosFinal[n - 1]) : 0;
                bool changed = d > 0.02;
                if (changed) widthsChanged++;
                bool restored = false;
                restoredAt.TryGetValue(n, out restored);
                Dictionary<string, object> w = new Dictionary<string, object>();
                w["changed"] = changed;
                w["restored"] = restored;
                if ((changed || restored) && (n - 1) < infosFinal.Count && (n - 1) < infosBefore.Count)
                {
                    w["before"] = GridColsPt(infosBefore[n - 1]);
                    w["after"] = GridColsPt(infosFinal[n - 1]);
                    if (d > 0.02)
                        globalWarnings.Add("table " + n + ": column widths still differ by "
                            + Convert.ToString(Math.Round(d * 100, 1), CultureInfo.InvariantCulture) + "% after the restore");
                }
                Dictionary<string, object> entry;
                if (entries.TryGetValue(n, out entry)) entry["widths"] = w;
            }

            int pagesAfter = Convert.ToInt32(TryGet(() => (object)(int)doc.ComputeStatistics(2), (object)0), CultureInfo.InvariantCulture);
            int paragraphsAfter = Convert.ToInt32(TryGet(() => (object)(int)doc.Paragraphs.Count, (object)0), CultureInfo.InvariantCulture);
            int revAfter = Convert.ToInt32(TryGet(() => (object)(int)doc.Revisions.Count, (object)0), CultureInfo.InvariantCulture);
            int comAfter = Convert.ToInt32(TryGet(() => (object)(int)doc.Comments.Count, (object)0), CultureInfo.InvariantCulture);
            if (revAfter != revBefore) globalWarnings.Add("revision count changed (" + revBefore + " -> " + revAfter + ")");
            if (comAfter != comBefore) globalWarnings.Add("comment count changed (" + comBefore + " -> " + comAfter + ")");

            // per-table page + header-repeat readback
            int capInRow = 0, notThick = 0, multipageMissing = 0, capPageMismatch = 0;
            HashSet<int> captionParas = new HashSet<int>();
            Dictionary<int, int> capPageOf = new Dictionary<int, int>();
            Dictionary<int, int> tablePageOf = new Dictionary<int, int>();
            int paraTotalNow = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            try { VerifyTableSweep(doc, app, from, to, paraTotalNow, captionParas, capPageOf, tablePageOf, true); }
            catch (Exception e) { globalWarnings.Add("verification sweep failed: " + Clean(e.Message)); }
            for (int n = from; n <= to; n++)
            {
                Tr("verify table " + n);
                dynamic table = null;
                try { table = doc.Tables.Item(n); } catch (Exception) { continue; }
                string firstCell = CleanPara(Convert.ToString(TryGet(() => (object)table.Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                if (IsCaptionText(firstCell)) capInRow++;
                int[] borders = BorderPair(table);
                if (!ThickBorders(borders)) notThick++;
                int[] span = PageSpanOf(doc, table);
                bool split = span[1] > span[0] && span[0] > 0;
                int heading = ReadHeadingFormat(app, table);
                if (split && heading != -1) multipageMissing++;
                int capPage = 0, tablePage = 0;
                capPageOf.TryGetValue(n, out capPage);
                tablePageOf.TryGetValue(n, out tablePage);
                if (capPage > 0 && tablePage > 0 && capPage != tablePage) capPageMismatch++;
                Dictionary<string, object> entry;
                if (entries.TryGetValue(n, out entry))
                {
                    entry["page"] = span[0];
                    // report the caption page from THIS sweep too, so page and caption.page can
                    // never disagree because they were read at different moments
                    object capRaw = null;
                    if (entry.TryGetValue("caption", out capRaw) && capRaw is Dictionary<string, object> && capPage > 0)
                        ((Dictionary<string, object>)capRaw)["page"] = capPage;
                    Dictionary<string, object> hr = new Dictionary<string, object>();
                    hr["applied"] = heading == -1;
                    entry["headerRepeat"] = hr;
                }
            }

            Dictionary<string, object> summary = new Dictionary<string, object>();
            summary["captionsMoved"] = moved;
            summary["captionsMovedFailed"] = moveFailed;
            summary["spacesNormalized"] = spaces;
            summary["renumbered"] = renumbered;
            summary["referencesUpdated"] = refsUpdated;
            summary["bordersApplied"] = bordersApplied;
            summary["bordersSkipped"] = bordersSkipped;
            summary["widthsRestored"] = widthsRestored;
            summary["headerRepeated"] = headerChanged;
            summary["paragraphsBefore"] = paraTotal;
            summary["paragraphsAfter"] = paragraphsAfter;
            // a moved caption is "-1 title row + 1 caption paragraph", so this must be 0 (the
            // 2026-11-03 trace once showed 4063 != 4069: that was the rollback leak)
            summary["paragraphDelta"] = paragraphsAfter - paraTotal;
            summary["pageBefore"] = pagesBefore;
            summary["pageAfter"] = pagesAfter;
            summary["revisionsBefore"] = revBefore;
            summary["revisionsAfter"] = revAfter;
            summary["commentsBefore"] = comBefore;
            summary["commentsAfter"] = comAfter;

            Dictionary<string, object> verification = new Dictionary<string, object>();
            verification["captionInFirstRow"] = capInRow;
            verification["bordersNotThick"] = notThick;
            verification["multipageWithoutHeaderRepeat"] = multipageMissing;
            verification["captionPageMismatch"] = capPageMismatch;
            verification["widthsChanged"] = widthsChanged;
            verification["pageDelta"] = pagesAfter - pagesBefore;

            Dictionary<string, object> r = new Dictionary<string, object>();
            r["action"] = "normalize";
            r["tableCount"] = to - from + 1;
            r["dryRun"] = false;
            r["summary"] = summary;
            r["verification"] = verification;
            r["tables"] = tableResults;
            r["failures"] = failures;
            if (needsHuman.Count > 0) r["needsHuman"] = needsHuman;
            if (globalWarnings.Count > 0) r["warnings"] = new List<object>(globalWarnings.ToArray());
            r["saved"] = false;
            r["note"] = "format-only, one undo step (Ctrl+Z takes the whole command back); nothing was "
                + "saved and no revision marks were produced. Caption moves are proven by reading the "
                + "page number back; column widths come from a WordOpenXML snapshot because WPS reports "
                + "Column.Width as wdUndefined after it rebuilds a table layout. Render the affected "
                + "pages before saving.";
            return r;
        }
        finally
        {
            if (screenSet) { try { app.ScreenUpdating = screenWas; } catch (Exception) { } }
            RestoreSelection(doc, selStart, selEnd);
        }
    }

    /// <summary>
    /// One read-only sweep that pairs every table with "the caption above it" and records both page
    /// numbers. captionParas is filled in the first pass so a caption is never attributed twice.
    /// </summary>
    private static void VerifyTableSweep(dynamic doc, dynamic app, int from, int to, int paraTotal,
        HashSet<int> captionParas, Dictionary<int, int> capPageOf, Dictionary<int, int> tablePageOf, bool fillPages)
    {
        for (int n = from; n <= to; n++)
        {
            int startPara = 0;
            try { startPara = TableStartPara(doc, doc.Tables.Item(n), paraTotal); }
            catch (Exception) { continue; }
            CaptionRec above = CaptionAbove(doc, startPara, captionParas);
            if (above != null && above.Para > 0) captionParas.Add(above.Para);
        }
        if (!fillPages) return;
        for (int n = from; n <= to; n++)
        {
            try
            {
                dynamic table = doc.Tables.Item(n);
                int[] span = PageSpanOf(doc, table);
                tablePageOf[n] = span[0];
                int startPara = TableStartPara(doc, table, paraTotal);
                CaptionRec above = CaptionAbove(doc, startPara, new HashSet<int>());
                if (above != null) capPageOf[n] = PageOfPara(doc, above.Para);
            }
            catch (Exception) { }
        }
    }

    // ---------------------------------------------------------- dry run

    /// <summary>
    /// -DryRun: the same decisions the real run would take, without writing anything (works on a
    /// read-only document). The insertion method is a STRUCTURAL prediction, not a page check.
    /// </summary>
    private static object NormalizeDryRun(dynamic doc, dynamic app, int total, int from, int to,
        List<TableXmlInfo> infos, double textWidth, bool doCaptions, bool doBorders, bool doRepeat, int gap,
        string borderMode, string fit)
    {
        int paraTotal = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        List<object> tables = new List<object>();
        int wouldMove = 0, wouldBorders = 0, wouldSkip = 0, wouldSpaces = 0, wouldNeedsHuman = 0;
        for (int n = from; n <= to; n++)
        {
            dynamic table = null;
            try { table = doc.Tables.Item(n); }
            catch (Exception) { continue; }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["table"] = n;
            int[] span = PageSpanOf(doc, table);
            one["page"] = span[0];
            string firstCell = CleanPara(Convert.ToString(TryGet(() => (object)table.Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
            bool capInRow = doCaptions && IsCaptionText(firstCell);
            if (capInRow)
            {
                int startPara = TableStartPara(doc, table, paraTotal);
                bool atSectionStart, sectPrBefore, pageBreakChar, pageBreakBefore;
                DetectBreakBefore(doc, table, startPara, out atSectionStart, out sectPrBefore, out pageBreakChar, out pageBreakBefore);
                // the real run always tries "table-para" first (Table.Range.InsertParagraphBefore):
                // that is the spelling that keeps a caption with its table across a page/section break
                string method = "table-para";
                Dictionary<string, object> cap = new Dictionary<string, object>();
                bool sectionBreak = sectPrBefore || atSectionStart;
                cap["wouldMove"] = !sectionBreak;
                if (sectionBreak) cap["needsHuman"] = "wps-section-start";
                else cap["method"] = method;
                cap["breakInFront"] = sectionBreak ? "section"
                    : ((pageBreakChar || pageBreakBefore) ? "page" : "none");
                cap["text"] = Head(firstCell, 60);
                one["caption"] = cap;
                if (sectionBreak) wouldNeedsHuman++; else wouldMove++;
            }
            if (doBorders)
            {
                TableXmlInfo info = (n - 1) < infos.Count ? infos[n - 1] : null;
                int[] borders = BorderPair(table);
                Dictionary<string, object> b = new Dictionary<string, object>();
                b["before"] = new List<object>(new object[] { borders[0], borders[1] });
                bool already = ThickBorders(borders) && (info == null || info.HasBorders);
                b["wouldApply"] = !already;
                if (already) { b["skipped"] = "already-thick"; wouldSkip++; }
                else wouldBorders++;
                one["borders"] = b;
            }
            if (gap > 0 && doCaptions)
            {
                int startPara = TableStartPara(doc, table, paraTotal);
                CaptionRec above = CaptionAbove(doc, startPara, new HashSet<int>());
                if (above != null)
                {
                    string want = above.Info.Label + " " + above.Info.Number
                        + (above.Info.Title.Trim().Length > 0 ? new string(' ', gap) + above.Info.Title.Trim() : "");
                    if (want.Trim() != ParaTextAt(doc, above.Para).Trim()) wouldSpaces++;
                }
            }
            tables.Add(one);
        }
        Dictionary<string, object> would = new Dictionary<string, object>();
        would["moveCaptions"] = wouldMove;
        would["needsHuman"] = wouldNeedsHuman;
        would["normalizeSpaces"] = wouldSpaces;
        would["applyBorders"] = wouldBorders;
        would["skipBorders"] = wouldSkip;
        would["repeatHeaders"] = doRepeat;
        would["borders"] = borderMode;
        would["fit"] = fit;
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "normalize";
        r["dryRun"] = true;
        r["tableCount"] = to - from + 1;
        r["would"] = would;
        r["tables"] = tables;
        r["saved"] = false;
        r["note"] = "dry run: nothing was written (no undo record either). caption.method is the method "
            + "the real run will TRY FIRST (chosen from the break in front of the table); the real run "
            + "falls through to the next method when the page number does not read back.";
        return r;
    }
}

