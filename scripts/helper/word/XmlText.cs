// v1.13 - paragraph texts from ONE WordOpenXML call.
//
// Field evidence (2026-11-01, a real 61-page / 7334-paragraph / 23-table report):
//   * doc.Content.Text produced 7328 blocks, misaligned from paragraph 2 on -> the v1.7-v1.10
//     block model could not be used at all, and snapshot/diff/read silently fell back to
//     ~5 COM roundtrips per paragraph (snapshot 45 s, diff 52 s, read 30 s per call).
//   * Paragraphs.Count = 7334 while the story XML holds 6779 w:p = exactly the number of table
//     ROWS fewer (sum of Tables(i).Rows.Count = 555). WPS counts every table row as one extra,
//     EMPTY paragraph at the end of that row - measured on table 1: cells p18..p21 then an empty
//     p22, cells p23..p26 then an empty p27, ... 8 rows -> 8 empty entries.
//   * doc.Content.WordOpenXML returns the whole story in ~0.4-1.7 s (8.2 MB here).
//
// So: one roundtrip, parse locally, re-create WPS's paragraph sequence by emitting one empty
// paragraph after every table row. The result is verified against Paragraphs.Count and by sampling
// real paragraphs; anything unexpected falls back to the per-paragraph loop.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

internal static partial class Program
{
    private const string W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>0-based indices of paragraphs the XML cannot render faithfully (re-read from COM).</summary>
    private static HashSet<int> _xmlRisky;

    /// <summary>Diagnostic summary of the last XmlDocTexts run ("risky=3, ms=1200").</summary>
    private static string _xmlNote;

    /// <summary>
    /// Paragraph texts (WPS paragraph numbering, i.e. one phantom empty paragraph per table row)
    /// from a single WordOpenXML fetch, or null with <paramref name="why"/> set.
    /// Semantics match Paragraphs.Item(k).Range.Text: unaccepted deletions are still present
    /// (w:delText is included).
    /// </summary>
    private static List<string> XmlDocTexts(dynamic doc, int total, out string why)
    {
        why = null;
        // v1.34: long-lived daemon - clear the diagnostics of the previous run up front, otherwise a
        // failed run would report the previous request's "risky=" note.
        _xmlNote = null;
        _xmlRisky = null;
        if (total <= 0) { why = "no paragraphs"; return null; }

        string xml;
        try { xml = Convert.ToString(doc.Content.WordOpenXML, CultureInfo.InvariantCulture); }
        catch (Exception e) { why = "WordOpenXML refused: " + Clean(e.Message); return null; }
        if (xml == null || xml.Length == 0) { why = "WordOpenXML empty"; return null; }

        List<string> texts;
        try
        {
            XmlDocument xd = new XmlDocument();
            xd.XmlResolver = null;
            xd.PreserveWhitespace = true;
            xd.LoadXml(xml);
            // WordOpenXML is flat OPC: <pkg:package><pkg:part name="/word/document.xml">
            // <pkg:xmlData><w:document><w:body>. Find the body wherever it sits.
            XmlNode body = null;
            XmlNodeList bodies = xd.GetElementsByTagName("body", W_NS);
            if (bodies != null && bodies.Count > 0) body = bodies[0];
            if (body == null) { why = "no w:body in WordOpenXML"; return null; }
            texts = new List<string>(total);
            _xmlRisky = new HashSet<int>();
            WalkBlocks(body, texts);
        }
        catch (Exception e) { why = "XML parse failed: " + Clean(e.Message); return null; }

        if (texts.Count != total)
        {
            why = "xml paragraphs=" + texts.Count + " total=" + total;
            return null;
        }

        // Paragraphs whose XML we cannot map to Range.Text with certainty (Wingdings/w:sym symbols
        // render as other characters, several inline objects collapse into one). A handful of these
        // is normal; re-read just those from COM, which keeps the fast path exact.
        _xmlNote = "risky=" + _xmlRisky.Count;
        if (_xmlRisky.Count > 0)
        {
            if (_xmlRisky.Count > 400) { why = "too many symbol paragraphs (" + _xmlRisky.Count + ")"; return null; }
            foreach (int idx in _xmlRisky)
            {
                // _xmlRisky holds 0-based positions in texts; WPS paragraphs are 1-based.
                int k = idx + 1;
                if (k < 1 || k > total) continue;
                try { texts[idx] = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); }
                catch (Exception) { why = "symbol re-read p" + k + " refused"; return null; }
            }
        }

        // Verify a spread of real paragraphs against COM. Cheap (a dozen roundtrips) and it is the
        // only thing standing between "fast" and "silently wrong".
        foreach (int k in SpotParagraphs(total))
        {
            string para;
            try { para = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); }
            catch (Exception) { why = "spot p" + k + " refused"; return null; }
            string xmlText = CleanPara(texts[k - 1]);
            if (para != xmlText)
            {
                why = "spot p" + k + " differs (xml " + Describe(xmlText) + " vs para " + Describe(para) + ")";
                return null;
            }
        }
        return texts;
    }

    /// <summary>
    /// Walk a block-level container (w:body / w:tc / w:txbxContent) in document order, appending
    /// one text per paragraph and one phantom empty paragraph per table row.
    /// </summary>
    private static void WalkBlocks(XmlNode container, List<string> texts)
    {
        foreach (XmlNode n in container.ChildNodes)
        {
            if (n.NodeType != XmlNodeType.Element) continue;
            if (n.NamespaceURI != W_NS) continue;
            switch (n.LocalName)
            {
                case "p":
                    {
                        int idx = texts.Count;
                        bool risky;
                        string pt = ParaTextOf(n, out risky);
                        texts.Add(pt);
                        if (risky) _xmlRisky.Add(idx);
                        break;
                    }
                case "tbl":
                    WalkTable(n, texts);
                    break;
                case "sdt":
                case "sdtContent":
                case "customXml":
                    WalkBlocks(n, texts);
                    break;
                default:
                    break;
            }
        }
    }

    private static void WalkTable(XmlNode tbl, List<string> texts)
    {
        foreach (XmlNode tr in tbl.ChildNodes)
        {
            if (tr.NodeType != XmlNodeType.Element || tr.NamespaceURI != W_NS || tr.LocalName != "tr") continue;
            foreach (XmlNode tc in tr.ChildNodes)
            {
                if (tc.NodeType != XmlNodeType.Element || tc.NamespaceURI != W_NS) continue;
                if (tc.LocalName == "tc") WalkBlocks(tc, texts);
                else if (tc.LocalName == "sdt" || tc.LocalName == "sdtContent" || tc.LocalName == "customXml") WalkBlocks(tc, texts);
            }
            // WPS quirk: every table row adds one EMPTY paragraph to the document paragraph count.
            texts.Add("");
        }
    }

    /// <summary>Text of one w:p as Range.Text would show it (accept-structure view: deleted runs kept).</summary>
    private static string ParaTextOf(XmlNode p, out bool risky)
    {
        risky = false;
        StringBuilder sb = new StringBuilder();
        AppendRunsWithRisk(p, sb, ref risky);
        // WPS shows a section break as a form feed at the end of its paragraph (Range.Text proof:
        // two paragraphs of the 2026-11-01 report are exactly "\x0c").
        foreach (XmlNode c in p.ChildNodes)
            if (c.NodeType == XmlNodeType.Element && c.NamespaceURI == W_NS && c.LocalName == "pPr")
                foreach (XmlNode g in c.ChildNodes)
                    if (g.NodeType == XmlNodeType.Element && g.NamespaceURI == W_NS && g.LocalName == "sectPr")
                    {
                        sb.Append('\f');
                        return sb.ToString();
                    }
        return sb.ToString();
    }

    private static void AppendRunsWithRisk(XmlNode node, StringBuilder sb, ref bool risky)
    {
        int objects = CountObjects(node);
        if (objects >= 2) risky = true;
        XmlNodeList all = node.SelectNodes(".//*");
        if (all != null)
            foreach (XmlNode c in all)
                if (c.NamespaceURI == W_NS && c.LocalName == "sym") { risky = true; break; }
        AppendRuns(node, sb);
    }

    private static int CountObjects(XmlNode node)
    {
        int n = 0;
        XmlNodeList all = node.SelectNodes(".//*");
        if (all == null) return 0;
        foreach (XmlNode c in all)
        {
            if (c.NamespaceURI != W_NS) continue;
            if (c.LocalName == "drawing" || c.LocalName == "pict" || c.LocalName == "object") n++;
        }
        return n;
    }

    private static void AppendRuns(XmlNode node, StringBuilder sb)
    {
        foreach (XmlNode n in node.ChildNodes)
        {
            if (n.NodeType == XmlNodeType.Text || n.NodeType == XmlNodeType.CDATA)
            {
                continue;   // whitespace between elements; real text lives in w:t
            }
            if (n.NodeType != XmlNodeType.Element) continue;
            if (n.NamespaceURI != W_NS) { continue; }
            switch (n.LocalName)
            {
                case "t":
                    sb.Append(n.InnerText);
                    break;
                case "delText":
                    break;                       // Range.Text hides unaccepted deletions
                case "tab":
                    sb.Append('\t');
                    break;
                case "br":
                case "cr":
                    {
                        string type = Attr(n, "type");
                        if (type == "page") sb.Append('\f');
                        else sb.Append('\v');
                        break;
                    }
                case "noBreakHyphen":
                    sb.Append('-');
                    break;
                case "softHyphen":
                    break;
                case "sym":
                    {
                        string ch = Attr(n, "char");
                        int code;
                        if (ch != null && int.TryParse(ch, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                            sb.Append((char)code);
                        break;
                    }
                case "instrText":
                case "del":
                case "ins":
                case "smartTag":
                case "hyperlink":
                case "fldSimple":
                case "sdt":
                case "sdtContent":
                case "customXml":
                case "rPr":
                case "pPr":
                case "bookmarkStart":
                case "bookmarkEnd":
                case "proofErr":
                    {
                        if (n.LocalName == "instrText") break;          // field CODE, not the result
                        if (n.LocalName == "rPr" || n.LocalName == "pPr") break;
                        if (n.LocalName == "bookmarkStart" || n.LocalName == "bookmarkEnd" || n.LocalName == "proofErr") break;
                        AppendRuns(n, sb);                              // wrappers: recurse
                        break;
                    }
                case "drawing":
                case "pict":
                case "object":
                    sb.Append('/');                                    // WPS shows inline objects as '/' in Range.Text
                    break;
                default:
                    AppendRuns(n, sb);
                    break;
            }
        }
    }

    private static string Attr(XmlNode n, string local)
    {
        XmlAttributeCollection at = n.Attributes;
        if (at == null) return null;
        foreach (XmlAttribute a in at)
            if (a.NamespaceURI == W_NS && a.LocalName == local) return a.Value;
        return null;
    }
}
