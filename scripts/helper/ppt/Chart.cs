// wps-office-cooperate (ppt): `chart` - native, editable charts from a data table.
//
// This is the one thing a picture-of-a-chart cannot do: the numbers stay in the deck,
// the reader can switch series or restyle it, and the data survives a save.
//
// Measured on WPS 12.1.0.29161, because every one of these cost a probe:
//   * Shapes.AddChart / AddChart2 work ONLY when the presentation has a window.  A deck
//     opened by the helper without -Show refuses with E_FAIL, and setting
//     Application.Visible afterwards does not rescue it.  So the error message tells the
//     operator to reopen with -Show instead of pretending.
//   * Chart.SeriesCollection(1).Values = array succeeds but is NOT persisted: after a
//     save the chart falls back to WPS's seed data (4.3 2.5 3.5 4.5).  The only route
//     that persists is the chart's embedded workbook - ChartData.Activate() then
//     ChartData.Workbook.ActiveSheet.Cells(r,c).Value - which also matches the note that
//     WPS takes literal values only (no formulas).
//   * The embedded workbook needs a retry loop (it appears a beat after Activate), and
//     WPS seeds it with 3 series x 4 categories: extra series are deleted, and the
//     leftover seed rows are blanked so a 3-point table does not drag a stray 4th point.
//   * SeriesCollection().NewSeries() fails (E_FAIL), so more series than the seed allows
//     is reported, not silently dropped.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

internal static partial class Program
{
    private static readonly string[] ChartTypeNames = new string[]
    {
        "column", "bar", "line", "pie", "area", "doughnut", "scatter"
    };

    /// <summary>PptChartType constants WPS accepted in the probe.</summary>
    private static int ChartTypeConstant(string name)
    {
        switch ((name ?? "column").Trim().ToLowerInvariant())
        {
            case "column": case "col": case "bar": return 51;      // xlColumnClustered
            case "barh": case "hbar": return 57;                    // xlBarClustered
            case "line": return 4;                                  // xlLine
            case "pie": return 5;                                   // xlPie
            case "area": return 1;                                  // xlArea
            case "doughnut": case "donut": return 80;               // xlDoughnut
            case "scatter": case "xy": return 74;                   // xlXYScatter
            default:
                throw new Fail("BAD_ARGS", "unknown -Type \"" + name + "\"; use column|barh|line|pie|area|doughnut|scatter");
        }
    }

    private sealed class ChartData
    {
        public List<string> Categories = new List<string>();
        public List<string> SeriesNames = new List<string>();
        public List<List<double>> Values = new List<List<double>>();
    }

    private static ChartData ReadChartData(string path)
    {
        if (path == null || path.Length == 0) throw new Fail("BAD_ARGS", "chart needs -Data <json file>");
        string full = FullPathOf(path);
        if (!File.Exists(full)) throw new Fail("BAD_ARGS", "data file not found: " + full);
        object parsed;
        try { parsed = Json.Parse(File.ReadAllText(full, Encoding.UTF8)); }
        catch (Exception e) { throw new Fail("BAD_ARGS", "chart data is not valid JSON: " + Clean(e.Message)); }
        Bag root = new Bag(parsed);

        ChartData data = new ChartData();
        foreach (object o in root.List("categories"))
            data.Categories.Add(Convert.ToString(o, CultureInfo.InvariantCulture));

        foreach (object o in root.List("series"))
        {
            Bag s = new Bag(o);
            data.SeriesNames.Add(s.Str("name", "series " + (data.SeriesNames.Count + 1)));
            List<double> values = new List<double>();
            foreach (object v in s.List("values"))
                values.Add(Convert.ToDouble(v, CultureInfo.InvariantCulture));
            data.Values.Add(values);
        }

        // a single flat "values" array is accepted too (one series, no names)
        if (data.Values.Count == 0 && root.Has("values"))
        {
            List<double> values = new List<double>();
            foreach (object v in root.List("values"))
                values.Add(Convert.ToDouble(v, CultureInfo.InvariantCulture));
            data.SeriesNames.Add("series 1");
            data.Values.Add(values);
        }
        if (data.Values.Count == 0) throw new Fail("BAD_ARGS", "chart data needs \"series\":[{\"name\":..,\"values\":[..]}] (or a flat \"values\")");
        if (data.Categories.Count == 0)
        {
            for (int i = 0; i < data.Values[0].Count; i++)
                data.Categories.Add((i + 1).ToString(CultureInfo.InvariantCulture));
        }
        foreach (List<double> list in data.Values)
            if (list.Count != data.Categories.Count)
                throw new Fail("BAD_ARGS", "every series needs one value per category (" + data.Categories.Count + " expected, got " + list.Count + ")");
        return data;
    }

    private static object ChartCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        if (n == 0) throw new Fail("BAD_ARGS", "this presentation has no slides");

        string action = (a.Str("action", "add") ?? "add").ToLowerInvariant();
        if (action == "list") return ChartList(p, n, ParseSlideSet(a.Str("slides", null), n));

        if (action != "add") throw new Fail("BAD_ARGS", "chart -Action must be add|list");
        int slideNo = a.Int("slide", 0);
        if (slideNo < 1 || slideNo > n) throw new Fail("BAD_ARGS", "chart needs -Slide 1.." + n);
        dynamic slide = SlideAt(p, slideNo);

        ChartData data = ReadChartData(a.Str("data", null));
        int constant = ChartTypeConstant(a.Str("type", "column"));

        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        double w = a.Dbl("w", 0), h = a.Dbl("h", 0);
        if (w <= 0) w = W * 0.55;
        if (h <= 0) h = H * 0.45;
        double l = a.Dbl("l", -1), t = a.Dbl("t", -1);
        if (l < 0) l = (W - w) / 2;
        if (t < 0) t = (H - h) / 2;

        dynamic shape = AddChartShape(slide, constant, l, t, w, h);
        dynamic chart = shape.Chart;

        int seededSeries;
        int seriesWritten = WriteChartData(chart, data, out seededSeries);
        ApplyChartDecorations(chart, a, data);

        int readSeries = 0;
        try { readSeries = (int)chart.SeriesCollection().Count; } catch (Exception) { }
        List<object> valuesRead = new List<object>();
        try
        {
            // Values comes back as a Variant array; casting straight to object[] drops it
            object raw = chart.SeriesCollection(1).Values;
            Array arr = raw as Array;
            if (arr != null)
            {
                foreach (object v in arr)
                    valuesRead.Add(Math.Round(Convert.ToDouble(v, CultureInfo.InvariantCulture), 2));
            }
        }
        catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slide"] = slideNo;
        r["shape"] = Convert.ToString(TryGet(() => (object)shape.Name, ""), CultureInfo.InvariantCulture);
        r["chartType"] = constant;
        r["series"] = readSeries;
        r["categories"] = new List<object>(data.Categories.ToArray());
        r["firstSeriesValues"] = valuesRead;
        if (LastChartRange != null) { r["rangeNote"] = LastChartRange; LastChartRange = null; }
        if (seededSeries > 0 && data.Values.Count > seededSeries)
            r["warning"] = "WPS only seeds " + seededSeries + " series and NewSeries is not supported: the extra series in "
                         + "your data were not written (" + (data.Values.Count - seededSeries) + " dropped)";
        r["note"] = "the data rides in the chart's embedded workbook, so the reader can edit the numbers in place "
                  + "(this is the difference from pasting a picture of a chart). Render the slide before moving on.";
        return r;
    }

    /// <summary>A chart can only be created in a presentation that has a window; retry once
    /// after showing the app, then say plainly what to do.</summary>
    private static dynamic AddChartShape(dynamic slide, int constant, double l, double t, double w, double h)
    {
        try
        {
            return slide.Shapes.AddChart(constant, (float)l, (float)t, (float)w, (float)h);
        }
        catch (Exception first)
        {
            dynamic app = App(false, false);
            if (app != null) { try { app.Visible = -1; } catch (Exception) { } }
            Thread.Sleep(400);
            try
            {
                return slide.Shapes.AddChart(constant, (float)l, (float)t, (float)w, (float)h);
            }
            catch (Exception)
            {
                throw new Fail("NO_WINDOW", "WPS can only create a chart in a presentation that has a window "
                    + "(AddChart fails with E_FAIL otherwise, and making the app visible afterwards does not help). "
                    + "Reopen the deck with `open -Show` and try again. Original error: " + Clean(first.Message));
            }
        }
    }

    private static int WriteChartData(dynamic chart, ChartData data, out int seededSeries)
    {
        seededSeries = 0;
        dynamic ws = null;
        for (int attempt = 0; attempt < 14 && ws == null; attempt++)
        {
            try { chart.ChartData.Activate(); } catch (Exception) { }
            Thread.Sleep(400);
            try
            {
                dynamic wb = chart.ChartData.Workbook;
                if (wb != null) ws = wb.ActiveSheet;
            }
            catch (Exception) { ws = null; }
        }
        if (ws == null)
            throw new Fail("COM_ERROR", "the chart's embedded workbook did not appear; the chart exists with WPS's sample "
                + "data, so delete it (shape delete) or try again. This happens when the spreadsheet part of WPS is busy.");

        try { seededSeries = (int)chart.SeriesCollection().Count; } catch (Exception) { seededSeries = 0; }

        int rows = data.Categories.Count + 1;              // header + points
        int existingRows = 0;
        try { existingRows = (int)ws.UsedRange.Rows.Count; } catch (Exception) { existingRows = rows; }

        // The chart is bound to the seeded range and SetSourceData is refused, so the only
        // lever is the sheet itself: inserting a row inside the range extends the plot,
        // deleting the rows past our data shrinks it.
        bool rangeAdjusted = false;
        try
        {
            if (rows > existingRows)
            {
                for (int i = 0; i < rows - existingRows; i++) ws.Rows(existingRows + i).Insert();
                rangeAdjusted = true;
            }
            else if (rows < existingRows)
            {
                for (int r = rows + 1; r <= existingRows; r++) ws.Rows(rows + 1).Delete();
                rangeAdjusted = true;
            }
        }
        catch (Exception) { rangeAdjusted = false; }

        if (!rangeAdjusted && existingRows > rows)
        {
            // could not delete rows: at least blank the tail so it is not plotted as zeros
            for (int r = rows + 1; r <= existingRows; r++)
            {
                for (int c = 1; c <= Math.Max(data.Values.Count, seededSeries) + 1; c++)
                {
                    try { ws.Cells(r, c).Value = ""; } catch (Exception) { }
                }
            }
        }

        for (int i = 0; i < data.Values.Count; i++)
        {
            try { ws.Cells(1, i + 2).Value = data.SeriesNames[i]; } catch (Exception) { }
            if (i >= seededSeries) break;                    // NewSeries is not supported
        }
        for (int c = 0; c < data.Categories.Count; c++)
        {
            try { ws.Cells(c + 2, 1).Value = data.Categories[c]; } catch (Exception) { }
        }
        for (int s = 0; s < data.Values.Count && s < Math.Max(seededSeries, 1); s++)
        {
            for (int c = 0; c < data.Categories.Count; c++)
            {
                try { ws.Cells(c + 2, s + 2).Value = data.Values[s][c]; } catch (Exception) { }
            }
        }

        // drop the series the seed brought and we do not want
        int wanted = Math.Min(data.Values.Count, Math.Max(seededSeries, 1));
        // Best effort only: WPS refused SetSourceData in every probe, the row adjustment
        // above is what actually moves the plot.
        try
        {
            dynamic rng = ws.Range(ws.Cells(1, 1), ws.Cells(rows, wanted + 1));
            chart.SetSourceData(rng);
        }
        catch (Exception) { }
        try
        {
            while ((int)chart.SeriesCollection().Count > wanted)
                chart.SeriesCollection((int)chart.SeriesCollection().Count).Delete();
        }
        catch (Exception) { }

        try { chart.ChartData.Workbook.Close(); } catch (Exception) { }
        int plotted = 0;
        try { plotted = ((Array)chart.SeriesCollection(1).Values).Length; } catch (Exception) { }
        if (plotted != data.Categories.Count)
        {
            LastChartRange = "the chart plots " + plotted + " points but the data has " + data.Categories.Count
                           + ": WPS keeps its own data range and refused SetSourceData, and the sheet rows could not be "
                           + "adjusted either. Delete the chart and try again, or use fewer points.";
        }
        return wanted;
    }

    /// <summary>Set while writing a chart when SetSourceData was refused.</summary>
    private static string LastChartRange;

    private static void ApplyChartDecorations(dynamic chart, Bag a, ChartData data)
    {
        string title = a.Str("title", null);
        if (title != null && title.Length > 0)
        {
            try { chart.HasTitle = true; chart.ChartTitle.Text = title; } catch (Exception) { }
        }
        string xTitle = a.Str("xtitle", null);
        if (xTitle != null && xTitle.Length > 0)
        {
            try { chart.Axes(2).HasTitle = true; chart.Axes(2).AxisTitle.Text = xTitle; } catch (Exception) { }
        }
        string yTitle = a.Str("ytitle", null);
        if (yTitle != null && yTitle.Length > 0)
        {
            try { chart.Axes(1).HasTitle = true; chart.Axes(1).AxisTitle.Text = yTitle; } catch (Exception) { }
        }
        bool legend = a.Flag("legend", data.Values.Count > 1);
        try { chart.HasLegend = legend; } catch (Exception) { }
    }

    // ---------------------------------------------------------------- list

    private static object ChartList(dynamic p, int n, HashSet<int> only)
    {
        List<object> slides = new List<object>();
        int total = 0;
        for (int k = 1; k <= n; k++)
        {
            if (only != null && !only.Contains(k)) continue;
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            int count = 0;
            try { count = (int)s.Shapes.Count; } catch (Exception) { }
            List<object> charts = new List<object>();
            for (int i = 1; i <= count; i++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(i); } catch (Exception) { continue; }
                bool has = false;
                try { has = (int)sh.HasChart != 0; } catch (Exception) { has = false; }
                if (!has) continue;
                Dictionary<string, object> entry = new Dictionary<string, object>();
                entry["name"] = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                entry["chartType"] = Convert.ToInt32(TryGet(() => (object)sh.Chart.ChartType, (object)0), CultureInfo.InvariantCulture);
                entry["series"] = Convert.ToInt32(TryGet(() => (object)sh.Chart.SeriesCollection().Count, (object)0), CultureInfo.InvariantCulture);
                entry["l"] = R(ToDouble(TryGet(() => (object)sh.Left, (object)0.0)));
                entry["t"] = R(ToDouble(TryGet(() => (object)sh.Top, (object)0.0)));
                entry["w"] = R(ToDouble(TryGet(() => (object)sh.Width, (object)0.0)));
                entry["h"] = R(ToDouble(TryGet(() => (object)sh.Height, (object)0.0)));
                charts.Add(entry);
            }
            if (charts.Count == 0) continue;
            total += charts.Count;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["slide"] = k;
            one["charts"] = charts;
            slides.Add(one);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["total"] = total;
        r["slides"] = slides;
        if (total == 0) r["note"] = "no native charts on this deck";
        return r;
    }
}
