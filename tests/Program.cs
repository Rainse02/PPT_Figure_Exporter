using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml.Linq;
using PPTFigureExporter;

internal static class Program
{
    private static int assertions;

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string output = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "output"));
            Directory.CreateDirectory(output);
            TestSizeModel();
            TestRibbon();
            if (args.Contains("--office"))
                TestOffice(output);
            if (args.Contains("--render-ui"))
                RenderUi(output);
            Console.WriteLine("PASS: " + assertions + " assertions");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static void Assert(bool condition, string message)
    {
        assertions++;
        if (!condition)
            throw new Exception("FAILED: " + message);
    }

    private static void Near(double actual, double expected, string message)
    {
        Assert(Math.Abs(actual - expected) < 0.0001, message + ": " + actual + " != " + expected);
    }

    private static void TestSizeModel()
    {
        var options = new ExportSettings { Dpi = 300 };
        var size = options.Target(144, 72);
        Near(size.Width, 144, "Keep width in pt");
        Assert(size.PixelWidth == 600 && size.PixelHeight == 300, "Keep 2x1 inch at 300 DPI");
        foreach (var unit in new[] { "mm", "cm", "inch" })
        {
            options.Unit = unit;
            options.Width = unit == "mm" ? 25.4 : unit == "cm" ? 2.54 : 1;
            options.SizeMode = "width";
            size = options.Target(144, 72);
            Near(size.Width, 72, "Unit conversion " + unit);
            Near(size.Height, 36, "Preserve ratio " + unit);
        }
        options.Unit = "inch";
        options.SizeMode = "height";
        options.Height = 2;
        size = options.Target(144, 72);
        Near(size.Width, 288, "Fit height maintains aspect ratio");
        options.SizeMode = "custom";
        options.Width = 3;
        size = options.Target(144, 72);
        Assert(size.PixelWidth == 900 && size.PixelHeight == 600, "Custom exact size");
        options.SizeMode = "scale";
        options.Percent = 50;
        size = options.Target(144, 72);
        Assert(size.PixelWidth == 300 && size.PixelHeight == 150, "50 percent changes both dimensions");
        options.SizeMode = "keep";
        options.Dpi = 72;
        size = options.Target(2.5, 3.5);
        Assert(size.PixelWidth == 2 && size.PixelHeight == 4, "Ties-to-even rounding");
        bool rejected = false;
        try { options.Target(100000, 100000); }
        catch (ArgumentException) { rejected = true; }
        Assert(rejected, "Reject allocation over 80 million pixels");
        options.Format = "pdf";
        size = options.Target(144, 72);
        Assert(size.PixelWidth == 0 && size.PixelHeight == 0, "PDF has no output pixel size");
        options.Format = "jpeg";
        Assert(!options.SupportsAlpha && options.UsesQuality, "JPEG white background and quality capability");
    }

    private static void TestRibbon()
    {
        Type[] contracts = typeof(Connect).GetInterfaces();
        Assert(!typeof(Connect).Assembly.GetReferencedAssemblies().Any(a => a.Name == "office" || a.Name == "Extensibility"),
            "Office interop contracts are embedded without a target-machine PIA dependency");
        var addin = new Connect();
        IntPtr unknown = Marshal.GetIUnknownForObject(addin);
        try
        {
            foreach (Type contract in contracts.Where(t => t.Name == "IDTExtensibility2" || t.Name == "IRibbonExtensibility" || t.Name == "IRibbonCallbacks"))
            {
                Guid id = contract.GUID;
                IntPtr pointer;
                int result = Marshal.QueryInterface(unknown, ref id, out pointer);
                Assert(result == 0, "COM interface is exposed: " + contract.Name);
                Marshal.Release(pointer);
            }
        }
        finally { Marshal.Release(unknown); }
        var document = XDocument.Parse(addin.GetCustomUI("Microsoft.PowerPoint.Presentation"));
        string[] callbackAttributes = { "onLoad", "onAction", "onChange", "getLabel", "getText", "getSelectedItemID", "getEnabled", "getPressed" };
        foreach (var attribute in document.Descendants().Attributes().Where(a => callbackAttributes.Contains(a.Name.LocalName)))
            Assert(typeof(Connect).GetMethod(attribute.Value) != null, "Ribbon callback exists: " + attribute.Value);
        Assert(document.Descendants().Any(e => (string)e.Attribute("id") == "ExportFigure"), "Export button present");
    }

    private static void TestOffice(string output)
    {
        if (Process.GetProcessesByName("POWERPNT").Length > 0)
            throw new InvalidOperationException("Close PowerPoint before the isolated Office integration test.");
        dynamic app = null;
        dynamic presentation = null;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            app.Visible = -1;
            presentation = app.Presentations.Add(-1);
            dynamic slide = presentation.Slides.Add(1, 12);
            dynamic frame = slide.Shapes.AddShape(1, 40F, 40F, 288F, 180F);
            frame.Name = "Frame";
            frame.Fill.Visible = 0;
            frame.Line.ForeColor.RGB = 0x333333;
            frame.Line.Weight = 0.5F;
            dynamic title = slide.Shapes.AddTextbox(1, 54F, 48F, 260F, 35F);
            title.Name = "GreekText";
            title.TextFrame.TextRange.Text = "Arial: α β γ Δ μ ± 10⁻³";
            title.TextFrame.TextRange.Font.Name = "Arial";
            title.TextFrame.TextRange.Font.Size = 16F;
            title.TextFrame.TextRange.Font.Color.RGB = 0x222222;
            dynamic circle = slide.Shapes.AddShape(9, 70F, 110F, 64F, 64F);
            circle.Name = "TransparentCircle";
            circle.Fill.ForeColor.RGB = 0xDD8822;
            circle.Fill.Transparency = 0.45F;
            circle.Line.Visible = 0;
            dynamic second = slide.Shapes.AddShape(9, 110F, 110F, 64F, 64F);
            second.Name = "NestedCircle";
            second.Fill.ForeColor.RGB = 0x44AA66;
            second.Fill.Transparency = 0.2F;
            second.Line.Visible = 0;
            dynamic nested = slide.Shapes.Range(new object[] { "TransparentCircle", "NestedCircle" }).Group();
            nested.Name = "NestedGroup";
            dynamic dashed = slide.Shapes.AddLine(60F, 195F, 210F, 195F);
            dashed.Name = "DashedLine";
            dashed.Line.DashStyle = 4;
            dashed.Line.Weight = 0.75F;
            dynamic arrow = slide.Shapes.AddLine(195F, 115F, 245F, 170F);
            arrow.Name = "Arrow";
            arrow.Line.Weight = 1F;
            arrow.Line.EndArrowheadStyle = 3;
            string bitmapPath = Path.Combine(output, "embedded.png");
            using (var bitmap = new Bitmap(16, 16))
            {
                for (int y = 0; y < 16; y++)
                    for (int x = 0; x < 16; x++)
                        bitmap.SetPixel(x, y, Color.FromArgb(255, x * 16, y * 16, 128));
                bitmap.Save(bitmapPath);
            }
            dynamic picture = slide.Shapes.AddPicture(bitmapPath, 0, -1, 260F, 100F, 40F, 40F);
            picture.Name = "RasterImage";
            string embeddedSvg = Path.Combine(output, "embedded.svg");
            File.WriteAllText(embeddedSvg, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40\" height=\"40\" viewBox=\"0 0 40 40\"><defs><clipPath id=\"clip\"><circle cx=\"20\" cy=\"20\" r=\"17\"/></clipPath></defs><g clip-path=\"url(#clip)\"><rect width=\"40\" height=\"40\" fill=\"#d97352\"/><path d=\"M0 0L40 40M0 40L40 0\" stroke=\"#253d70\" stroke-width=\"6\"/></g></svg>", new UTF8Encoding(false));
            dynamic svgPicture = slide.Shapes.AddPicture(embeddedSvg, 0, -1, 260F, 158F, 40F, 40F);
            svgPicture.Name = "EmbeddedSvg";
            dynamic group = slide.Shapes.Range(new object[] { "Frame", "GreekText", "NestedGroup", "DashedLine", "Arrow", "RasterImage", "EmbeddedSvg" }).Group();
            group.Name = "ScientificFigure";
            dynamic unrelated = slide.Shapes.AddShape(1, 400F, 250F, 150F, 100F);
            unrelated.Name = "MustNotExport";
            unrelated.Fill.ForeColor.RGB = 0x0000FF;
            string sourcePath = Path.Combine(output, "scientific-fixture.pptx");
            presentation.SaveAs(sourcePath, 24);
            title.TextFrame.TextRange.Text = "Arial: α β γ Δ μ ± 10⁻³";
            presentation.Saved = 0;
            group.Select(-1);
            string before = Snapshot(presentation, group, unrelated);
            using (var selected = SelectedFigure.Read((object)app))
            {
                string folder = Path.Combine(output, "office-intermediate");
                Directory.CreateDirectory(folder);
                string svg = Path.Combine(folder, "figure.svg");
                selected.ExportSvg(svg);
                Assert(File.Exists(svg), "PowerPoint Shape.Export generates SVG");
                var vector = XDocument.Load(svg);
                Assert(vector.Descendants().Any(e => e.Name.LocalName == "path"), "Figure retains vector paths");
                Assert(vector.Descendants().Any(e => e.Name.LocalName == "text"), "Editable SVG text retained");
                Assert(vector.Descendants().Any(e => e.Name.LocalName == "image"), "Embedded raster retained");
                Assert(before == Snapshot(presentation, group, unrelated), "Source presentation, selection geometry, groups and saved state unchanged");
                foreach (string format in new[] { "svg", "pdf", "png", "tiff", "jpeg", "webp" })
                {
                    var options = new ExportSettings { Format = format, Dpi = 300, Transparent = true };
                    string destination = Path.Combine(output, "office-figure." + options.Extension);
                    string request = ExportService.WriteRequest(folder, svg, destination, selected, options, true);
                    var result = ExportService.RunAsync(request, CancellationToken.None).GetAwaiter().GetResult();
                    Assert(result.ok && File.Exists(destination), "Office fixture exports " + format);
                    Near(result.width_pt, selected.WidthPt, format + " physical width");
                    Near(result.height_pt, selected.HeightPt, format + " physical height");
                }
                group.Export(Path.Combine(output, "powerpoint-reference.png"), 2, 1920, 1080);
                File.WriteAllText(Path.Combine(output, "office-result.json"), new JavaScriptSerializer().Serialize(new
                {
                    office_version = (string)app.Version,
                    width_pt = selected.WidthPt,
                    height_pt = selected.HeightPt,
                    source_unchanged = before == Snapshot(presentation, group, unrelated),
                    formats = new[] { "svg", "pdf", "png", "tiff", "jpeg", "webp" }
                }), new UTF8Encoding(false));
            }
            slide.Shapes.Range(new object[] { "ScientificFigure", "MustNotExport" }).Select(-1);
            bool multipleRejected = false;
            try { using (SelectedFigure.Read((object)app)) { } }
            catch (InvalidOperationException ex) { multipleRejected = ex.Message.Contains("Ctrl+G"); }
            Assert(multipleRejected, "Multiple shapes require grouping");
            Console.WriteLine("Office integration output: " + output);
        }
        finally
        {
            if (presentation != null)
            {
                presentation.Saved = -1;
                presentation.Close();
                Marshal.ReleaseComObject((object)presentation);
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject((object)app);
            }
        }
    }

    private static string Snapshot(dynamic presentation, dynamic group, dynamic unrelated)
    {
        var values = new List<string>
        {
            Convert.ToString(presentation.Saved, CultureInfo.InvariantCulture),
            Convert.ToString(presentation.Slides.Count, CultureInfo.InvariantCulture)
        };
        Action<dynamic> capture = shape =>
        {
            foreach (object value in new object[] { shape.Id, shape.Left, shape.Top, shape.Width, shape.Height, shape.Rotation, shape.Name })
                values.Add(Convert.ToString(value, CultureInfo.InvariantCulture));
        };
        capture(group);
        capture(unrelated);
        for (int index = 1; index <= (int)group.GroupItems.Count; index++)
            capture(group.GroupItems[index]);
        return string.Join("|", values);
    }

    private static void RenderUi(string output)
    {
        Application.EnableVisualStyles();
        foreach (string format in new[] { "tiff", "pdf", "webp" })
        {
            using (var dialog = new SettingsDialog(new ExportSettings { Format = format }, 288, 180))
            {
                dialog.Show();
                Application.DoEvents();
                using (var bitmap = new Bitmap(dialog.Width, dialog.Height))
                {
                    dialog.DrawToBitmap(bitmap, new Rectangle(0, 0, dialog.Width, dialog.Height));
                    bitmap.Save(Path.Combine(output, "settings-" + format + ".png"));
                }
                dialog.Close();
            }
        }
    }
}
