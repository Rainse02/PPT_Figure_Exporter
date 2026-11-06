using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using PPTFigureExporter;

internal static class NativePdfTests
{
    public static void Run(string output, Action<bool, string> assert)
    {
        string folder = Path.Combine(output, "pdf-quality");
        Directory.CreateDirectory(folder);
        string png = Path.Combine(folder, "high-detail-2048x1024.png");
        CreateDetailImage(png);
        dynamic app = null;
        dynamic presentation = null;
        object slide = null;
        var objects = new List<object>();
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            presentation = app.Presentations.Add(0);
            object pageSetup = presentation.PageSetup;
            try
            {
                ((dynamic)pageSetup).SlideWidth = 720F;
                ((dynamic)pageSetup).SlideHeight = 540F;
            }
            finally
            {
                Release(pageSetup);
            }
            slide = presentation.Slides.Add(1, 12);
            dynamic small = Own(((dynamic)slide).Shapes.AddShape(1, 500F, 400F, 24F, 18F), objects);
            small.Name = "SmallObject";
            small.Fill.ForeColor.RGB = 0x886633;

            dynamic unrelated = Own(((dynamic)slide).Shapes.AddShape(1, 580F, 440F, 120F, 50F), objects);
            unrelated.Name = "UnselectedObject";
            unrelated.TextFrame.TextRange.Text = "UNSELECTED_SENTINEL";
            unrelated.TextFrame.TextRange.Font.Size = 10F;
            string sourcePath = Path.Combine(folder, "pdf-quality-source.pptx");
            if (File.Exists(sourcePath))
                File.Delete(sourcePath);
            presentation.SaveAs(sourcePath, 24);
            assert((int)presentation.Saved == -1, "Small-object fixture uses an already saved PPTX");
            string unrelatedState = NativeRasterTests.Snapshot((object)unrelated, (object)presentation);
            int slideId = (int)((dynamic)slide).SlideID;

            CheckShape((object)small, (object)presentation, slideId, folder, "small-object",
                new[] { new ExportSettings { Format = "pdf" } }, true, assert);
            assert(NativeRasterTests.Snapshot((object)unrelated, (object)presentation) == unrelatedState,
                "Saved-document PDF export preserves unselected source objects");

            // Insert and edit after saving, so the PDF must include current in-memory content.
            object group = CreateGroup(slide, png, objects);
            object children = ((dynamic)group).GroupItems;
            object editedText = null;
            try
            {
                editedText = ((dynamic)children)["EquationText"];
                ((dynamic)editedText).TextFrame.TextRange.Font.Bold = -1;
            }
            finally
            {
                Release(editedText);
                Release(children);
            }
            assert((int)presentation.Saved == 0, "Mixed-group fixture includes unsaved insertions and edits");
            unrelatedState = NativeRasterTests.Snapshot((object)unrelated, (object)presentation);
            CheckShape(group, (object)presentation, slideId, folder, "mixed-group",
                new[]
                {
                    new ExportSettings { Format = "pdf" },
                    new ExportSettings { Format = "pdf", SizeMode = "width", Unit = "inch", Width = 2.75 },
                    new ExportSettings
                    {
                        Format = "pdf",
                        SizeMode = "custom",
                        Unit = "inch",
                        Width = 3.25,
                        Height = 1.25
                    }
                }, false, assert);

            assert(NativeRasterTests.Snapshot((object)unrelated, (object)presentation) == unrelatedState,
                "PDF export preserves unselected source objects");
            string[] presentationFiles = Directory.GetFiles(folder, "*.pptx", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal)).ToArray();
            assert(presentationFiles.Length == 1 && File.Exists(sourcePath),
                "PDF export removes its temporary presentation copies");
            Console.WriteLine("Native PDF regression output: " + folder);
        }
        finally
        {
            if (presentation != null)
            {
                presentation.Saved = -1;
                presentation.Close();
            }
            for (int index = objects.Count - 1; index >= 0; index--)
                Release(objects[index]);
            Release(slide);
            Release((object)presentation);
            Release((object)app);
        }
    }

    private static object CreateGroup(object slide, string png, List<object> objects)
    {
        dynamic shapes = ((dynamic)slide).Shapes;
        try
        {
            dynamic text = Own(shapes.AddTextbox(1, 40F, 40F, 280F, 45F), objects);
            text.Name = "EquationText";
            text.TextFrame.TextRange.Text = "E = mc2; alpha α ± beta β; 10⁻³";
            text.TextFrame.TextRange.Font.Name = "Arial";
            text.TextFrame.TextRange.Font.Size = 18F;
            text.TextFrame.TextRange.Characters(7, 1).Font.Superscript = -1;

            dynamic line = Own(shapes.AddLine(40F, 105F, 310F, 105F), objects);
            line.Name = "FineLine";
            line.Line.Weight = 0.35F;

            dynamic first = Own(shapes.AddShape(1, 40F, 125F, 96F, 50F), objects);
            first.Name = "TransparentRectangle";
            first.Fill.ForeColor.RGB = 0xCC6600;
            first.Fill.Transparency = 0.3F;
            first.Line.Visible = 0;
            dynamic second = Own(shapes.AddShape(9, 80F, 140F, 96F, 50F), objects);
            second.Name = "TransparentEllipse";
            second.Fill.ForeColor.RGB = 0x0099CC;
            second.Fill.Transparency = 0.4F;
            second.Line.Visible = 0;
            dynamic nested = Own(shapes.Range(new object[]
            {
                "TransparentRectangle", "TransparentEllipse"
            }).Group(), objects);
            nested.Name = "NestedTransparency";

            dynamic picture = Own(shapes.AddPicture(png, 0, -1, 40F, 220F, 288F, 144F), objects);
            picture.Name = "HighDetailPicture";
            dynamic shadowed = Own(shapes.AddShape(1, 350F, 110F, 90F, 50F), objects);
            shadowed.Name = "RotatedShadow";
            shadowed.Rotation = 23F;
            shadowed.Shadow.Visible = -1;
            shadowed.Shadow.OffsetX = 5F;
            shadowed.Shadow.OffsetY = 4F;
            shadowed.Shadow.Blur = 4F;

            dynamic group = Own(shapes.Range(new object[]
            {
                "EquationText", "FineLine", "NestedTransparency", "HighDetailPicture", "RotatedShadow"
            }).Group(), objects);
            group.Name = "NativePdfFixture";
            return (object)group;
        }
        finally
        {
            Release((object)shapes);
        }
    }

    private static void CheckShape(object shape, object presentation, int slideId, string folder,
        string name, ExportSettings[] modes, bool smallObject, Action<bool, string> assert)
    {
        string input = Path.Combine(folder, name + "-native.pdf");
        string before = NativeRasterTests.Snapshot(shape, presentation);
        double width = (double)((dynamic)shape).Width;
        double height = (double)((dynamic)shape).Height;
        var referenceSettings = new ExportSettings { Format = "png", Dpi = 144 };
        PowerPointRasterExporter.Export(shape, presentation, Path.Combine(folder, name + "-reference.png"),
            referenceSettings.Target(width, height), false);
        PdfCrop crop = PowerPointPdfExporter.Export(shape, presentation, slideId, input);
        assert(IsPdf(input), name + ": PowerPoint generates a native PDF");
        assert(NativeRasterTests.Snapshot(shape, presentation) == before,
            name + ": native PDF preserves source geometry, saved state and slide dimensions");
        assert(Finite(crop.left_pt) && Finite(crop.top_pt) &&
            Finite(crop.width_pt) && crop.width_pt > 0 &&
            Finite(crop.height_pt) && crop.height_pt > 0,
            name + ": native PDF provides a finite, positive crop rectangle");

        foreach (ExportSettings settings in modes)
        {
            SizePoints target = settings.Target(width, height);
            string destination = Path.Combine(folder, name + "-" + settings.SizeMode + ".pdf");
            string request = Path.Combine(folder, name + "-" + settings.SizeMode + "-request.json");
            File.WriteAllText(request, new JavaScriptSerializer().Serialize(new
            {
                input_path = input,
                output_path = destination,
                format = "pdf",
                source_width_pt = width,
                source_height_pt = height,
                size_mode = settings.SizeMode,
                width = settings.Width,
                height = settings.Height,
                unit = settings.Unit,
                percent = settings.Percent,
                dpi = settings.Dpi,
                transparent = settings.Transparent,
                overwrite = true,
                pdf_crop = crop
            }), new UTF8Encoding(false));
            CoreResult result = ExportService.RunAsync(request, CancellationToken.None).GetAwaiter().GetResult();
            string label = name + " " + settings.SizeMode;
            assert(result.ok && IsPdf(destination), label + ": core writes a PDF");
            assert(Math.Abs(result.width_pt - target.Width) < 0.0001 &&
                Math.Abs(result.height_pt - target.Height) < 0.0001,
                label + ": output has the requested physical dimensions");
            assert(result.pixel_width == null && result.pixel_height == null,
                label + ": PDF result has no raster pixel dimensions");
            if (smallObject)
                assert(result.width_pt < 72 && result.height_pt < 72,
                    "Final small-object PDF remains below one inch in both dimensions");
            assert(NativeRasterTests.Snapshot(shape, presentation) == before,
                label + ": exporting preserves the source presentation");
        }
    }

    private static void CreateDetailImage(string path)
    {
        using (var image = new Bitmap(2048, 1024))
        using (Graphics graphics = Graphics.FromImage(image))
        {
            graphics.Clear(Color.White);
            for (int x = 0; x < image.Width; x += 8)
                graphics.FillRectangle(Brushes.Black, x, 0, 4, image.Height);
            graphics.FillRectangle(Brushes.DarkRed, 0, image.Height / 2, image.Width, 2);
            image.Save(path, ImageFormat.Png);
        }
    }

    private static bool IsPdf(string path)
    {
        if (!File.Exists(path))
            return false;
        using (var input = File.OpenRead(path))
        {
            var header = new byte[5];
            return input.Read(header, 0, header.Length) == header.Length &&
                Encoding.ASCII.GetString(header) == "%PDF-";
        }
    }

    private static bool Finite(double value)
    {
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static object Own(object value, List<object> objects)
    {
        objects.Add(value);
        return value;
    }

    private static void Release(object value)
    {
        if (value != null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
