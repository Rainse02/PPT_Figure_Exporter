using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Media.Imaging;
using System.Xml;

namespace PPTFigureExporter
{
    internal static class NativePictureEffects
    {
        private sealed class RenderTarget
        {
            public NativePicture Picture;
            public int Width;
            public int Height;
            public string Name { get { return "NativeBitmap" + Picture.ShapeId.ToString(CultureInfo.InvariantCulture); } }
        }

        public static void Render(object presentation, string sourcePath, int slideId,
            List<NativePicture> pictures, string parent)
        {
            var targets = pictures.Select(CreateTarget).Where(target => target != null).ToArray();
            if (targets.Length == 0)
                return;

            string folder = Path.Combine(parent, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string snapshot = Path.Combine(folder, "image-effects.pptx");
            object app = null;
            object presentations = null;
            object copy = null;
            object slides = null;
            object slide = null;
            object shapes = null;
            try
            {
                if (string.Equals(Path.GetExtension(sourcePath), ".pptx", StringComparison.OrdinalIgnoreCase))
                    File.Copy(sourcePath, snapshot);
                else
                    ((dynamic)presentation).SaveCopyAs(snapshot, 24, 0);
                PrepareSnapshot(snapshot, slideId, targets);
                app = ((dynamic)presentation).Application;
                presentations = ((dynamic)app).Presentations;
                copy = ((dynamic)presentations).Open(snapshot, -1, 0, 0);
                slides = ((dynamic)copy).Slides;
                slide = ((dynamic)slides).FindBySlideID(slideId);
                shapes = ((dynamic)slide).Shapes;
                foreach (RenderTarget target in targets)
                {
                    string png = Path.Combine(folder, target.Name + ".png");
                    object shape = ((dynamic)shapes)[target.Name];
                    try
                    {
                        int side = Math.Max(target.Width, target.Height);
                        ((dynamic)shape).Export(png, 2, side, side, 1);
                    }
                    finally { Release(shape); }
                    byte[] data = File.ReadAllBytes(png);
                    using (var input = new MemoryStream(data, false))
                    {
                        BitmapFrame frame = BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                        if (frame.PixelWidth != target.Width || frame.PixelHeight != target.Height)
                            throw new InvalidOperationException("PowerPoint 图片校正输出尺寸为 " + frame.PixelWidth + " x " + frame.PixelHeight + "，期望 " + target.Width + " x " + target.Height + "。");
                    }
                    target.Picture.RenderedData = data;
                }
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        ((dynamic)copy).Saved = -1;
                        ((dynamic)copy).Close();
                    }
                    catch (COMException ex) { SettingsStore.Log("关闭图片校正临时文稿失败。", ex); }
                }
                Release(shapes);
                Release(slide);
                Release(slides);
                Release(copy);
                Release(presentations);
                Release(app);
                try { Directory.Delete(folder, true); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    SettingsStore.Log("清理图片校正临时文件失败。", ex);
                }
            }
        }

        private static RenderTarget CreateTarget(NativePicture picture)
        {
            using (var input = new MemoryStream(picture.Data, false))
            {
                BitmapFrame frame = BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.None);
                bool nativeColor = frame.Format.ToString().StartsWith("Cmyk", StringComparison.OrdinalIgnoreCase) ||
                    (frame.ColorContexts != null && frame.ColorContexts.Count != 0);
                if (!picture.HasImageEffects && !nativeColor)
                    return null;
                return new RenderTarget { Picture = picture, Width = frame.PixelWidth, Height = frame.PixelHeight };
            }
        }

        private static void PrepareSnapshot(string path, int slideId, RenderTarget[] targets)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var package = new ZipArchive(file, ZipArchiveMode.Update))
            {
                // Set the rendering canvas in package XML before PowerPoint opens it,
                // so changing slide size cannot scale the normalized source pictures.
                XmlDocument presentation = PresentationPictures.ReadXml(package, "ppt/presentation.xml");
                XmlNamespaceManager presentationNamespaces = PresentationPictures.Namespaces(presentation);
                XmlElement size = (XmlElement)presentation.SelectSingleNode("/p:presentation/p:sldSz", presentationNamespaces);
                size.SetAttribute("cx", "1219200");
                size.SetAttribute("cy", "1219200");
                size.RemoveAttribute("type");
                WriteXml(package, "ppt/presentation.xml", presentation);
                string slidePart = PresentationPictures.FindSlidePart(package, slideId);
                XmlDocument document = PresentationPictures.ReadXml(package, slidePart);
                XmlNamespaceManager namespaces = PresentationPictures.Namespaces(document);
                XmlNode tree = document.SelectSingleNode("/p:sld/p:cSld/p:spTree", namespaces);
                foreach (XmlNode node in tree.ChildNodes.Cast<XmlNode>().ToArray())
                {
                    if (node.NodeType == XmlNodeType.Element && node.LocalName != "nvGrpSpPr" && node.LocalName != "grpSpPr")
                        tree.RemoveChild(node);
                }
                foreach (RenderTarget target in targets)
                    tree.AppendChild(CreatePicture(document, namespaces, target));
                WriteXml(package, slidePart, document);

                // A format conversion or a save-time compression setting must not substitute
                // a reduced source image in the disposable processing document.
                foreach (var picture in targets.Select(target => target.Picture).GroupBy(picture => picture.PartName).Select(group => group.First()))
                    WriteBytes(package, picture.PartName, picture.Data);
            }
        }

        private static XmlElement CreatePicture(XmlDocument document, XmlNamespaceManager namespaces, RenderTarget target)
        {
            string p = namespaces.LookupNamespace("p");
            string a = namespaces.LookupNamespace("a");
            XmlElement picture = document.CreateElement("p", "pic", p);
            XmlElement nonVisual = Append(picture, "p", "nvPicPr", p);
            XmlElement properties = Append(nonVisual, "p", "cNvPr", p);
            properties.SetAttribute("id", target.Picture.ShapeId.ToString(CultureInfo.InvariantCulture));
            properties.SetAttribute("name", target.Name);
            Append(nonVisual, "p", "cNvPicPr", p);
            Append(nonVisual, "p", "nvPr", p);
            XmlElement fill = Append(picture, "p", "blipFill", p);
            var source = new XmlDocument { XmlResolver = null };
            source.LoadXml(target.Picture.BlipXml);
            fill.AppendChild(document.ImportNode(source.DocumentElement, true));
            Append(Append(fill, "a", "stretch", a), "a", "fillRect", a);
            XmlElement shape = Append(picture, "p", "spPr", p);
            XmlElement transform = Append(shape, "a", "xfrm", a);
            XmlElement offset = Append(transform, "a", "off", a);
            offset.SetAttribute("x", "0");
            offset.SetAttribute("y", "0");
            XmlElement extent = Append(transform, "a", "ext", a);
            double side = Math.Max(target.Width, target.Height);
            extent.SetAttribute("cx", Math.Round(914400 * target.Width / side).ToString(CultureInfo.InvariantCulture));
            extent.SetAttribute("cy", Math.Round(914400 * target.Height / side).ToString(CultureInfo.InvariantCulture));
            XmlElement geometry = Append(shape, "a", "prstGeom", a);
            geometry.SetAttribute("prst", "rect");
            Append(geometry, "a", "avLst", a);
            return picture;
        }

        private static XmlElement Append(XmlElement parent, string prefix, string name, string ns)
        {
            XmlElement element = parent.OwnerDocument.CreateElement(prefix, name, ns);
            parent.AppendChild(element);
            return element;
        }

        private static void WriteXml(ZipArchive package, string part, XmlDocument document)
        {
            ZipArchiveEntry previous = package.GetEntry(part);
            if (previous != null)
                previous.Delete();
            using (Stream output = package.CreateEntry(part).Open())
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
                document.Save(writer);
        }

        private static void WriteBytes(ZipArchive package, string part, byte[] bytes)
        {
            ZipArchiveEntry previous = package.GetEntry(part);
            if (previous != null)
                previous.Delete();
            using (Stream output = package.CreateEntry(part).Open())
                output.Write(bytes, 0, bytes.Length);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
