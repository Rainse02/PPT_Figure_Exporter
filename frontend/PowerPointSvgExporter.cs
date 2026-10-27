using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Xml;
using System.Windows.Media.Imaging;

namespace PPTFigureExporter
{
    internal static class PowerPointSvgExporter
    {
        private const string SvgNamespace = "http://www.w3.org/2000/svg";

        public static void Export(object shape, object presentation, int slideId, string path)
        {
            ExportDirect(shape, path);
            XmlDocument document = ReadSvg(path);
            List<XmlAttribute> images = ImageReferences(document);
            if (images.Count == 0)
                return;

            string snapshot = null;
            try
            {
                string sourcePath = (string)((dynamic)presentation).FullName;
                string extension = Path.GetExtension(sourcePath).ToLowerInvariant();
                bool savedPackage = (int)((dynamic)presentation).Saved == -1 && File.Exists(sourcePath) &&
                    new[] { ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm" }.Contains(extension);
                if (!savedPackage)
                {
                    // SaveCopyAs captures unsaved image insertions without changing the source document.
                    snapshot = Path.Combine(Path.GetDirectoryName(path), Guid.NewGuid().ToString("N") + ".pptx");
                    ((dynamic)presentation).SaveCopyAs(snapshot, 24, 0);
                    sourcePath = snapshot;
                }
                List<NativePicture> originals = PresentationPictures.Read(sourcePath, slideId, (int)((dynamic)shape).Id);
                NativePictureEffects.Render(presentation, sourcePath, slideId, originals, Path.GetDirectoryName(path));
                var replacements = new Dictionary<string, NativePicture>();
                foreach (NativePicture original in originals)
                {
                    string key = PictureReference(shape, original, path, document);
                    if (key == null)
                        continue;
                    NativePicture previous;
                    if (replacements.TryGetValue(key, out previous) && !(previous.RenderedData ?? previous.Data).SequenceEqual(original.RenderedData ?? original.Data))
                        throw new InvalidOperationException("PowerPoint 将不同原图合并成了同一预览，无法可靠恢复原始图片。");
                    replacements[key] = original;
                }

                var restored = new HashSet<string>();
                var payloads = new Dictionary<string, string>();
                foreach (XmlAttribute image in images)
                {
                    string key = ImageKey(image.Value);
                    NativePicture original;
                    if (!replacements.TryGetValue(key, out original))
                        continue;
                    string payload;
                    string payloadKey = original.PartName + (original.RenderedData == null ? "" : "#" + original.ShapeId);
                    if (!payloads.TryGetValue(payloadKey, out payload))
                    {
                        payload = DataUri(original);
                        payloads.Add(payloadKey, payload);
                    }
                    image.Value = payload;
                    restored.Add(key);
                }
                if (restored.Count != replacements.Count)
                    throw new InvalidOperationException("无法将 SVG 中的图片与 PPT 原图准确对应，已停止导出以避免使用低清预览。");
                if (restored.Count > 0)
                {
                    // Keep image viewports, clip paths and transforms: only replace the image payload.
                    document.Save(path);
                }
            }
            finally
            {
                DeleteTemporary(snapshot);
            }
        }

        private static string PictureReference(object root, NativePicture picture, string path, XmlDocument rootDocument)
        {
            object current = root;
            bool ownsCurrent = false;
            string temporary = null;
            try
            {
                if ((int)((dynamic)root).Id != picture.ShapeId)
                {
                    current = FindPicture(root, picture.ShapeId, picture.ShapeName);
                    ownsCurrent = true;
                    if (current == null)
                        throw new InvalidOperationException("无法在当前选区中定位原始图片。");
                }
                if ((int)((dynamic)current).Id != picture.ShapeId)
                    throw new InvalidOperationException("PPT 图片结构与导出选区不一致。");
                XmlDocument document = rootDocument;
                if (ownsCurrent)
                {
                    temporary = Path.Combine(Path.GetDirectoryName(path), Guid.NewGuid().ToString("N") + ".svg");
                    ExportDirect(current, temporary);
                    document = ReadSvg(temporary);
                }
                string[] keys = ImageReferences(document).Select(image => ImageKey(image.Value)).Distinct().ToArray();
                if (keys.Length > 1)
                    throw new InvalidOperationException("该图片导出为多个独立位图，无法直接恢复原图。");
                return keys.Length == 1 ? keys[0] : null;
            }
            finally
            {
                if (ownsCurrent)
                    Release(current);
                DeleteTemporary(temporary);
            }
        }

        private static object FindPicture(object root, int id, string name)
        {
            if ((int)((dynamic)root).Type != 6)
                return null;
            object children = ((dynamic)root).GroupItems;
            try
            {
                // GroupItems can flatten a hierarchy that remains nested in the saved XML.
                // Names provide a fast lookup; IDs determine whether the match is correct.
                object named = null;
                try
                {
                    named = ((dynamic)children)[name];
                    if ((int)((dynamic)named).Id == id)
                    {
                        object found = named;
                        named = null;
                        return found;
                    }
                }
                catch (Exception ex) when (ex is COMException || ex is ArgumentException) { }
                finally { Release(named); }

                for (int index = 1; index <= (int)((dynamic)children).Count; index++)
                {
                    object child = ((dynamic)children)[index];
                    try
                    {
                        if ((int)((dynamic)child).Id == id)
                        {
                            object found = child;
                            child = null;
                            return found;
                        }
                        object nested = FindPicture(child, id, name);
                        if (nested != null)
                            return nested;
                    }
                    finally { Release(child); }
                }
                return null;
            }
            finally { Release(children); }
        }

        private static string DataUri(NativePicture picture)
        {
            // These SVG-native bitmap containers are embedded byte-for-byte. Detect actual
            // file content so the result does not depend on a producer's extension or MIME alias.
            byte[] data = picture.RenderedData ?? picture.Data;
            string type = null;
            if (data.Length >= 8 && data[0] == 137 && data[1] == 80 && data[2] == 78 && data[3] == 71 &&
                data[4] == 13 && data[5] == 10 && data[6] == 26 && data[7] == 10)
                type = "image/png";
            else if (data.Length >= 3 && data[0] == 255 && data[1] == 216 && data[2] == 255)
                type = "image/jpeg";
            else if (data.Length >= 6 && data[0] == 71 && data[1] == 73 && data[2] == 70 && data[3] == 56 &&
                (data[4] == 55 || data[4] == 57) && data[5] == 97)
                type = "image/gif";
            if (type != null)
                return "data:" + type + ";base64," + Convert.ToBase64String(data);

            // WIC decodes other bitmap containers at their stored pixel size. Preserve pixel
            // format and bit depth where PNG supports them; no thumbnail or resize is requested.
            using (var input = new MemoryStream(data, false))
            using (var output = new MemoryStream())
            {
                BitmapFrame image;
                try
                {
                    image = BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                }
                catch (Exception ex) when (ex is NotSupportedException || ex is System.IO.FileFormatException)
                {
                    throw new InvalidOperationException("无法解码 PPT 原始位图：" + picture.PartName + " (" + picture.MediaType + ")。", ex);
                }
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(image);
                encoder.Save(output);
                return "data:image/png;base64," + Convert.ToBase64String(output.ToArray());
            }
        }

        private static string ImageKey(string value)
        {
            int comma = value.IndexOf(',');
            byte[] bytes = Convert.FromBase64String(value.Substring(comma + 1));
            using (SHA256 algorithm = SHA256.Create())
                return Convert.ToBase64String(algorithm.ComputeHash(bytes));
        }

        private static List<XmlAttribute> ImageReferences(XmlDocument document)
        {
            var result = new List<XmlAttribute>();
            foreach (XmlNode image in document.GetElementsByTagName("image", SvgNamespace))
            {
                foreach (XmlAttribute attribute in image.Attributes)
                {
                    if (attribute.LocalName == "href" && attribute.Value.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase) &&
                        !attribute.Value.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase))
                        result.Add(attribute);
                }
            }
            return result;
        }

        private static void ExportDirect(object shape, string path)
        {
            try { ((dynamic)shape).Export(path, 6); }
            catch (COMException ex)
            {
                throw new InvalidOperationException("PowerPoint 无法通过自动化接口导出 SVG。请使用支持该接口的新版 PowerPoint，并检查所选图形内容。", ex);
            }
        }

        private static XmlDocument ReadSvg(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length == 0)
                throw new IOException("PowerPoint 未生成有效的 SVG 文件。");
            try
            {
                var options = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
                using (var reader = XmlReader.Create(path, options))
                    document.Load(reader);
                if (document.DocumentElement == null || document.DocumentElement.LocalName != "svg" ||
                    document.DocumentElement.NamespaceURI != SvgNamespace)
                    throw new XmlException("根元素不是 SVG。");
                return document;
            }
            catch (XmlException ex)
            {
                throw new InvalidOperationException("此 PowerPoint 自动化接口没有生成真实 SVG。手动另存为 SVG 与自动化导出的版本支持可能不同，请更新 Office。", ex);
            }
        }

        private static void DeleteTemporary(string path)
        {
            if (path == null)
                return;
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                SettingsStore.Log("清理原图恢复临时文件失败。", ex);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
