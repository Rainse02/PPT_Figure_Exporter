using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace PPTFigureExporter
{
    internal sealed class NativePicture
    {
        public int ShapeId;
        public string ShapeName;
        public string MediaType;
        public string PartName;
        public string BlipXml;
        public bool HasImageEffects;
        public byte[] RenderedData;
        public byte[] Data;
    }

    internal static class PresentationPictures
    {
        private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string SvgNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

        public static List<NativePicture> Read(string path, int slideId, int shapeId)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ZipArchive(file, ZipArchiveMode.Read))
            {
                string slidePart = FindSlidePart(package, slideId);
                XmlDocument document = ReadXml(package, slidePart);
                XmlNamespaceManager namespaces = Namespaces(document);
                XmlNode properties = document.SelectSingleNode("//p:cNvPr[@id='" + shapeId + "']", namespaces);
                if (properties == null)
                    throw new InvalidDataException("文稿中找不到当前选中图形。");
                XmlNode selected = properties.ParentNode.ParentNode;
                var pictures = new List<NativePicture>();
                Collect(package, selected, ReadRelationships(package, slidePart), ReadContentTypes(package), namespaces, pictures);
                return pictures;
            }
        }

        private static void Collect(ZipArchive package, XmlNode shape,
            Dictionary<string, string> relationships, Dictionary<string, string> contentTypes,
            XmlNamespaceManager namespaces, List<NativePicture> pictures)
        {
            XmlNode blip = shape.SelectSingleNode("p:blipFill/a:blip|p:spPr/a:blipFill/a:blip", namespaces);
            if (blip != null && blip.SelectSingleNode(".//asvg:svgBlip", namespaces) == null)
            {
                XmlAttribute embedded = blip.Attributes["embed", namespaces.LookupNamespace("r")];
                string part;
                if (embedded != null && relationships.TryGetValue(embedded.Value, out part))
                {
                    string mediaType = MediaType(part, contentTypes);
                    if (mediaType != null)
                    {
                        XmlNode properties = shape.SelectSingleNode("*/p:cNvPr", namespaces);
                        pictures.Add(new NativePicture
                        {
                            ShapeId = int.Parse(properties.Attributes["id"].Value, System.Globalization.CultureInfo.InvariantCulture),
                            ShapeName = properties.Attributes["name"].Value,
                            MediaType = mediaType,
                            PartName = part,
                            BlipXml = blip.OuterXml,
                            HasImageEffects = HasImageEffects(blip),
                            Data = ReadBytes(package, part)
                        });
                    }
                }
            }
            if (shape.LocalName != "grpSp")
                return;

            foreach (XmlNode child in shape.ChildNodes)
            {
                XmlNode candidate = child;
                if (candidate.LocalName == "AlternateContent")
                    candidate = candidate.SelectSingleNode("*/*[p:nvPicPr or p:nvSpPr or p:nvGrpSpPr or p:nvGraphicFramePr]", namespaces);
                if (candidate == null || candidate.SelectSingleNode("*/p:cNvPr", namespaces) == null)
                    continue;
                Collect(package, candidate, relationships, contentTypes, namespaces, pictures);
            }
        }

        internal static string FindSlidePart(ZipArchive package, int slideId)
        {
            XmlDocument presentation = ReadXml(package, "ppt/presentation.xml");
            XmlNamespaceManager namespaces = Namespaces(presentation);
            XmlNode slide = presentation.SelectSingleNode("/p:presentation/p:sldIdLst/p:sldId[@id='" + slideId + "']", namespaces);
            if (slide == null)
                throw new InvalidDataException("文稿中找不到当前幻灯片。");
            return ReadRelationships(package, "ppt/presentation.xml")[slide.Attributes["id", namespaces.LookupNamespace("r")].Value];
        }

        private static bool HasImageEffects(XmlNode blip)
        {
            foreach (XmlNode effect in blip.ChildNodes)
            {
                if (effect.NodeType != XmlNodeType.Element || effect.LocalName == "extLst")
                    continue;
                if (effect.LocalName == "alphaModFix" && effect.Attributes["amt"]?.Value == "100000")
                    continue;
                if (effect.LocalName == "lum" && (effect.Attributes["bright"]?.Value ?? "0") == "0" &&
                    (effect.Attributes["contrast"]?.Value ?? "0") == "0")
                    continue;
                return true;
            }
            return blip.SelectSingleNode(".//*[local-name()='imgEffect' or local-name()='imgEffectLst']") != null;
        }

        private static Dictionary<string, string> ReadRelationships(ZipArchive package, string part)
        {
            int separator = part.LastIndexOf('/');
            string relationshipsPart = part.Substring(0, separator + 1) + "_rels/" + part.Substring(separator + 1) + ".rels";
            XmlDocument document = ReadXml(package, relationshipsPart);
            var result = new Dictionary<string, string>();
            foreach (XmlNode relationship in document.DocumentElement.ChildNodes)
            {
                if (relationship.NodeType != XmlNodeType.Element || relationship.Attributes["TargetMode"]?.Value == "External")
                    continue;
                var uri = new Uri(new Uri("https://package.local/" + part), relationship.Attributes["Target"].Value);
                result.Add(relationship.Attributes["Id"].Value, Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/')));
            }
            return result;
        }

        private static Dictionary<string, string> ReadContentTypes(ZipArchive package)
        {
            XmlDocument document = ReadXml(package, "[Content_Types].xml");
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlNode node in document.DocumentElement.ChildNodes)
            {
                if (node.LocalName == "Default")
                    result["." + node.Attributes["Extension"].Value] = node.Attributes["ContentType"].Value;
                else if (node.LocalName == "Override")
                    result[node.Attributes["PartName"].Value.TrimStart('/')] = node.Attributes["ContentType"].Value;
            }
            return result;
        }

        private static string MediaType(string part, Dictionary<string, string> contentTypes)
        {
            string type;
            if (!contentTypes.TryGetValue(part, out type) && !contentTypes.TryGetValue(Path.GetExtension(part), out type))
                type = "application/octet-stream";
            // Leave native vector resources to PowerPoint's SVG writer. All bitmap containers
            // are handled by their content and Windows imaging codecs, rather than an extension list.
            string extension = Path.GetExtension(part).ToLowerInvariant();
            if (type == "image/svg+xml" || extension == ".svg" || extension == ".emf" || extension == ".wmf")
                return null;
            return type;
        }

        private static byte[] ReadBytes(ZipArchive package, string part)
        {
            ZipArchiveEntry entry = package.GetEntry(part);
            if (entry == null)
                throw new InvalidDataException("文稿缺少图片资源：" + part);
            using (Stream source = entry.Open())
            using (var output = new MemoryStream())
            {
                source.CopyTo(output);
                return output.ToArray();
            }
        }

        internal static XmlDocument ReadXml(ZipArchive package, string part)
        {
            ZipArchiveEntry entry = package.GetEntry(part);
            if (entry == null)
                throw new InvalidDataException("文稿缺少组件：" + part);
            using (Stream source = entry.Open())
            using (var reader = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            {
                var document = new XmlDocument { XmlResolver = null };
                document.Load(reader);
                return document;
            }
        }

        internal static XmlNamespaceManager Namespaces(XmlDocument document)
        {
            var namespaces = new XmlNamespaceManager(document.NameTable);
            bool strict = document.DocumentElement.NamespaceURI.StartsWith("http://purl.oclc.org/ooxml/", StringComparison.Ordinal);
            namespaces.AddNamespace("p", strict ? "http://purl.oclc.org/ooxml/presentationml/main" : PresentationNamespace);
            namespaces.AddNamespace("a", strict ? "http://purl.oclc.org/ooxml/drawingml/main" : DrawingNamespace);
            namespaces.AddNamespace("r", strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : RelationshipNamespace);
            namespaces.AddNamespace("asvg", SvgNamespace);
            return namespaces;
        }
    }
}
