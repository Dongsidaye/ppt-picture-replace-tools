using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Dongsidaye.PowerPointTools
{
    [ComVisible(false)]
    internal sealed class PictureSourceCrop
    {
        public double l { get; set; }
        public double t { get; set; }
        public double r { get; set; }
        public double b { get; set; }
    }

    [ComVisible(false)]
    internal sealed class PictureSourceEntry
    {
        public string fingerprint { get; set; }
        public string path { get; set; }
        public string ext { get; set; }
        public PictureSourceCrop crop { get; set; }
        public PictureSourceDisplay display { get; set; }
        public string appearance { get; set; }
    }

    [ComVisible(false)]
    internal sealed class PictureSourceDisplay
    {
        public double width { get; set; }
        public double height { get; set; }
        public double rotation { get; set; }
        public bool flipH { get; set; }
        public bool flipV { get; set; }
    }

    // Reads only an owned snapshot. No Office COM calls, changes to the source
    // presentation, external relationships, or clipboard operations occur here.
    [ComVisible(false)]
    internal static class PictureSourceManifest
    {
        private const long MediaLimit = 32L * 1024 * 1024;
        private const long XmlLimit = 4L * 1024 * 1024;
        private const long TotalLimit = 256L * 1024 * 1024;
        private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

        private sealed class Relationship
        {
            internal string Target, Type;
        }

        private sealed class Reader
        {
            private readonly Dictionary<string, ZipArchiveEntry> entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
            private readonly Dictionary<string, XDocument> xml = new Dictionary<string, XDocument>(StringComparer.Ordinal);
            private readonly Dictionary<string, Dictionary<string, Relationship>> relationships = new Dictionary<string, Dictionary<string, Relationship>>(StringComparer.Ordinal);
            private readonly Dictionary<string, PictureSourceEntry> media = new Dictionary<string, PictureSourceEntry>(StringComparer.Ordinal);
            private readonly HashSet<string> consumed = new HashSet<string>(StringComparer.Ordinal);
            private readonly string output;
            private long total;
            private string styleAppearance;

            internal Reader(ZipArchive archive, string outputDirectory)
            {
                output = Path.GetFullPath(outputDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                Directory.CreateDirectory(output);
                // The caller owns this directory; never follow an existing link
                // from that boundary into another location.
                if ((File.GetAttributes(output) & FileAttributes.ReparsePoint) != 0) throw new IOException("Snapshot output directory is a reparse point.");
                if (archive.Entries.Count > 50000) throw new InvalidDataException("Snapshot has too many package entries.");
                var ambiguous = new HashSet<string>(StringComparer.Ordinal);
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string normalized = NormalizePart(entry.FullName, false);
                    if (normalized == null || normalized != entry.FullName || ambiguous.Contains(normalized)) continue;
                    if (entries.ContainsKey(normalized)) { entries.Remove(normalized); ambiguous.Add(normalized); }
                    else entries.Add(normalized, entry);
                }
            }

            private byte[] ReadBytes(string part, long limit)
            {
                ZipArchiveEntry entry;
                if (part == null || !entries.TryGetValue(part, out entry) || entry.Length < 1 || entry.Length > limit) return null;
                if (!consumed.Contains(part))
                {
                    if (entry.Length > TotalLimit - total) return null;
                    total += entry.Length;
                    consumed.Add(part);
                }
                using (Stream input = entry.Open())
                using (var buffer = new MemoryStream((int)entry.Length))
                {
                    byte[] chunk = new byte[65536];
                    int count;
                    while ((count = input.Read(chunk, 0, chunk.Length)) > 0)
                    {
                        if (buffer.Length + count > entry.Length || buffer.Length + count > limit) throw new InvalidDataException("Package entry exceeded its declared size.");
                        buffer.Write(chunk, 0, count);
                    }
                    return buffer.Length == entry.Length ? buffer.ToArray() : null;
                }
            }

            internal XDocument ReadXml(string part)
            {
                if (part == null) return null;
                XDocument document;
                if (xml.TryGetValue(part, out document)) return document;
                document = null;
                try
                {
                    byte[] bytes = ReadBytes(part, XmlLimit);
                    if (bytes != null)
                    {
                        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = XmlLimit };
                        using (var stream = new MemoryStream(bytes, false))
                        using (XmlReader reader = XmlReader.Create(stream, settings)) document = XDocument.Load(reader, LoadOptions.None);
                    }
                }
                catch (Exception ex) { if (!Recoverable(ex)) throw; }
                xml[part] = document;
                return document;
            }

            private string StyleAppearance()
            {
                if (styleAppearance != null) return styleAppearance;
                bool complete = true;
                // An inherited placeholder/theme effect can change while its
                // p:pic remains byte-for-byte identical. A conservative global
                // style revision avoids reusing that picture's old native pixels.
                // ReadXml shares the same XML and cumulative expansion limits.
                using (SHA256 sha = SHA256.Create())
                {
                    foreach (string part in entries.Keys.Where(x =>
                        ((x.StartsWith("ppt/theme/", StringComparison.Ordinal) || x.StartsWith("ppt/slideLayouts/", StringComparison.Ordinal) ||
                          x.StartsWith("ppt/slideMasters/", StringComparison.Ordinal)) && x.EndsWith(".xml", StringComparison.Ordinal)) ||
                        (x.StartsWith("ppt/", StringComparison.Ordinal) && x.EndsWith(".rels", StringComparison.Ordinal))).OrderBy(x => x, StringComparer.Ordinal))
                    {
                        XDocument document = ReadXml(part);
                        if (document == null) complete = false;
                        string content = document == null ? "unverified" : document.ToString(SaveOptions.DisableFormatting);
                        byte[] bytes = Encoding.UTF8.GetBytes(part.Length + ":" + part + ":" + content.Length + ":" + content);
                        sha.TransformBlock(bytes, 0, bytes.Length, bytes, 0);
                    }
                    sha.TransformFinalBlock(new byte[0], 0, 0);
                    styleAppearance = BitConverter.ToString(sha.Hash).Replace("-", "").ToLowerInvariant();
                }
                // Unreadable/over-budget styles cannot prove cached appearance.
                // Keep identities, but make every new snapshot fail cache reuse.
                if (!complete) styleAppearance += ":unverified:" + Guid.NewGuid().ToString("N");
                return styleAppearance;
            }

            private Dictionary<string, Relationship> Relationships(string owner)
            {
                Dictionary<string, Relationship> result;
                if (relationships.TryGetValue(owner, out result)) return result;
                result = new Dictionary<string, Relationship>(StringComparer.Ordinal);
                relationships[owner] = result;
                int slash = owner.LastIndexOf('/');
                string folder = slash >= 0 ? owner.Substring(0, slash + 1) : "";
                string name = owner.Substring(slash + 1);
                XDocument document = ReadXml(folder + "_rels/" + name + ".rels");
                if (document == null || document.Root == null || document.Root.Name != Rel + "Relationships") return result;
                var ambiguous = new HashSet<string>(StringComparer.Ordinal);
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (XElement item in document.Root.Elements(Rel + "Relationship"))
                {
                    string id = (string)item.Attribute("Id");
                    if (String.IsNullOrEmpty(id) || ambiguous.Contains(id)) continue;
                    if (!seen.Add(id)) { result.Remove(id); ambiguous.Add(id); continue; }
                    if (String.Equals((string)item.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    string target = ResolvePart(owner, (string)item.Attribute("Target"));
                    string type = (string)item.Attribute("Type");
                    if (target != null && type != null) result.Add(id, new Relationship { Target = target, Type = type });
                }
                return result;
            }

            internal string RelatedPart(string owner, XElement element, string kind)
            {
                if (element == null) return null;
                return RelatedPart(owner, (string)element.Attribute(R + "id"), kind);
            }

            private string RelatedPart(string owner, string id, string kind)
            {
                Relationship rel;
                if (id == null || !Relationships(owner).TryGetValue(id, out rel) || rel.Type != R.NamespaceName + "/" + kind) return null;
                return rel.Target;
            }

            private PictureSourceEntry ReadMedia(string part)
            {
                PictureSourceEntry result;
                if (part == null || !part.StartsWith("ppt/media/", StringComparison.Ordinal)) return null;
                if (media.TryGetValue(part, out result)) return result;
                result = null;
                try
                {
                    byte[] bytes = ReadBytes(part, MediaLimit);
                    if (bytes != null)
                    {
                        string hash;
                        using (SHA256 sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
                        string extension = Path.GetExtension(part).ToLowerInvariant();
                        if (!new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg", ".webp", ".wdp", ".jxr" }.Contains(extension)) extension = ".bin";
                        string destination = Path.GetFullPath(Path.Combine(output, hash + extension));
                        if (!destination.StartsWith(output, StringComparison.OrdinalIgnoreCase)) throw new IOException("Media output is outside its owned directory.");
                        if (File.Exists(destination))
                        {
                            if ((File.GetAttributes(destination) & FileAttributes.ReparsePoint) != 0) throw new IOException("Media output is a reparse point.");
                            using (var existing = File.OpenRead(destination))
                            using (SHA256 sha = SHA256.Create())
                                if (BitConverter.ToString(sha.ComputeHash(existing)).Replace("-", "").ToLowerInvariant() != hash) throw new IOException("Existing media output does not match.");
                        }
                        else
                        {
                            // CreateNew refuses an existing file/link; a malicious
                            // archive never supplies any filesystem destination.
                            using (var file = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)) file.Write(bytes, 0, bytes.Length);
                        }
                        result = new PictureSourceEntry { fingerprint = "sha256:" + hash, path = destination, ext = extension };
                    }
                }
                catch (Exception ex) { if (!Recoverable(ex)) throw; }
                media[part] = result;
                return result;
            }

            internal void AddPictures(string owner, string prefix, Dictionary<string, PictureSourceEntry> result)
            {
                XDocument document = ReadXml(owner);
                if (document == null || document.Root == null) return;
                XElement common = document.Root.Element(P + "cSld");
                XElement tree = common == null ? null : common.Element(P + "spTree");
                if (tree == null) return;
                var duplicateIds = new HashSet<string>(tree.Descendants(P + "cNvPr").GroupBy(x => (string)x.Attribute("id")).Where(x => x.Count() != 1).Select(x => x.Key), StringComparer.Ordinal);
                foreach (XElement picture in tree.Descendants(P + "pic"))
                {
                    XElement nonvisual = picture.Element(P + "nvPicPr");
                    XElement properties = nonvisual == null ? null : nonvisual.Element(P + "cNvPr");
                    string id = properties == null ? null : (string)properties.Attribute("id");
                    uint numericId;
                    if (id == null || !UInt32.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out numericId) || numericId == 0 || id != numericId.ToString(CultureInfo.InvariantCulture) || duplicateIds.Contains(id)) continue;
                    XElement fill = picture.Element(P + "blipFill");
                    XElement blip = fill == null ? null : fill.Element(A + "blip");
                    string embedded = blip == null ? null : (string)blip.Attribute(R + "embed");
                    PictureSourceCrop crop = ReadCrop(fill == null ? null : fill.Element(A + "srcRect"));
                    if (embedded == null || crop == null) continue;
                    // SVG/other alternate payloads can share one low-fidelity
                    // PNG fallback. Never identify those by that fallback alone.
                    if (blip.Descendants().Attributes(R + "embed").Any(x => x.Value != embedded)) continue;
                    PictureSourceEntry source = ReadMedia(RelatedPart(owner, embedded, "image"));
                    if (source == null) continue;
                    string key = prefix + numericId.ToString(CultureInfo.InvariantCulture);
                    XElement properties2 = picture.Element(P + "spPr");
                    XElement transform = properties2 == null ? null : properties2.Element(A + "xfrm");
                    XElement size = transform == null ? null : transform.Element(A + "ext");
                    PictureSourceDisplay display = null;
                    double width, height, rotation;
                    if (CanRenderSimplePicture(picture, fill, blip, properties2, transform) && size != null &&
                        Double.TryParse((string)size.Attribute("cx"), NumberStyles.Float, CultureInfo.InvariantCulture, out width) &&
                        Double.TryParse((string)size.Attribute("cy"), NumberStyles.Float, CultureInfo.InvariantCulture, out height) &&
                        IsFinite(width) && IsFinite(height) && width > 0 && height > 0)
                    {
                        Double.TryParse((string)transform.Attribute("rot"), NumberStyles.Float, CultureInfo.InvariantCulture, out rotation);
                        string flipH = (string)transform.Attribute("flipH"), flipV = (string)transform.Attribute("flipV");
                        display = new PictureSourceDisplay { width = width / 12700, height = height / 12700, rotation = rotation / 60000,
                            flipH = flipH == "1" || flipH == "true", flipV = flipV == "1" || flipV == "true" };
                    }
                    result[key] = new PictureSourceEntry { fingerprint = source.fingerprint, path = source.path, ext = source.ext,
                        crop = crop, display = display, appearance = AppearanceFingerprint(picture, StyleAppearance()) };
                }
            }
        }

        internal static Dictionary<string, PictureSourceEntry> Read(string snapshotPath, string ownedOutputDirectory)
        {
            if (String.IsNullOrWhiteSpace(snapshotPath) || String.IsNullOrWhiteSpace(ownedOutputDirectory)) throw new ArgumentException("Snapshot and owned output directory are required.");
            var result = new Dictionary<string, PictureSourceEntry>(StringComparer.Ordinal);
            using (var file = new FileStream(snapshotPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Read))
            {
                var reader = new Reader(archive, ownedOutputDirectory);
                const string presentation = "ppt/presentation.xml";
                XDocument document = reader.ReadXml(presentation);
                if (document == null || document.Root == null) return result;
                XElement slides = document.Root.Element(P + "sldIdLst");
                if (slides != null)
                {
                    var uniqueIds = new HashSet<string>(slides.Elements(P + "sldId").GroupBy(x => (string)x.Attribute("id")).Where(x => x.Count() == 1).Select(x => x.Key), StringComparer.Ordinal);
                    foreach (XElement slide in slides.Elements(P + "sldId"))
                    {
                        string id = (string)slide.Attribute("id"); uint numericId;
                        if (id == null || !uniqueIds.Contains(id) || !UInt32.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out numericId) || numericId == 0 || id != numericId.ToString(CultureInfo.InvariantCulture)) continue;
                        reader.AddPictures(reader.RelatedPart(presentation, slide, "slide"), "slide:" + numericId.ToString(CultureInfo.InvariantCulture) + ":", result);
                    }
                }
                XElement masters = document.Root.Element(P + "sldMasterIdLst");
                int masterIndex = 0;
                if (masters != null) foreach (XElement master in masters.Elements(P + "sldMasterId"))
                {
                    masterIndex++;
                    string owner = reader.RelatedPart(presentation, master, "slideMaster");
                    reader.AddPictures(owner, "master:" + masterIndex + ":", result);
                    XDocument masterDoc = reader.ReadXml(owner);
                    XElement layouts = masterDoc == null || masterDoc.Root == null ? null : masterDoc.Root.Element(P + "sldLayoutIdLst");
                    int layoutIndex = 0;
                    if (layouts != null) foreach (XElement layout in layouts.Elements(P + "sldLayoutId"))
                    {
                        layoutIndex++;
                        reader.AddPictures(reader.RelatedPart(owner, layout, "slideLayout"), "layout:" + masterIndex + ":" + layoutIndex + ":", result);
                    }
                }
            }
            return result;
        }

        private static bool Recoverable(Exception ex)
        {
            return ex is IOException || ex is UnauthorizedAccessException || ex is XmlException || ex is ArgumentException || ex is InvalidOperationException || ex is NotSupportedException;
        }

        // A source preview is an optimization, not a second OOXML renderer.
        // Unknown drawing semantics must use Shape.Export, which applies the
        // actual Office geometry, theme, image effects and group transforms.
        private static bool CanRenderSimplePicture(XElement picture, XElement fill, XElement blip, XElement properties, XElement transform)
        {
            if (picture.Parent == null || picture.Parent.Name != P + "spTree" || properties == null || transform == null) return false;
            if (!AllowedAttributes(picture) || picture.Elements().Any(x => x.Name != P + "nvPicPr" && x.Name != P + "blipFill" && x.Name != P + "spPr")) return false;
            if (picture.Elements(P + "blipFill").Count() != 1 || picture.Elements(P + "spPr").Count() != 1) return false;
            XElement nonvisual = picture.Element(P + "nvPicPr");
            if (nonvisual != null && nonvisual.Descendants(P + "ph").Any()) return false;
            if (!AllowedAttributes(fill, "dpi", "rotWithShape") || !TrueOrMissing((string)fill.Attribute("rotWithShape"))) return false;
            if (fill.Elements(A + "blip").Count() != 1 || fill.Elements(A + "srcRect").Count() > 1 || fill.Elements(A + "stretch").Count() > 1) return false;
            foreach (XElement child in fill.Elements())
            {
                if (child.Name == A + "blip") continue;
                if (child.Name == A + "srcRect")
                {
                    if (!AllowedAttributes(child, "l", "t", "r", "b") || child.HasElements) return false;
                }
                else if (child.Name == A + "stretch")
                {
                    if (!AllowedAttributes(child) || child.Elements().Any(x => x.Name != A + "fillRect") || child.Elements().Count() > 1) return false;
                    XElement rectangle = child.Element(A + "fillRect");
                    if (rectangle != null && (!AllowedAttributes(rectangle, "l", "t", "r", "b") || rectangle.HasElements ||
                        rectangle.Attributes().Where(x => !x.IsNamespaceDeclaration).Any(x => x.Value != "0"))) return false;
                }
                else return false;
            }
            if (blip.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != R + "embed" && x.Name != "cstate")) return false;
            // useLocalDpi only influences physical size, already fixed by xfrm.
            // Any other blip extension or color/alpha effect needs native pixels.
            foreach (XElement extensionList in blip.Elements())
            {
                if (extensionList.Name != A + "extLst" || !AllowedAttributes(extensionList)) return false;
                foreach (XElement extension in extensionList.Elements())
                {
                    if (extension.Name != A + "ext" || !AllowedAttributes(extension, "uri")) return false;
                    foreach (XElement setting in extension.Elements())
                    {
                        if (setting.Name != XName.Get("useLocalDpi", "http://schemas.microsoft.com/office/drawing/2010/main") ||
                            !AllowedAttributes(setting, "val") || setting.HasElements || !BooleanOrMissing((string)setting.Attribute("val"))) return false;
                    }
                }
            }
            if (!AllowedAttributes(properties) || properties.Elements(A + "xfrm").Count() != 1 ||
                !AllowedAttributes(transform, "rot", "flipH", "flipV") ||
                !BooleanOrMissing((string)transform.Attribute("flipH")) || !BooleanOrMissing((string)transform.Attribute("flipV"))) return false;
            double rotation;
            string rawRotation = (string)transform.Attribute("rot");
            if (rawRotation != null && (!Double.TryParse(rawRotation, NumberStyles.Float, CultureInfo.InvariantCulture, out rotation) || !IsFinite(rotation))) return false;
            if (transform.Elements().Any(x => x.Name != A + "off" && x.Name != A + "ext") || transform.Elements(A + "ext").Count() != 1 || transform.Elements(A + "off").Count() > 1) return false;
            foreach (XElement child in properties.Elements())
            {
                if (child.Name == A + "xfrm") continue;
                if (child.Name == A + "prstGeom")
                {
                    if (!AllowedAttributes(child, "prst") || (string)child.Attribute("prst") != "rect" ||
                        child.Elements().Any(x => x.Name != A + "avLst" || x.HasElements || !AllowedAttributes(x))) return false;
                }
                else if (child.Name == A + "noFill" || child.Name == A + "effectLst")
                {
                    if (child.HasElements || !AllowedAttributes(child)) return false;
                }
                else if (child.Name == A + "ln")
                {
                    if (child.Elements().Count() != 1 || child.Element(A + "noFill") == null || child.Element(A + "noFill").HasElements) return false;
                }
                else return false;
            }
            return true;
        }

        private static bool AllowedAttributes(XElement element, params string[] names)
        {
            return element != null && element.Attributes().All(x => x.IsNamespaceDeclaration || (x.Name.NamespaceName.Length == 0 && names.Contains(x.Name.LocalName)));
        }

        private static bool TrueOrMissing(string value) { return value == null || value == "1" || value == "true"; }
        private static bool BooleanOrMissing(string value) { return TrueOrMissing(value) || value == "0" || value == "false"; }
        private static bool IsFinite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }

        private static string AppearanceFingerprint(XElement picture, string styles)
        {
            var description = new StringBuilder(picture.ToString(SaveOptions.DisableFormatting));
            description.Append("|styles:").Append(styles);
            foreach (XElement group in picture.Ancestors(P + "grpSp"))
            {
                // Avoid repeating every sibling subtree for each nesting level.
                // Only ancestor drawing/nonvisual properties affect this object.
                foreach (XAttribute attribute in group.Attributes()) description.Append(attribute.ToString());
                foreach (XElement properties in group.Elements().Where(x => x.Name == P + "grpSpPr" || x.Name == P + "nvGrpSpPr" || x.Name == P + "style"))
                    description.Append(properties.ToString(SaveOptions.DisableFormatting));
            }
            using (SHA256 sha = SHA256.Create())
                return "sha256:" + BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(description.ToString()))).Replace("-", "").ToLowerInvariant();
        }

        private static PictureSourceCrop ReadCrop(XElement element)
        {
            double[] parts = new double[4];
            string[] names = { "l", "t", "r", "b" };
            for (int i = 0; i < names.Length; i++)
            {
                string raw = element == null ? null : (string)element.Attribute(names[i]);
                int value;
                if (raw != null && !Int32.TryParse(raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value)) return null;
                parts[i] = raw == null ? 0 : Int32.Parse(raw, CultureInfo.InvariantCulture) / 100000.0;
            }
            if (1 - parts[2] <= parts[0] || 1 - parts[3] <= parts[1]) return null;
            return new PictureSourceCrop { l = parts[0], t = parts[1], r = 1 - parts[2], b = 1 - parts[3] };
        }

        private static string ResolvePart(string owner, string target)
        {
            if (String.IsNullOrEmpty(target) || target.IndexOf('\\') >= 0 || target.IndexOf('#') >= 0 || target.IndexOf('?') >= 0 || target.StartsWith("//", StringComparison.Ordinal)) return null;
            try { target = Uri.UnescapeDataString(target); } catch (UriFormatException) { return null; }
            if (target.IndexOf(':') >= 0 || target.IndexOf('\\') >= 0 || target.IndexOf('\0') >= 0 || target.IndexOf('#') >= 0 || target.IndexOf('?') >= 0 || target.StartsWith("//", StringComparison.Ordinal)) return null;
            int slash = owner.LastIndexOf('/');
            return NormalizePart(target.StartsWith("/", StringComparison.Ordinal) ? target.Substring(1) : (slash >= 0 ? owner.Substring(0, slash + 1) : "") + target, true);
        }

        private static string NormalizePart(string part, bool allowParent)
        {
            if (String.IsNullOrEmpty(part) || part[0] == '/' || part.IndexOf('\\') >= 0 || part.IndexOf(':') >= 0 || part.IndexOf('\0') >= 0) return null;
            var segments = new List<string>();
            foreach (string segment in part.Split('/'))
            {
                if (segment == "..")
                {
                    if (!allowParent || segments.Count == 0) return null;
                    segments.RemoveAt(segments.Count - 1);
                }
                else if (segment == ".") { if (!allowParent) return null; }
                else if (segment.Length == 0) return null;
                else segments.Add(segment);
            }
            return segments.Count == 0 ? null : String.Join("/", segments);
        }
    }
}
