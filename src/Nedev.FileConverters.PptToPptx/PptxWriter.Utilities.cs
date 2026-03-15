using System;
using System.IO;
using System.Xml;

namespace Nedev.FileConverters.PptToPptx
{
    public partial class PptxWriter
    {
        /// <summary>
        /// Converts master units (1/576 inch) to EMU (English Metric Units).
        /// 1 inch = 914400 EMU, so 1 master unit = 914400/576 = 1587.5 EMU
        /// </summary>
        internal static long MasterUnitsToEmu(long masterUnits)
        {
            long n = masterUnits * 914400L;
            long d = 576L;
            return n >= 0 ? (n + (d / 2)) / d : (n - (d / 2)) / d;
        }

        /// <summary>
        /// Converts master units to points * 100 (hundredths of a point).
        /// 1 master unit = 1/576 inch; 1 inch = 72pt
        /// </summary>
        internal static int MasterUnitsToPoints100(int masterUnits)
        {
            int n = masterUnits * 25;
            return n >= 0 ? (n + 1) / 2 : (n - 1) / 2;
        }

        /// <summary>
        /// Writes a relationship element to the XML writer.
        /// </summary>
        internal static void WriteRelationship(XmlWriter writer, string id, string type, string target)
        {
            writer.WriteStartElement("Relationship", NS_RELS);
            writer.WriteAttributeString("Id", id);
            writer.WriteAttributeString("Type", type);
            writer.WriteAttributeString("Target", target);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a default content type entry.
        /// </summary>
        internal static void WriteDefaultContentType(XmlWriter writer, string extension, string contentType)
        {
            writer.WriteStartElement("Default", NS_CT);
            writer.WriteAttributeString("Extension", extension);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes an override content type entry.
        /// </summary>
        internal static void WriteOverrideContentType(XmlWriter writer, string partName, string contentType)
        {
            writer.WriteStartElement("Override", NS_CT);
            writer.WriteAttributeString("PartName", partName);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Creates a new XML writer with standard settings.
        /// </summary>
        internal static XmlWriter CreateXmlWriter(string path)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                NewLineChars = "\n",
                NewLineHandling = NewLineHandling.Replace,
                OmitXmlDeclaration = false
            };
            return XmlWriter.Create(path, settings);
        }

        /// <summary>
        /// Determines if the output should be macro-enabled based on presentation content.
        /// </summary>
        internal static bool IsMacroEnabled(Presentation presentation)
        {
            return presentation?.VbaProject?.ProjectData != null && presentation.VbaProject.ProjectData.Length > 0;
        }

        /// <summary>
        /// Gets the appropriate file extension for an image content type.
        /// </summary>
        internal static string GetImageExtension(string? contentType)
        {
            return contentType?.ToLowerInvariant() switch
            {
                "image/png" => "png",
                "image/jpeg" => "jpg",
                "image/gif" => "gif",
                "image/bmp" => "bmp",
                "image/tiff" => "tiff",
                "image/x-emf" => "emf",
                "image/x-wmf" => "wmf",
                _ => "png"
            };
        }

        /// <summary>
        /// Sanitizes a string for use in XML content.
        /// </summary>
        internal static string SanitizeForXml(string? input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            var result = new System.Text.StringBuilder(input.Length);
            foreach (char c in input)
            {
                if (IsValidXmlChar(c))
                    result.Append(c);
            }
            return result.ToString();
        }

        /// <summary>
        /// Checks if a character is valid for XML 1.0.
        /// </summary>
        private static bool IsValidXmlChar(char c)
        {
            return c == 0x9 || c == 0xA || c == 0xD ||
                   (c >= 0x20 && c <= 0xD7FF) ||
                   (c >= 0xE000 && c <= 0xFFFD);
        }

        /// <summary>
        /// Ensures a directory exists, creating it if necessary.
        /// </summary>
        internal static void EnsureDirectoryExists(string path)
        {
            string? directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }
    }
}
