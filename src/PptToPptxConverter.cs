using System;
using System.IO;

namespace Nefdev.PptToPptx
{
    public class PptToPptxConverter
    {
        public static void Convert(string pptPath, string pptxPath)
        {
            Convert(pptPath, pptxPath, null);
        }

        public static void Convert(string pptPath, string pptxPath, ConversionOptions? options)
        {
            if (string.IsNullOrWhiteSpace(pptPath))
                throw new ArgumentException("Input .ppt path must be provided.", nameof(pptPath));
            if (string.IsNullOrWhiteSpace(pptxPath))
                throw new ArgumentException("Output .pptx/.pptm path must be provided.", nameof(pptxPath));

            if (!File.Exists(pptPath))
                throw new FileNotFoundException("Input .ppt file not found.", pptPath);

            if (Path.GetFullPath(pptPath).Equals(Path.GetFullPath(pptxPath), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Output path must be different from input path.", nameof(pptxPath));

            var outDir = Path.GetDirectoryName(pptxPath);
            if (!string.IsNullOrEmpty(outDir))
                Directory.CreateDirectory(outDir);

            using var pptReader = new PptReader(pptPath, options);
            using var pptxWriter = new PptxWriter(pptxPath, options);
            
            pptxWriter.WritePresentation(pptReader.ReadPresentation());
        }
    }
}
