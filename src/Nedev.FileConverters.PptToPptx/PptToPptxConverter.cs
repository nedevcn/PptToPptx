using System;
using System.IO;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// Provides methods for converting legacy PowerPoint (.ppt) files to modern OpenXML (.pptx) format.
    /// </summary>
    public static class PptToPptxConverter
    {
        /// <summary>
        /// Converts a .ppt file to .pptx format.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <exception cref="ArgumentException">Thrown when path arguments are invalid.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
        public static void Convert(string pptPath, string pptxPath)
        {
            Convert(pptPath, pptxPath, null);
        }

        /// <summary>
        /// Converts a .ppt file to .pptx format with optional configuration.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <exception cref="ArgumentException">Thrown when path arguments are invalid.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
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

            options?.ReportProgress(ConversionPhase.Initializing, 0, "Starting conversion...");

            try
            {
                Presentation presentation;

                // Read phase
                options?.ReportProgress(ConversionPhase.Reading, 10, "Reading PPT file...");
                using (var pptReader = new PptReader(pptPath, options))
                {
                    presentation = pptReader.ReadPresentation();
                }

                options?.ReportProgress(ConversionPhase.ProcessingStructure, 30, $"Found {presentation.Slides.Count} slides...", 0, presentation.Slides.Count);

                // Write phase
                options?.ReportProgress(ConversionPhase.Writing, 50, "Writing PPTX file...");
                using (var pptxWriter = new PptxWriter(pptxPath, options))
                {
                    pptxWriter.WritePresentation(presentation);
                }

                options?.ReportProgress(ConversionPhase.Finalizing, 90, "Finalizing...");
                options?.ReportProgress(ConversionPhase.Completed, 100, "Conversion completed successfully.", presentation.Slides.Count, presentation.Slides.Count);
            }
            catch (Exception ex)
            {
                options?.ReportProgress(ConversionPhase.Failed, 0, $"Conversion failed: {ex.Message}");
                throw;
            }
        }
    }
}
