using System;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// Represents the progress of a conversion operation.
    /// </summary>
    public sealed class ConversionProgress
    {
        /// <summary>
        /// Gets the current phase of the conversion.
        /// </summary>
        public ConversionPhase Phase { get; }

        /// <summary>
        /// Gets the progress percentage (0-100).
        /// </summary>
        public int PercentComplete { get; }

        /// <summary>
        /// Gets the current operation message.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Gets the number of slides processed (if applicable).
        /// </summary>
        public int SlidesProcessed { get; }

        /// <summary>
        /// Gets the total number of slides (if applicable).
        /// </summary>
        public int TotalSlides { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionProgress"/> class.
        /// </summary>
        public ConversionProgress(ConversionPhase phase, int percentComplete, string message, int slidesProcessed = 0, int totalSlides = 0)
        {
            Phase = phase;
            PercentComplete = Math.Clamp(percentComplete, 0, 100);
            Message = message ?? string.Empty;
            SlidesProcessed = slidesProcessed;
            TotalSlides = totalSlides;
        }
    }

    /// <summary>
    /// Defines the phases of the conversion process.
    /// </summary>
    public enum ConversionPhase
    {
        /// <summary>Initial phase before conversion starts.</summary>
        Initializing,
        /// <summary>Reading and parsing the PPT file.</summary>
        Reading,
        /// <summary>Processing document structure.</summary>
        ProcessingStructure,
        /// <summary>Extracting images and media.</summary>
        ExtractingMedia,
        /// <summary>Processing slides.</summary>
        ProcessingSlides,
        /// <summary>Writing the PPTX file.</summary>
        Writing,
        /// <summary>Finalizing and packaging.</summary>
        Finalizing,
        /// <summary>Conversion completed successfully.</summary>
        Completed,
        /// <summary>Conversion failed.</summary>
        Failed
    }

    /// <summary>
    /// Options for configuring the PPT to PPTX conversion process.
    /// </summary>
    public sealed class ConversionOptions
    {
        /// <summary>
        /// Optional log sink for diagnostic messages. When null, the library stays silent.
        /// </summary>
        public Action<string>? Log { get; set; }

        /// <summary>
        /// Optional progress callback for reporting conversion progress.
        /// </summary>
        public Action<ConversionProgress>? Progress { get; set; }

        /// <summary>
        /// Keep a copy of the generated package directory next to the output file as "temp_pptx".
        /// Useful for debugging invalid packages. Default: false.
        /// </summary>
        public bool KeepTempFiles { get; set; } = false;

        /// <summary>
        /// Override the ANSI code page used for legacy PPT TextBytesAtom decoding.
        /// When null, uses the current culture's ANSI code page on Windows.
        /// </summary>
        public int? PptAnsiCodePageOverride { get; set; }

        /// <summary>
        /// Override the ANSI code page used for BIFF8 chart strings (when not Unicode).
        /// When null, uses the workbook CODEPAGE record (or 1252 as fallback).
        /// </summary>
        public int? BiffAnsiCodePageOverride { get; set; }

        /// <summary>
        /// Reports a progress update if a progress callback is configured.
        /// </summary>
        internal void ReportProgress(ConversionPhase phase, int percentComplete, string message, int slidesProcessed = 0, int totalSlides = 0)
        {
            Progress?.Invoke(new ConversionProgress(phase, percentComplete, message, slidesProcessed, totalSlides));
        }

        /// <summary>
        /// Logs a message if a log callback is configured.
        /// </summary>
        internal void LogMessage(string message)
        {
            Log?.Invoke(message);
        }
    }
}
