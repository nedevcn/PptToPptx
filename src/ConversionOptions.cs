using System;

namespace Nefdev.PptToPptx
{
    public sealed class ConversionOptions
    {
        /// <summary>
        /// Optional log sink. When null, the library stays silent.
        /// </summary>
        public Action<string>? Log { get; set; }

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
    }
}

