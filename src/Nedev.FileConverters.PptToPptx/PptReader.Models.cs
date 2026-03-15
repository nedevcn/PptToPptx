namespace Nedev.FileConverters.PptToPptx
{
    public partial class PptReader
    {
        /// <summary>
        /// OLE object information structure for tracking embedded objects.
        /// </summary>
        private struct OleObjectInfo
        {
            public string StorageName;
            public string ProgId;
        }

        /// <summary>
        /// Record header structure for PPT binary records.
        /// </summary>
        private struct RecordHeader
        {
            public int RecVer;
            public int RecInstance;
            public ushort RecType;
            public uint RecLen;

            public bool IsContainer => RecVer == 0x0F;
        }

        /// <summary>
        /// Text range information for hyperlink mapping.
        /// </summary>
        private readonly struct TextRange
        {
            public int Start { get; }
            public int End { get; }
            public int HyperlinkId { get; }

            public TextRange(int start, int end, int hyperlinkId)
            {
                Start = start;
                End = end;
                HyperlinkId = hyperlinkId;
            }
        }
    }
}
