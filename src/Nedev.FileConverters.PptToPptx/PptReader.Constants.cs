namespace Nedev.FileConverters.PptToPptx
{
    public partial class PptReader
    {
        // PPT Record type constants - Document structure
        internal const ushort RT_Document = 1000;
        internal const ushort RT_DocumentAtom = 1001;
        internal const ushort RT_Slide = 1006;
        internal const ushort RT_SlideAtom = 1007;
        internal const ushort RT_SlideListWithText = 1008;
        internal const ushort RT_Notes = 1008;
        internal const ushort RT_NotesAtom = 1009;
        internal const ushort RT_Environment = 1010;
        internal const ushort RT_SlidePersistAtom = 1011;
        internal const ushort RT_SlideShowSlideInfoAtom = 1012;
        internal const ushort RT_MainMaster = 1016;
        internal const ushort RT_SlideMasterAtom = 1017;
        internal const ushort RT_ColorSchemeAtom = 2032;
        internal const ushort RT_FontCollection = 2005;
        internal const ushort RT_FontEntityAtom = 4023;

        // Text records
        internal const ushort RT_TextHeaderAtom = 3999;
        internal const ushort RT_TextCharsAtom = 4000;
        internal const ushort RT_StyleTextPropAtom = 4001;
        internal const ushort RT_TextBytesAtom = 4008;

        // User edit and persist
        internal const ushort RT_UserEditAtom = 4085;
        internal const ushort RT_CurrentUserAtom = 4086;
        internal const ushort RT_PersistDirectoryAtom = 6002;

        // Hyperlink records
        internal const ushort RT_ExObjList = 1033;
        internal const ushort RT_ExObjListAtom = 1034;
        internal const ushort RT_ExHyperlink = 4055;
        internal const ushort RT_ExHyperlinkAtom = 4051;
        internal const ushort RT_InteractiveInfo = 4082;
        internal const ushort RT_InteractiveInfoAtom = 4083;
        internal const ushort RT_TextInteractiveInfoAtom = 4084;
        internal const ushort RT_CString = 4056;

        // Animation records
        internal const ushort RT_AnimationInfoContainer = 4072;
        internal const ushort RT_AnimationInfoAtom = 4073;

        // OLE / ExObj records
        internal const ushort RT_ExObjRefAtom = 3009;
        internal const ushort RT_ExOleObjStg = 4113;
        internal const ushort RT_ExOleObjAtom = 4035;
        internal const ushort RT_ExEmbed = 4044;
        internal const ushort RT_ExOleEmbed = 4034;
        internal const ushort RT_ExOleLink = 4036;

        // Programmable Tags
        internal const ushort RT_ProgTags = 5000;
        internal const ushort RT_ProgStringTag = 5001;
        internal const ushort RT_ProgBinaryTag = 5002;
        internal const ushort RT_BinaryTagData = 5003;

        // Escher record types
        internal const ushort ESCHER_DggContainer = 0xF000;
        internal const ushort ESCHER_BStoreContainer = 0xF001;
        internal const ushort ESCHER_DgContainer = 0xF002;
        internal const ushort ESCHER_SpgrContainer = 0xF003;
        internal const ushort ESCHER_SpContainer = 0xF004;
        internal const ushort ESCHER_Sp = 0xF00A;
        internal const ushort ESCHER_Opt = 0xF00B;
        internal const ushort ESCHER_ClientTextbox = 0xF00D;
        internal const ushort ESCHER_ChildAnchor = 0xF00F;
        internal const ushort ESCHER_ClientAnchor = 0xF010;
        internal const ushort ESCHER_ClientData = 0xF011;
        internal const ushort ESCHER_BlipFirst = 0xF018;
        internal const ushort ESCHER_BlipLast = 0xF117;
    }
}
