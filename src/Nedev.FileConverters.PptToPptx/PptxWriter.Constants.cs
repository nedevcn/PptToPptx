namespace Nedev.FileConverters.PptToPptx
{
    public partial class PptxWriter
    {
        // OpenXML Namespaces
        internal const string NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        internal const string NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal const string NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal const string NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
        internal const string NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal const string NS_DC = "http://purl.org/dc/elements/1.1/";
        internal const string NS_DCTERMS = "http://purl.org/dc/terms/";
        internal const string NS_DCMITYPE = "http://purl.org/dc/dcmitype/";
        internal const string NS_XSI = "http://www.w3.org/2001/XMLSchema-instance";
        internal const string NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        internal const string NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        internal const string NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        // Relationship Types
        internal const string REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        internal const string REL_CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        internal const string REL_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        internal const string REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        internal const string REL_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
        internal const string REL_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        internal const string REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        internal const string REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        internal const string REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        internal const string REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        internal const string REL_VBA_PROJECT = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        internal const string REL_OLE_OBJECT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

        // Content Types
        internal const string CT_PRESENTATION = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
        internal const string CT_PRESENTATION_MACRO = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
        internal const string CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        internal const string CT_SLIDE_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        internal const string CT_SLIDE_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        internal const string CT_NOTES_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
        internal const string CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
        internal const string CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        internal const string CT_CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml";
        internal const string CT_EXT_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        internal const string CT_VBA_PROJECT = "application/vnd.ms-office.vbaProject";
        internal const string CT_OLE_OBJECT = "application/vnd.openxmlformats-officedocument.oleObject";
    }
}
