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
            using var pptReader = new PptReader(pptPath, options);
            using var pptxWriter = new PptxWriter(pptxPath, options);
            
            pptxWriter.WritePresentation(pptReader.ReadPresentation());
        }
    }
}
