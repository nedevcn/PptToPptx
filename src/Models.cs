namespace Nefdev.PptToPptx
{
    public class Presentation
    {
        public List<Slide> Slides { get; set; }
        public VbaProject VbaProject { get; set; }
        public int SlideWidth { get; set; } = 9144000;   // EMU (10 inches)
        public int SlideHeight { get; set; } = 6858000;  // EMU (7.5 inches)
        public List<ImageResource> Images { get; set; } = new List<ImageResource>();
        
        public Presentation()
        {
            Slides = new List<Slide>();
        }
    }
    
    public class Slide
    {
        public int Index { get; set; }
        public List<Shape> Shapes { get; set; }
        public List<TextParagraph> TextContent { get; set; }
        
        public Slide()
        {
            Shapes = new List<Shape>();
            TextContent = new List<TextParagraph>();
        }
    }
    
    public class Shape
    {
        public string Type { get; set; }  // "TextBox", "Rectangle", "Picture", "Chart", "Group", "Placeholder"
        public string Text { get; set; }
        public Chart Chart { get; set; }
        
        // 位置和大小 (EMU)
        public long Left { get; set; }
        public long Top { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        
        // 文本内容 (富文本)
        public List<TextParagraph> Paragraphs { get; set; }
        
        // 图片引用
        public int? ImageId { get; set; }
        public byte[] ImageData { get; set; }
        public string ImageContentType { get; set; }
        
        // 填充颜色 (RRGGBB 格式)
        public string FillColor { get; set; }
        
        // 线条颜色
        public string LineColor { get; set; }
        
        // 超链接
        public string Hyperlink { get; set; }
        
        public Shape()
        {
            Paragraphs = new List<TextParagraph>();
        }
    }
    
    public class ImageResource
    {
        public int Id { get; set; }
        public byte[] Data { get; set; }
        public string Extension { get; set; }
        public string ContentType { get; set; }
    }
    
    public class TextParagraph
    {
        public List<TextRun> Runs { get; set; }
        public TextAlignment Alignment { get; set; } = TextAlignment.Left;
        
        // Advanced formatting
        public int IndentLevel { get; set; } = 0;
        public bool HasBullet { get; set; }
        public char? BulletChar { get; set; }
        public string BulletFont { get; set; }
        
        public TextParagraph()
        {
            Runs = new List<TextRun>();
        }
        
        public string GetPlainText()
        {
            return string.Join("", Runs.ConvertAll(r => r.Text));
        }
    }
    
    public class TextRun
    {
        public string Text { get; set; } = "";
        public string FontName { get; set; }
        public int FontSize { get; set; } = 1800; // hundredths of a point (18pt default)
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string Color { get; set; }  // RRGGBB
        
        // 超链接
        public string Hyperlink { get; set; }
    }
    
    public enum TextAlignment
    {
        Left,
        Center,
        Right,
        Justify
    }
    
    public class Chart
    {
        public string Type { get; set; }
        public string Title { get; set; }
        public string CategoryAxisTitle { get; set; }
        public string ValueAxisTitle { get; set; }
        public bool ShowLegend { get; set; } = true;
        public string LegendPosition { get; set; } = "r"; // r, l, t, b, tr
        public List<ChartSeries> Series { get; set; }
        
        public Chart()
        {
            Series = new List<ChartSeries>();
        }
    }
    
    public class ChartSeries
    {
        public string Name { get; set; }
        public List<string> Categories { get; set; }
        public List<double> Values { get; set; }
        
        public ChartSeries()
        {
            Categories = new List<string>();
            Values = new List<double>();
        }
    }
    
    public class VbaProject
    {
        public string Name { get; set; }
        public byte[] ProjectData { get; set; }
        public List<VbaModule> Modules { get; set; }
        
        public VbaProject()
        {
            Modules = new List<VbaModule>();
        }
    }
    
    public class VbaModule
    {
        public string Name { get; set; }
        public string Code { get; set; }
    }
}
