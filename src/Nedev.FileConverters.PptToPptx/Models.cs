namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// Represents a PowerPoint presentation containing slides, masters, and resources.
    /// </summary>
    public class Presentation
    {
        /// <summary>
        /// Gets or sets the collection of slides in the presentation.
        /// </summary>
        public List<Slide> Slides { get; set; }

        /// <summary>
        /// Gets or sets the collection of master slides.
        /// </summary>
        public List<Slide> Masters { get; set; }

        /// <summary>
        /// Gets or sets the VBA project if the presentation contains macros.
        /// </summary>
        public VbaProject? VbaProject { get; set; }

        /// <summary>
        /// Gets or sets the slide width in EMU (English Metric Units). Default is 9144000 (10 inches).
        /// </summary>
        public int SlideWidth { get; set; } = 9144000;

        /// <summary>
        /// Gets or sets the slide height in EMU (English Metric Units). Default is 6858000 (7.5 inches).
        /// </summary>
        public int SlideHeight { get; set; } = 6858000;

        /// <summary>
        /// Gets or sets the collection of images used in the presentation.
        /// </summary>
        public List<ImageResource> Images { get; set; } = new List<ImageResource>();

        /// <summary>
        /// Gets or sets the collection of embedded resources (OLE objects, media).
        /// </summary>
        public List<EmbeddedResource> EmbeddedResources { get; set; } = new List<EmbeddedResource>();

        /// <summary>
        /// Initializes a new instance of the <see cref="Presentation"/> class.
        /// </summary>
        public Presentation()
        {
            Slides = new List<Slide>();
            Masters = new List<Slide>();
        }
    }

    /// <summary>
    /// Represents a single slide in a presentation.
    /// </summary>
    public class Slide
    {
        /// <summary>
        /// Gets or sets the slide index (1-based).
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or sets the unique slide identifier.
        /// </summary>
        public int SlideId { get; set; }

        /// <summary>
        /// Gets or sets the collection of shapes on this slide.
        /// </summary>
        public List<Shape> Shapes { get; set; }

        /// <summary>
        /// Gets or sets the text content of the slide.
        /// </summary>
        public List<TextParagraph> TextContent { get; set; }

        /// <summary>
        /// Gets or sets the notes text for this slide.
        /// </summary>
        public string? Notes { get; set; }

        /// <summary>
        /// Gets or sets the transition effect for this slide.
        /// </summary>
        public SlideTransition? Transition { get; set; }

        /// <summary>
        /// Gets or sets the color scheme applied to this slide.
        /// </summary>
        public ColorScheme? ColorScheme { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Slide"/> class.
        /// </summary>
        public Slide()
        {
            Shapes = new List<Shape>();
            TextContent = new List<TextParagraph>();
        }
    }

    /// <summary>
    /// Represents a shape on a slide (text box, image, chart, etc.).
    /// </summary>
    public class Shape
    {
        /// <summary>
        /// Gets or sets the shape type. Common values: "TextBox", "Rectangle", "Picture", "Chart", "Group", "Placeholder".
        /// </summary>
        public string Type { get; set; } = "Rectangle";

        /// <summary>
        /// Gets or sets the plain text content of the shape.
        /// </summary>
        public string? Text { get; set; }

        /// <summary>
        /// Gets or sets the chart data if this shape contains a chart.
        /// </summary>
        public Chart? Chart { get; set; }

        /// <summary>
        /// Gets or sets the placeholder type. Common values: "title", "body", "dt" (date), "ftr" (footer), "sldNum" (slide number).
        /// </summary>
        public string? PlaceholderType { get; set; }

        /// <summary>
        /// Gets or sets the placeholder index.
        /// </summary>
        public int? PlaceholderIndex { get; set; }

        /// <summary>
        /// Gets or sets the embedded resource ID if this shape references an embedded object.
        /// </summary>
        public int? EmbeddedResourceId { get; set; }

        /// <summary>
        /// Gets or sets the click action URL (e.g., "ppaction://hlinkshowjump?jump=nextslide").
        /// </summary>
        public string? ClickAction { get; set; }

        /// <summary>
        /// Gets or sets the left position in EMU.
        /// </summary>
        public long Left { get; set; }

        /// <summary>
        /// Gets or sets the top position in EMU.
        /// </summary>
        public long Top { get; set; }

        /// <summary>
        /// Gets or sets the width in EMU.
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Gets or sets the height in EMU.
        /// </summary>
        public long Height { get; set; }

        /// <summary>
        /// Gets or sets the rich text paragraphs.
        /// </summary>
        public List<TextParagraph> Paragraphs { get; set; }

        /// <summary>
        /// Gets or sets the table data if this shape contains a table.
        /// </summary>
        public Table? Table { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this is a native PowerPoint table.
        /// </summary>
        public bool IsNativeTable { get; set; }

        /// <summary>
        /// Gets or sets the image ID if this shape references an image.
        /// </summary>
        public int? ImageId { get; set; }

        /// <summary>
        /// Gets or sets the raw image data (temporary storage during parsing).
        /// </summary>
        public byte[]? ImageData { get; set; }

        /// <summary>
        /// Gets or sets the image content type (e.g., "image/png").
        /// </summary>
        public string? ImageContentType { get; set; }

        /// <summary>
        /// Gets or sets the fill color in RRGGBB format.
        /// </summary>
        public string? FillColor { get; set; }

        /// <summary>
        /// Gets or sets the line color in RRGGBB format.
        /// </summary>
        public string? LineColor { get; set; }

        /// <summary>
        /// Gets or sets the line width in EMU.
        /// </summary>
        public long? LineWidth { get; set; }

        /// <summary>
        /// Gets or sets the line dash style ("solid", "dash", "dot", "dashDot", "lgDash").
        /// </summary>
        public string? LineDash { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this shape has a gradient fill.
        /// </summary>
        public bool HasGradientFill { get; set; }

        /// <summary>
        /// Gets or sets the background color for gradient fills in RRGGBB format.
        /// </summary>
        public string? FillBackColor { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this shape has a shadow.
        /// </summary>
        public bool HasShadow { get; set; }

        /// <summary>
        /// Gets or sets the shadow color in RRGGBB format.
        /// </summary>
        public string? ShadowColor { get; set; }

        /// <summary>
        /// Gets or sets the hyperlink URL.
        /// </summary>
        public string? Hyperlink { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment ("t" = top, "ctr" = center, "b" = bottom).
        /// </summary>
        public string VerticalAlignment { get; set; } = "t";

        /// <summary>
        /// Gets or sets the left margin in EMU.
        /// </summary>
        public long? MarginLeft { get; set; }

        /// <summary>
        /// Gets or sets the top margin in EMU.
        /// </summary>
        public long? MarginTop { get; set; }

        /// <summary>
        /// Gets or sets the right margin in EMU.
        /// </summary>
        public long? MarginRight { get; set; }

        /// <summary>
        /// Gets or sets the bottom margin in EMU.
        /// </summary>
        public long? MarginBottom { get; set; }

        /// <summary>
        /// Gets or sets the shape geometry information.
        /// </summary>
        public ShapeGeometry? Geometry { get; set; }

        /// <summary>
        /// Gets or sets the animation information for this shape.
        /// </summary>
        public ShapeAnimation? Animation { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Shape"/> class.
        /// </summary>
        public Shape()
        {
            Paragraphs = new List<TextParagraph>();
        }
    }

    /// <summary>
    /// Represents an image resource embedded in the presentation.
    /// </summary>
    public class ImageResource
    {
        /// <summary>
        /// Gets or sets the unique identifier for this image.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Gets or sets the raw image data.
        /// </summary>
        public byte[] Data { get; set; } = Array.Empty<byte>();

        /// <summary>
        /// Gets or sets the file extension (e.g., "png", "jpg").
        /// </summary>
        public string? Extension { get; set; }

        /// <summary>
        /// Gets or sets the MIME content type (e.g., "image/png").
        /// </summary>
        public string? ContentType { get; set; }
    }

    /// <summary>
    /// Represents a paragraph of text with formatting.
    /// </summary>
    public class TextParagraph
    {
        /// <summary>
        /// Gets or sets the collection of text runs in this paragraph.
        /// </summary>
        public List<TextRun> Runs { get; set; }

        /// <summary>
        /// Gets or sets the text alignment.
        /// </summary>
        public TextAlignment Alignment { get; set; } = TextAlignment.Left;

        /// <summary>
        /// Gets or sets the indentation level (0-based).
        /// </summary>
        public int IndentLevel { get; set; } = 0;

        /// <summary>
        /// Gets or sets a value indicating whether this paragraph has a bullet.
        /// </summary>
        public bool HasBullet { get; set; }

        /// <summary>
        /// Gets or sets the bullet character.
        /// </summary>
        public char? BulletChar { get; set; }

        /// <summary>
        /// Gets or sets the bullet font name.
        /// </summary>
        public string BulletFont { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the line spacing in master units.
        /// </summary>
        public short? LineSpacing { get; set; }

        /// <summary>
        /// Gets or sets the space before the paragraph in master units.
        /// </summary>
        public short? SpaceBefore { get; set; }

        /// <summary>
        /// Gets or sets the space after the paragraph in master units.
        /// </summary>
        public short? SpaceAfter { get; set; }

        /// <summary>
        /// Gets or sets the left margin in master units.
        /// </summary>
        public short? LeftMargin { get; set; }

        /// <summary>
        /// Gets or sets the indent in master units.
        /// </summary>
        public short? Indent { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="TextParagraph"/> class.
        /// </summary>
        public TextParagraph()
        {
            Runs = new List<TextRun>();
        }

        /// <summary>
        /// Gets the plain text content of this paragraph by concatenating all runs.
        /// </summary>
        /// <returns>The plain text content.</returns>
        public string GetPlainText()
        {
            return string.Join("", Runs.ConvertAll(r => r.Text));
        }
    }

    /// <summary>
    /// Represents a run of text with consistent formatting.
    /// </summary>
    public class TextRun
    {
        /// <summary>
        /// Gets or sets the text content.
        /// </summary>
        public string Text { get; set; } = "";

        /// <summary>
        /// Gets or sets the font name.
        /// </summary>
        public string? FontName { get; set; }

        /// <summary>
        /// Gets or sets the font size in hundredths of a point (default 1800 = 18pt).
        /// </summary>
        public int FontSize { get; set; } = 1800;

        /// <summary>
        /// Gets or sets a value indicating whether the text is bold.
        /// </summary>
        public bool Bold { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the text is italic.
        /// </summary>
        public bool Italic { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the text is underlined.
        /// </summary>
        public bool Underline { get; set; }

        /// <summary>
        /// Gets or sets the text color in RRGGBB format.
        /// </summary>
        public string? Color { get; set; }

        /// <summary>
        /// Gets or sets the hyperlink URL.
        /// </summary>
        public string? Hyperlink { get; set; }

        /// <summary>
        /// Gets or sets the click action URL.
        /// </summary>
        public string? ClickAction { get; set; }
    }

    /// <summary>
    /// Represents an embedded resource (OLE object or media file).
    /// </summary>
    public class EmbeddedResource
    {
        /// <summary>
        /// Gets or sets the unique identifier.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Gets or sets the resource kind ("ole", "media", "unknown").
        /// </summary>
        public string Kind { get; set; } = "ole";

        /// <summary>
        /// Gets or sets the program ID for OLE objects.
        /// </summary>
        public string? ProgId { get; set; }

        /// <summary>
        /// Gets or sets the file name.
        /// </summary>
        public string? FileName { get; set; }

        /// <summary>
        /// Gets or sets the file extension.
        /// </summary>
        public string? Extension { get; set; }

        /// <summary>
        /// Gets or sets the MIME content type.
        /// </summary>
        public string? ContentType { get; set; }

        /// <summary>
        /// Gets or sets the raw data.
        /// </summary>
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }

    /// <summary>
    /// Specifies the text alignment options.
    /// </summary>
    public enum TextAlignment
    {
        /// <summary>Left alignment.</summary>
        Left,
        /// <summary>Center alignment.</summary>
        Center,
        /// <summary>Right alignment.</summary>
        Right,
        /// <summary>Justified alignment.</summary>
        Justify
    }

    /// <summary>
    /// Represents a chart embedded in a shape.
    /// </summary>
    public class Chart
    {
        /// <summary>
        /// Gets or sets the chart type ("bar", "line", "pie", "area", "scatter", "radar").
        /// </summary>
        public string Type { get; set; } = "bar";

        /// <summary>
        /// Gets or sets the chart title.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the category axis title.
        /// </summary>
        public string? CategoryAxisTitle { get; set; }

        /// <summary>
        /// Gets or sets the value axis title.
        /// </summary>
        public string? ValueAxisTitle { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to show the legend.
        /// </summary>
        public bool ShowLegend { get; set; } = true;

        /// <summary>
        /// Gets or sets the legend position ("r", "l", "t", "b", "tr").
        /// </summary>
        public string LegendPosition { get; set; } = "r";

        /// <summary>
        /// Gets or sets the collection of data series.
        /// </summary>
        public List<ChartSeries> Series { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Chart"/> class.
        /// </summary>
        public Chart()
        {
            Series = new List<ChartSeries>();
        }
    }

    /// <summary>
    /// Represents a data series in a chart.
    /// </summary>
    public class ChartSeries
    {
        /// <summary>
        /// Gets or sets the series name.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the category labels.
        /// </summary>
        public List<string> Categories { get; set; }

        /// <summary>
        /// Gets or sets the data values.
        /// </summary>
        public List<double> Values { get; set; }

        /// <summary>
        /// Gets or sets the series color in RRGGBB format.
        /// </summary>
        public string? Color { get; set; }

        /// <summary>
        /// Gets or sets the marker type for line/scatter charts.
        /// </summary>
        public string? MarkerType { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSeries"/> class.
        /// </summary>
        public ChartSeries()
        {
            Categories = new List<string>();
            Values = new List<double>();
        }
    }

    /// <summary>
    /// Represents a VBA project containing macros.
    /// </summary>
    public class VbaProject
    {
        /// <summary>
        /// Gets or sets the project name.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the raw project data.
        /// </summary>
        public byte[] ProjectData { get; set; } = Array.Empty<byte>();

        /// <summary>
        /// Gets or sets the collection of VBA modules.
        /// </summary>
        public List<VbaModule> Modules { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaProject"/> class.
        /// </summary>
        public VbaProject()
        {
            Modules = new List<VbaModule>();
        }
    }

    /// <summary>
    /// Represents a VBA code module.
    /// </summary>
    public class VbaModule
    {
        /// <summary>
        /// Gets or sets the module name.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the VBA code content.
        /// </summary>
        public string? Code { get; set; }
    }

    /// <summary>
    /// Represents a table structure.
    /// </summary>
    public class Table
    {
        /// <summary>
        /// Gets or sets the collection of rows.
        /// </summary>
        public List<TableRow> Rows { get; set; }

        /// <summary>
        /// Gets or sets the column widths in EMU.
        /// </summary>
        public List<long> ColumnWidths { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Table"/> class.
        /// </summary>
        public Table()
        {
            Rows = new List<TableRow>();
            ColumnWidths = new List<long>();
        }
    }

    /// <summary>
    /// Represents a row in a table.
    /// </summary>
    public class TableRow
    {
        /// <summary>
        /// Gets or sets the collection of cells.
        /// </summary>
        public List<TableCell> Cells { get; set; }

        /// <summary>
        /// Gets or sets the row height in EMU.
        /// </summary>
        public long Height { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="TableRow"/> class.
        /// </summary>
        public TableRow()
        {
            Cells = new List<TableCell>();
        }
    }

    /// <summary>
    /// Represents a cell in a table.
    /// </summary>
    public class TableCell
    {
        /// <summary>
        /// Gets or sets the text content.
        /// </summary>
        public List<TextParagraph> TextContent { get; set; }

        /// <summary>
        /// Gets or sets the fill color in RRGGBB format.
        /// </summary>
        public string? FillColor { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment ("t", "ctr", "b").
        /// </summary>
        public string VerticalAlignment { get; set; } = "t";

        /// <summary>
        /// Gets or sets the left margin in EMU.
        /// </summary>
        public long MarginLeft { get; set; } = 91440;

        /// <summary>
        /// Gets or sets the top margin in EMU.
        /// </summary>
        public long MarginTop { get; set; } = 45720;

        /// <summary>
        /// Gets or sets the right margin in EMU.
        /// </summary>
        public long MarginRight { get; set; } = 91440;

        /// <summary>
        /// Gets or sets the bottom margin in EMU.
        /// </summary>
        public long MarginBottom { get; set; } = 45720;

        /// <summary>
        /// Initializes a new instance of the <see cref="TableCell"/> class.
        /// </summary>
        public TableCell()
        {
            TextContent = new List<TextParagraph>();
        }
    }

    /// <summary>
    /// Represents a slide transition effect.
    /// </summary>
    public class SlideTransition
    {
        /// <summary>
        /// Gets or sets the transition type ("fade", "push", "wipe", etc.).
        /// </summary>
        public string Type { get; set; } = "none";

        /// <summary>
        /// Gets or sets the transition speed ("slow", "med", "fast").
        /// </summary>
        public string Speed { get; set; } = "fast";

        /// <summary>
        /// Gets or sets the auto-advance time in milliseconds.
        /// </summary>
        public int AdvanceTime { get; set; } = 0;

        /// <summary>
        /// Gets or sets a value indicating whether the slide auto-advances.
        /// </summary>
        public bool HasAutoAdvance { get; set; }
    }

    /// <summary>
    /// Represents a color scheme for the presentation.
    /// </summary>
    public class ColorScheme
    {
        /// <summary>Gets or sets the background color.</summary>
        public string Background { get; set; } = "FFFFFF";

        /// <summary>Gets or sets the text and lines color.</summary>
        public string TextAndLines { get; set; } = "000000";

        /// <summary>Gets or sets the shadows color.</summary>
        public string Shadows { get; set; } = "808080";

        /// <summary>Gets or sets the title text color.</summary>
        public string TitleText { get; set; } = "000000";

        /// <summary>Gets or sets the fills color.</summary>
        public string Fills { get; set; } = "00FF00";

        /// <summary>Gets or sets the accent color.</summary>
        public string Accent { get; set; } = "0000FF";

        /// <summary>Gets or sets the accent and hyperlink color.</summary>
        public string AccentAndHyperlink { get; set; } = "0000FF";

        /// <summary>Gets or sets the accent and following hyperlink color.</summary>
        public string AccentAndFollowingHyperlink { get; set; } = "800080";
    }

    /// <summary>
    /// Represents animation information for a shape.
    /// </summary>
    public class ShapeAnimation
    {
        /// <summary>
        /// Gets or sets the animation type ("fly", "wipe", "fade", etc.).
        /// </summary>
        public string Type { get; set; } = "none";

        /// <summary>
        /// Gets or sets the animation direction.
        /// </summary>
        public string Direction { get; set; } = "none";

        /// <summary>
        /// Gets or sets a value indicating whether the animation triggers on click.
        /// </summary>
        public bool TriggerOnClick { get; set; } = true;

        /// <summary>
        /// Gets or sets the animation order.
        /// </summary>
        public int Order { get; set; } = 0;

        /// <summary>
        /// Gets or sets the animation duration in milliseconds.
        /// </summary>
        public int DurationMs { get; set; } = 500;
    }

    /// <summary>
    /// Represents shape geometry information.
    /// </summary>
    public class ShapeGeometry
    {
        /// <summary>Gets or sets the geometry left coordinate.</summary>
        public long GeoLeft { get; set; } = 0;

        /// <summary>Gets or sets the geometry top coordinate.</summary>
        public long GeoTop { get; set; } = 0;

        /// <summary>Gets or sets the geometry right coordinate.</summary>
        public long GeoRight { get; set; } = 21600;

        /// <summary>Gets or sets the geometry bottom coordinate.</summary>
        public long GeoBottom { get; set; } = 21600;

        /// <summary>Gets or sets the collection of geometry paths.</summary>
        public List<GeometryPath> Paths { get; set; } = new List<GeometryPath>();
    }

    /// <summary>
    /// Represents a geometry path.
    /// </summary>
    public class GeometryPath
    {
        /// <summary>Gets or sets the path width.</summary>
        public long Width { get; set; } = 21600;

        /// <summary>Gets or sets the path height.</summary>
        public long Height { get; set; } = 21600;

        /// <summary>Gets or sets the collection of geometry commands.</summary>
        public List<GeometryCommand> Commands { get; set; } = new List<GeometryCommand>();
    }

    /// <summary>
    /// Represents a geometry command (move, line, curve, etc.).
    /// </summary>
    public class GeometryCommand
    {
        /// <summary>
        /// Gets or sets the command type ("moveTo", "lnTo", "arcTo", "cubicBezTo", "close").
        /// </summary>
        public string Type { get; set; } = "";

        /// <summary>
        /// Gets or sets the collection of points for this command.
        /// </summary>
        public List<GeometryPoint> Points { get; set; } = new List<GeometryPoint>();
    }

    /// <summary>
    /// Represents a point in geometry coordinates.
    /// </summary>
    public struct GeometryPoint
    {
        /// <summary>Gets or sets the X coordinate.</summary>
        public long X { get; set; }

        /// <summary>Gets or sets the Y coordinate.</summary>
        public long Y { get; set; }
    }
}
