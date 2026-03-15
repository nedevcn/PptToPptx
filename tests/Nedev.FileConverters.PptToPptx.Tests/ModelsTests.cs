using System.Collections.Generic;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class ModelsTests
    {
        [Fact]
        public void Presentation_DefaultValues_AreCorrect()
        {
            var presentation = new Presentation();

            Assert.NotNull(presentation.Slides);
            Assert.NotNull(presentation.Masters);
            Assert.NotNull(presentation.Images);
            Assert.NotNull(presentation.EmbeddedResources);
            Assert.Null(presentation.VbaProject);
            Assert.Equal(9144000, presentation.SlideWidth);
            Assert.Equal(6858000, presentation.SlideHeight);
        }

        [Fact]
        public void Slide_DefaultValues_AreCorrect()
        {
            var slide = new Slide();

            Assert.NotNull(slide.Shapes);
            Assert.NotNull(slide.TextContent);
            Assert.Null(slide.Notes);
            Assert.Null(slide.Transition);
            Assert.Null(slide.ColorScheme);
        }

        [Fact]
        public void Shape_DefaultValues_AreCorrect()
        {
            var shape = new Shape();

            Assert.Equal("Rectangle", shape.Type);
            Assert.NotNull(shape.Paragraphs);
            Assert.Equal("t", shape.VerticalAlignment);
            Assert.Null(shape.Table);
            Assert.False(shape.IsNativeTable);
            Assert.Null(shape.ImageId);
            Assert.Null(shape.FillColor);
            Assert.Null(shape.LineColor);
            Assert.False(shape.HasGradientFill);
            Assert.False(shape.HasShadow);
            Assert.Null(shape.Hyperlink);
        }

        [Fact]
        public void TextParagraph_GetPlainText_ReturnsConcatenatedText()
        {
            var paragraph = new TextParagraph();
            paragraph.Runs.Add(new TextRun { Text = "Hello " });
            paragraph.Runs.Add(new TextRun { Text = "World" });

            var result = paragraph.GetPlainText();

            Assert.Equal("Hello World", result);
        }

        [Fact]
        public void TextParagraph_DefaultValues_AreCorrect()
        {
            var paragraph = new TextParagraph();

            Assert.NotNull(paragraph.Runs);
            Assert.Equal(TextAlignment.Left, paragraph.Alignment);
            Assert.Equal(0, paragraph.IndentLevel);
            Assert.False(paragraph.HasBullet);
            Assert.Null(paragraph.BulletChar);
        }

        [Fact]
        public void TextRun_DefaultValues_AreCorrect()
        {
            var run = new TextRun();

            Assert.Equal("", run.Text);
            Assert.Equal(1800, run.FontSize);
            Assert.False(run.Bold);
            Assert.False(run.Italic);
            Assert.False(run.Underline);
            Assert.Null(run.Color);
            Assert.Null(run.Hyperlink);
        }

        [Fact]
        public void Chart_DefaultValues_AreCorrect()
        {
            var chart = new Chart();

            Assert.Equal("bar", chart.Type);
            Assert.Null(chart.Title);
            Assert.Null(chart.CategoryAxisTitle);
            Assert.Null(chart.ValueAxisTitle);
            Assert.True(chart.ShowLegend);
            Assert.Equal("r", chart.LegendPosition);
            Assert.NotNull(chart.Series);
        }

        [Fact]
        public void ChartSeries_DefaultValues_AreCorrect()
        {
            var series = new ChartSeries();

            Assert.Null(series.Name);
            Assert.NotNull(series.Categories);
            Assert.NotNull(series.Values);
            Assert.Null(series.Color);
            Assert.Null(series.MarkerType);
        }

        [Fact]
        public void Table_DefaultValues_AreCorrect()
        {
            var table = new Table();

            Assert.NotNull(table.Rows);
            Assert.NotNull(table.ColumnWidths);
        }

        [Fact]
        public void TableCell_DefaultValues_AreCorrect()
        {
            var cell = new TableCell();

            Assert.NotNull(cell.TextContent);
            Assert.Equal("t", cell.VerticalAlignment);
            Assert.Equal(91440, cell.MarginLeft);
            Assert.Equal(45720, cell.MarginTop);
            Assert.Equal(91440, cell.MarginRight);
            Assert.Equal(45720, cell.MarginBottom);
        }

        [Fact]
        public void SlideTransition_DefaultValues_AreCorrect()
        {
            var transition = new SlideTransition();

            Assert.Equal("none", transition.Type);
            Assert.Equal("fast", transition.Speed);
            Assert.Equal(0, transition.AdvanceTime);
            Assert.False(transition.HasAutoAdvance);
        }

        [Fact]
        public void ColorScheme_DefaultValues_AreCorrect()
        {
            var scheme = new ColorScheme();

            Assert.Equal("FFFFFF", scheme.Background);
            Assert.Equal("000000", scheme.TextAndLines);
            Assert.Equal("808080", scheme.Shadows);
            Assert.Equal("000000", scheme.TitleText);
            Assert.Equal("00FF00", scheme.Fills);
            Assert.Equal("0000FF", scheme.Accent);
            Assert.Equal("0000FF", scheme.AccentAndHyperlink);
            Assert.Equal("800080", scheme.AccentAndFollowingHyperlink);
        }

        [Fact]
        public void ImageResource_DefaultValues_AreCorrect()
        {
            var resource = new ImageResource();

            Assert.Equal(0, resource.Id);
            Assert.NotNull(resource.Data);
            Assert.Empty(resource.Data);
            Assert.Null(resource.Extension);
            Assert.Null(resource.ContentType);
        }

        [Fact]
        public void EmbeddedResource_DefaultValues_AreCorrect()
        {
            var resource = new EmbeddedResource();

            Assert.Equal(0, resource.Id);
            Assert.Equal("ole", resource.Kind);
            Assert.Null(resource.ProgId);
            Assert.Null(resource.FileName);
            Assert.Null(resource.Extension);
            Assert.Null(resource.ContentType);
            Assert.NotNull(resource.Data);
            Assert.Empty(resource.Data);
        }
    }
}
