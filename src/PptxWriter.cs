using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Collections.Generic;

namespace Nefdev.PptToPptx
{
    public class PptxWriter : IDisposable
    {
        private readonly string _outputPath;
        
        // 命名空间常量
        private const string NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string NS_DC = "http://purl.org/dc/elements/1.1/";
        private const string NS_DCTERMS = "http://purl.org/dc/terms/";
        private const string NS_DCMITYPE = "http://purl.org/dc/dcmitype/";
        private const string NS_XSI = "http://www.w3.org/2001/XMLSchema-instance";
        private const string NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private const string NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private const string NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        
        private const string REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string REL_CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string REL_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        private const string REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        private const string REL_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
        private const string REL_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        private const string REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        private const string REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        private const string REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        
        public PptxWriter(string path)
        {
            _outputPath = path;
        }
        
        public void WritePresentation(Presentation presentation)
        {
            var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);
            
            try
            {
                // 确保至少有一个幻灯片
                if (presentation.Slides.Count == 0)
                {
                    presentation.Slides.Add(new Slide { Index = 1 });
                }
                
                CreateDirectoryStructure(tempDir, presentation);
                WriteContentTypes(tempDir, presentation);
                WritePackageRelationships(tempDir);
                
                // 写入图片资源
                WriteMediaFiles(tempDir, presentation);
                
                WritePresentationXml(tempDir, presentation);
                WritePresentationRelationships(tempDir, presentation);
                WriteSlidesXml(tempDir, presentation);
                WriteSlideLayouts(tempDir, presentation);
                WriteSlideLayoutRelationships(tempDir, presentation);
                WriteSlideMasters(tempDir, presentation);
                WriteSlideMasterRelationships(tempDir, presentation);
                WriteTheme(tempDir, presentation);
                WriteCoreProperties(tempDir);
                WriteExtendedProperties(tempDir);
                
                if (presentation.VbaProject != null)
                {
                    WriteVbaProject(tempDir, presentation.VbaProject);
                }
                
                PackageAsPptx(tempDir, _outputPath);
                
                Console.WriteLine($"PPTX file written to: {_outputPath}");
            }
            finally
            {
                // 调试：复制临时目录
                string copyDir = Path.Combine(Path.GetDirectoryName(_outputPath), "temp_pptx");
                if (Directory.Exists(copyDir))
                    Directory.Delete(copyDir, recursive: true);
                Directory.CreateDirectory(copyDir);
                
                foreach (var file in Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories))
                {
                    string relativePath = Path.GetRelativePath(tempDir, file);
                    string destPath = Path.Combine(copyDir, relativePath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                    File.Copy(file, destPath);
                }
                
                Directory.Delete(tempDir, recursive: true);
            }
        }
        
        private void CreateDirectoryStructure(string baseDir, Presentation presentation)
        {
            Directory.CreateDirectory(Path.Combine(baseDir, "_rels"));
            Directory.CreateDirectory(Path.Combine(baseDir, "docProps"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "_rels"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "slides"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "slides", "_rels"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "slideLayouts"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "slideLayouts", "_rels"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "slideMasters"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "slideMasters", "_rels"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "theme"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "notesSlides"));
            Directory.CreateDirectory(Path.Combine(baseDir, "ppt", "notesSlides", "_rels"));
        }
        
        #region Content Types
        
        private void WriteContentTypes(string baseDir, Presentation presentation)
        {
            var path = Path.Combine(baseDir, "[Content_Types].xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Types", NS_CT);
            
            // Default
            DirectoryPropertyDefault(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            DirectoryPropertyDefault(writer, "xml", "application/xml");
            DirectoryPropertyDefault(writer, "png", "image/png");
            DirectoryPropertyDefault(writer, "jpeg", "image/jpeg");
            DirectoryPropertyDefault(writer, "jpg", "image/jpeg");
            DirectoryPropertyDefault(writer, "gif", "image/gif");
            DirectoryPropertyDefault(writer, "emf", "image/x-emf");
            DirectoryPropertyDefault(writer, "wmf", "image/x-wmf");
            DirectoryPropertyDefault(writer, "bmp", "image/bmp");
            DirectoryPropertyDefault(writer, "tiff", "image/tiff");
            
            // Override — presentation
            WriteOverride(writer, "/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");
            
            // Override — slides (动态)
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                WriteOverride(writer, $"/ppt/slides/slide{i + 1}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
                if (!string.IsNullOrEmpty(presentation.Slides[i].Notes))
                {
                    WriteOverride(writer, $"/ppt/notesSlides/notesSlide{i + 1}.xml", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml");
                }
            }
            
            // Override — slideLayout, slideMaster, theme
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                WriteOverride(writer, $"/ppt/slideLayouts/slideLayout{i + 1}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml");
                WriteOverride(writer, $"/ppt/slideMasters/slideMaster{i + 1}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml");
            }
            WriteOverride(writer, "/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
            
            // Override — docProps
            WriteOverride(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
            WriteOverride(writer, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            
            // Charts (动态)
            int chartCount = CountCharts(presentation);
            for (int i = 0; i < chartCount; i++)
            {
                WriteOverride(writer, $"/ppt/charts/chart{i + 1}.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
            }
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        private void DirectoryPropertyDefault(XmlWriter writer, string ext, string contentType)
        {
            writer.WriteStartElement("Default", NS_CT);
            writer.WriteAttributeString("Extension", ext);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }
        
        private void WriteDefault(XmlWriter writer, string ext, string contentType)
        {
            writer.WriteStartElement("Default", NS_CT);
            writer.WriteAttributeString("Extension", ext);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }
        
        private void WriteOverride(XmlWriter writer, string partName, string contentType)
        {
            writer.WriteStartElement("Override", NS_CT);
            writer.WriteAttributeString("PartName", partName);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }
        
        #endregion
        
        #region Package Relationships
        
        private void WritePackageRelationships(string baseDir)
        {
            var path = Path.Combine(baseDir, "_rels", ".rels");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", NS_RELS);
            
            WriteRelationship(writer, "rId1", REL_OFFICE_DOC, "ppt/presentation.xml");
            WriteRelationship(writer, "rId2", REL_CORE_PROPS, "docProps/core.xml");
            WriteRelationship(writer, "rId3", REL_EXT_PROPS, "docProps/app.xml");
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        #endregion
        
        #region Presentation
        
        private void WritePresentationXml(string baseDir, Presentation presentation)
        {
            var path = Path.Combine(baseDir, "ppt", "presentation.xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("p", "presentation", NS_P);
            writer.WriteAttributeString("xmlns", "a", null, NS_A);
            writer.WriteAttributeString("xmlns", "r", null, NS_R);
            
            // sldMasterIdLst
            writer.WriteStartElement("p", "sldMasterIdLst", NS_P);
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                writer.WriteStartElement("p", "sldMasterId", NS_P);
                writer.WriteAttributeString("id", $"{2147483648 + i}");
                writer.WriteAttributeString("r", "id", NS_R, $"rId{presentation.Slides.Count + i + 1}");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            
            // sldIdLst
            writer.WriteStartElement("p", "sldIdLst", NS_P);
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                writer.WriteStartElement("p", "sldId", NS_P);
                writer.WriteAttributeString("id", $"{256 + i}");
                writer.WriteAttributeString("r", "id", NS_R, $"rId{i + 1}");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            
            // sldSz
            writer.WriteStartElement("p", "sldSz", NS_P);
            writer.WriteAttributeString("cx", presentation.SlideWidth.ToString());
            writer.WriteAttributeString("cy", presentation.SlideHeight.ToString());
            writer.WriteEndElement();
            
            // notesSz
            writer.WriteStartElement("p", "notesSz", NS_P);
            writer.WriteAttributeString("cx", "6858000");
            writer.WriteAttributeString("cy", "9144000");
            writer.WriteEndElement();
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        private void WritePresentationRelationships(string baseDir, Presentation presentation)
        {
            var path = Path.Combine(baseDir, "ppt", "_rels", "presentation.xml.rels");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", NS_RELS);
            
            // Slides: rId1 .. rIdN
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                WriteRelationship(writer, $"rId{i + 1}", REL_SLIDE, $"slides/slide{i + 1}.xml");
            }
            
            // SlideMaster: rId(N+1)...rId(N+masterCount)
            int masterRidBase = presentation.Slides.Count + 1;
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                WriteRelationship(writer, $"rId{masterRidBase + i}", REL_SLIDE_MASTER, $"slideMasters/slideMaster{i + 1}.xml");
            }
            
            // Theme: rId(N+masterCount+1)
            int themeRid = presentation.Slides.Count + masterCount + 1;
            WriteRelationship(writer, $"rId{themeRid}", REL_THEME, "theme/theme1.xml");
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        #endregion
        
        #region Slides
        
        private void WriteSlidesXml(string baseDir, Presentation presentation)
        {
            var slides = presentation.Slides;
            for (int i = 0; i < slides.Count; i++)
            {
                int slideNum = i + 1;
                WriteSlideXml(baseDir, presentation, slides[i], slideNum);
                WriteSlideRelationship(baseDir, presentation, slides[i], slideNum);
                
                if (!string.IsNullOrEmpty(slides[i].Notes))
                {
                    WriteNotesSlideXml(baseDir, slides[i], slideNum);
                    WriteNotesSlideRelationships(baseDir, slideNum);
                }
            }
        }
        
        private void WriteSlideXml(string baseDir, Presentation presentation, Slide slide, int slideNum)
        {
            var path = Path.Combine(baseDir, "ppt", "slides", $"slide{slideNum}.xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("p", "sld", NS_P);
            writer.WriteAttributeString("xmlns", "a", null, NS_A);
            writer.WriteAttributeString("xmlns", "r", null, NS_R);
            
            writer.WriteStartElement("p", "cSld", NS_P);
            writer.WriteStartElement("p", "spTree", NS_P);
            
            // spTree 必须的 nvGrpSpPr 和 grpSpPr
            WriteGroupShapeProperties(writer);
            
            // 写入形状
            int shapeId = 2;  // 1 is reserved for the group shape
            int chartId = 1;
            int imageRid = 100; // Start high for images to avoid collision
            
            // 先写入 slide 形状
            foreach (var shape in slide.Shapes)
            {
                if (shape.Type == "Chart" && shape.Chart != null)
                {
                    WriteChartFrame(writer, shape, shapeId, chartId);
                    WriteChartXml(baseDir, shape.Chart, chartId);
                    chartId++;
                }
                else if (shape.Type == "Table" && shape.Table != null)
                {
                    WriteGraphicFrame(writer, shape, shapeId);
                }
                else if (shape.Type == "Picture" && shape.ImageId != null)
                {
                    if (presentation.Images.Any(img => img.Id == shape.ImageId.Value))
                    {
                        WritePictureShape(writer, shape, shapeId, imageRid, slideNum, slide.Shapes.IndexOf(shape));
                        imageRid++;
                    }
                }
                else
                {
                    WriteTextBoxShape(writer, shape, shapeId, slideNum, slide.Shapes.IndexOf(shape));
                }
                shapeId++;
            }
            
            // 如果没有形状但有文本内容，创建文本框
            if (slide.Shapes.Count == 0 && slide.TextContent.Count > 0)
            {
                long yPos = 457200; // 0.5 inch
                int fallbackShapeIndex = slide.Shapes.Count;
                foreach (var para in slide.TextContent)
                {
                    string text = para.GetPlainText();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var shape = new Shape
                        {
                            Type = "TextBox",
                            Text = text,
                            Left = 457200,
                            Top = yPos,
                            Width = 8229600,
                            Height = 457200
                        };
                        shape.Paragraphs.Add(para);
                        WriteTextBoxShape(writer, shape, shapeId, slideNum, fallbackShapeIndex++);
                        shapeId++;
                        yPos += 500000;
                    }
                }
            }
            
            writer.WriteEndElement(); // spTree
            writer.WriteEndElement(); // cSld
            
            // 写入切换动画
            if (slide.Transition != null)
            {
                WriteTransitionXml(writer, slide.Transition);
            }
            
            writer.WriteEndElement(); // sld
            writer.WriteEndDocument();
        }
        
        private void WriteTransitionXml(XmlWriter writer, SlideTransition transition)
        {
            writer.WriteStartElement("p", "transition", NS_P);
            
            if (transition.Speed != "fast")
            {
                writer.WriteAttributeString("spd", transition.Speed);
            }
            
            if (transition.HasAutoAdvance && transition.AdvanceTime > 0)
            {
                writer.WriteAttributeString("p", "advTm", NS_P, transition.AdvanceTime.ToString());
            }
            
            // write the specific transition effect element
            if (transition.Type != "none")
            {
                writer.WriteStartElement("p", transition.Type, NS_P);
                writer.WriteEndElement();
            }
            
            writer.WriteEndElement(); // transition
        }
        
        private void WriteGroupShapeProperties(XmlWriter writer)
        {
            // nvGrpSpPr
            writer.WriteStartElement("p", "nvGrpSpPr", NS_P);
            
            writer.WriteStartElement("p", "cNvPr", NS_P);
            writer.WriteAttributeString("id", "1");
            writer.WriteAttributeString("name", "");
            writer.WriteEndElement();
            
            writer.WriteStartElement("p", "cNvGrpSpPr", NS_P);
            writer.WriteEndElement();
            
            writer.WriteStartElement("p", "nvPr", NS_P);
            writer.WriteEndElement();
            
            writer.WriteEndElement(); // nvGrpSpPr
            
            // grpSpPr
            writer.WriteStartElement("p", "grpSpPr", NS_P);
            
            writer.WriteStartElement("a", "xfrm", NS_A);
            writer.WriteStartElement("a", "off", NS_A);
            writer.WriteAttributeString("x", "0");
            writer.WriteAttributeString("y", "0");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ext", NS_A);
            writer.WriteAttributeString("cx", "0");
            writer.WriteAttributeString("cy", "0");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "chOff", NS_A);
            writer.WriteAttributeString("x", "0");
            writer.WriteAttributeString("y", "0");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "chExt", NS_A);
            writer.WriteAttributeString("cx", "0");
            writer.WriteAttributeString("cy", "0");
            writer.WriteEndElement();
            writer.WriteEndElement(); // xfrm
            
            writer.WriteEndElement(); // grpSpPr
        }
        
        private void WriteTextBoxShape(XmlWriter writer, Shape shape, int shapeId, int slideNum, int shapeIndex)
        {
            writer.WriteStartElement("p", "sp", NS_P);
            
            // nvSpPr
            writer.WriteStartElement("p", "nvSpPr", NS_P);
            
            writer.WriteStartElement("p", "cNvPr", NS_P);
            writer.WriteAttributeString("id", shapeId.ToString());
            writer.WriteAttributeString("name", $"TextBox {shapeId}");
            
            if (!string.IsNullOrEmpty(shape.Hyperlink))
            {
                int hId = 1000 + (slideNum * 100) + shapeIndex;
                writer.WriteStartElement("a", "hlinkClick", NS_A);
                writer.WriteAttributeString("r", "id", NS_R, $"hId{hId}");
                writer.WriteEndElement();
            }
            
            writer.WriteEndElement();
            
            writer.WriteStartElement("p", "cNvSpPr", NS_P);
            writer.WriteAttributeString("txBox", "1");
            writer.WriteEndElement();
            
            writer.WriteStartElement("p", "nvPr", NS_P);
            writer.WriteEndElement();
            
            writer.WriteEndElement(); // nvSpPr
            
            // spPr (使用 a: 命名空间)
            writer.WriteStartElement("p", "spPr", NS_P);
            
            writer.WriteStartElement("a", "xfrm", NS_A);
            writer.WriteStartElement("a", "off", NS_A);
            writer.WriteAttributeString("x", shape.Left.ToString());
            writer.WriteAttributeString("y", shape.Top.ToString());
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ext", NS_A);
            // Height and width must be non-negative in OpenXML
            writer.WriteAttributeString("cx", Math.Max(0, Math.Abs(shape.Width)).ToString());
            writer.WriteAttributeString("cy", Math.Max(0, Math.Abs(shape.Height)).ToString());
            writer.WriteEndElement();
            writer.WriteEndElement(); // xfrm
            
            writer.WriteStartElement("a", "prstGeom", NS_A);
            writer.WriteAttributeString("prst", "rect");
            writer.WriteStartElement("a", "avLst", NS_A);
            writer.WriteEndElement();
            writer.WriteEndElement(); // prstGeom
            
            // 填充颜色
            if (!string.IsNullOrEmpty(shape.FillColor))
            {
                writer.WriteStartElement("a", "solidFill", NS_A);
                writer.WriteStartElement("a", "srgbClr", NS_A);
                writer.WriteAttributeString("val", shape.FillColor);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            else
            {
                writer.WriteStartElement("a", "noFill", NS_A);
                writer.WriteEndElement();
            }
            
            // Line/border
            if (!string.IsNullOrEmpty(shape.LineColor))
            {
                writer.WriteStartElement("a", "ln", NS_A);
                writer.WriteStartElement("a", "solidFill", NS_A);
                writer.WriteStartElement("a", "srgbClr", NS_A);
                writer.WriteAttributeString("val", shape.LineColor);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            
            writer.WriteEndElement(); // spPr
            
            // txBody
            writer.WriteStartElement("p", "txBody", NS_P);
            
            writer.WriteStartElement("a", "bodyPr", NS_A);
            writer.WriteAttributeString("wrap", "square");
            writer.WriteAttributeString("rtlCol", "0");
            writer.WriteEndElement();
            
            writer.WriteStartElement("a", "lstStyle", NS_A);
            writer.WriteEndElement();
            
            // 段落
            int runIndexOffset = 0;
            if (shape.Paragraphs.Count > 0)
            {
                foreach (var para in shape.Paragraphs)
                {
                    WriteParagraph(writer, para, slideNum, shapeIndex, runIndexOffset);
                    runIndexOffset += para.Runs.Count;
                }
            }
            else if (!string.IsNullOrEmpty(shape.Text))
            {
                // 使用简单文本
                var simplePara = new TextParagraph();
                simplePara.Runs.Add(new TextRun { Text = shape.Text });
                WriteParagraph(writer, simplePara, slideNum, shapeIndex, 0);
            }
            else
            {
                // 空段落
                writer.WriteStartElement("a", "p", NS_A);
                writer.WriteStartElement("a", "endParaRPr", NS_A);
                writer.WriteAttributeString("lang", "en-US");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            
            writer.WriteEndElement(); // txBody
            
            writer.WriteEndElement(); // sp
        }

        private void WriteGraphicFrame(XmlWriter writer, Shape shape, int shapeId)
        {
            writer.WriteStartElement("p", "graphicFrame", NS_P);
            
            writer.WriteStartElement("p", "nvGraphicFramePr", NS_P);
            writer.WriteStartElement("p", "cNvPr", NS_P);
            writer.WriteAttributeString("id", shapeId.ToString());
            writer.WriteAttributeString("name", $"Table {shapeId}");
            writer.WriteEndElement();
            writer.WriteStartElement("p", "cNvGraphicFramePr", NS_P);
            writer.WriteEndElement();
            writer.WriteStartElement("p", "nvPr", NS_P);
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("p", "xfrm", NS_P);
            writer.WriteStartElement("a", "off", NS_A);
            writer.WriteAttributeString("x", shape.Left.ToString());
            writer.WriteAttributeString("y", shape.Top.ToString());
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ext", NS_A);
            writer.WriteAttributeString("cx", shape.Width.ToString());
            writer.WriteAttributeString("cy", shape.Height.ToString());
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("a", "graphic", NS_A);
            writer.WriteStartElement("a", "graphicData", NS_A);
            writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");

            WriteTable(writer, shape.Table);

            writer.WriteEndElement(); // graphicData
            writer.WriteEndElement(); // graphic
            
            writer.WriteEndElement(); // graphicFrame
        }

        private void WriteTable(XmlWriter writer, Table table)
        {
            writer.WriteStartElement("a", "tbl", NS_A);
            
            writer.WriteStartElement("a", "tblPr", NS_A);
            writer.WriteAttributeString("firstRow", "1");
            writer.WriteAttributeString("bandRow", "1");
            writer.WriteStartElement("a", "tableStyleId", NS_A);
            writer.WriteString("{5C1824F4-7C08-47E0-8931-811103223AAE}"); // Medium Style 2 - Accent 1
            writer.WriteEndElement();
            writer.WriteEndElement();

            // tblGrid
            writer.WriteStartElement("a", "tblGrid", NS_A);
            if (table.ColumnWidths.Count > 0)
            {
                foreach (var width in table.ColumnWidths)
                {
                    writer.WriteStartElement("a", "gridCol", NS_A);
                    writer.WriteAttributeString("w", width.ToString());
                    writer.WriteEndElement();
                }
            }
            else if (table.Rows.Count > 0)
            {
                int colCount = table.Rows[0].Cells.Count;
                // Rough estimate for col width if not available
                long gridColWidth = 9144000 / Math.Max(1, colCount); 
                for (int i = 0; i < colCount; i++)
                {
                    writer.WriteStartElement("a", "gridCol", NS_A);
                    writer.WriteAttributeString("w", gridColWidth.ToString());
                    writer.WriteEndElement();
                }
            }
            writer.WriteEndElement();

            foreach (var row in table.Rows)
            {
                writer.WriteStartElement("a", "tr", NS_A);
                writer.WriteAttributeString("h", row.Height.ToString()); 

                foreach (var cell in row.Cells)
                {
                    writer.WriteStartElement("a", "tc", NS_A);
                    
                    writer.WriteStartElement("a", "txBody", NS_A);
                    writer.WriteStartElement("a", "bodyPr", NS_A);
                    writer.WriteEndElement();
                    writer.WriteStartElement("a", "lstStyle", NS_A);
                    writer.WriteEndElement();

                    if (cell.TextContent != null && cell.TextContent.Count > 0)
                    {
                        foreach (var para in cell.TextContent)
                        {
                            WriteParagraph(writer, para, 0, 0, 0);
                        }
                    }
                    else
                    {
                        writer.WriteStartElement("a", "p", NS_A);
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // txBody

                    writer.WriteStartElement("a", "tcPr", NS_A);
                    writer.WriteAttributeString("marL", cell.MarginLeft.ToString());
                    writer.WriteAttributeString("marT", cell.MarginTop.ToString());
                    writer.WriteAttributeString("marR", cell.MarginRight.ToString());
                    writer.WriteAttributeString("marB", cell.MarginBottom.ToString());
                    writer.WriteAttributeString("anchor", cell.VerticalAlignment);
                    if (!string.IsNullOrEmpty(cell.FillColor))
                    {
                        writer.WriteStartElement("a", "solidFill", NS_A);
                        writer.WriteStartElement("a", "srgbClr", NS_A);
                        writer.WriteAttributeString("val", cell.FillColor);
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // tcPr

                    writer.WriteEndElement(); // tc
                }
                writer.WriteEndElement(); // tr
            }

            writer.WriteEndElement(); // tbl
        }

        private void WritePictureShape(XmlWriter writer, Shape shape, int shapeId, int imageRid, int slideNum, int shapeIndex)
        {
            writer.WriteStartElement("p", "pic", NS_P);

            // nvPicPr
            writer.WriteStartElement("p", "nvPicPr", NS_P);
            writer.WriteStartElement("p", "cNvPr", NS_P);
            writer.WriteAttributeString("id", shapeId.ToString());
            writer.WriteAttributeString("name", $"Picture {shapeId}");
            
            if (!string.IsNullOrEmpty(shape.Hyperlink))
            {
                int hId = 1000 + (slideNum * 100) + shapeIndex;
                writer.WriteStartElement("a", "hlinkClick", NS_A);
                writer.WriteAttributeString("r", "id", NS_R, $"hId{hId}");
                writer.WriteEndElement();
            }
            
            writer.WriteEndElement();
            writer.WriteStartElement("p", "cNvPicPr", NS_P);
            writer.WriteEndElement();
            writer.WriteStartElement("p", "nvPr", NS_P);
            writer.WriteEndElement();
            writer.WriteEndElement(); // nvPicPr

            // blipFill
            writer.WriteStartElement("p", "blipFill", NS_P);
            writer.WriteStartElement("a", "blip", NS_A);
            writer.WriteAttributeString("r", "embed", NS_R, $"rId{imageRid}");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "stretch", NS_A);
            writer.WriteStartElement("a", "fillRect", NS_A);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // blipFill

            // spPr
            writer.WriteStartElement("p", "spPr", NS_P);
            writer.WriteStartElement("a", "xfrm", NS_A);
            writer.WriteStartElement("a", "off", NS_A);
            writer.WriteAttributeString("x", shape.Left.ToString());
            writer.WriteAttributeString("y", shape.Top.ToString());
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ext", NS_A);
            // Height and width must be non-negative in OpenXML
            writer.WriteAttributeString("cx", Math.Max(0, Math.Abs(shape.Width)).ToString());
            writer.WriteAttributeString("cy", Math.Max(0, Math.Abs(shape.Height)).ToString());
            writer.WriteEndElement();
            writer.WriteEndElement(); // xfrm
            writer.WriteStartElement("a", "prstGeom", NS_A);
            writer.WriteAttributeString("prst", "rect");
            writer.WriteStartElement("a", "avLst", NS_A);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // spPr

            writer.WriteEndElement(); // pic
        }
        
        private void WriteParagraph(XmlWriter writer, TextParagraph para, int slideNum, int shapeIndex, int runIndexOffset)
        {
            writer.WriteStartElement("a", "p", NS_A);
            
            // 段落属性
            writer.WriteStartElement("a", "pPr", NS_A);
            switch (para.Alignment)
            {
                case TextAlignment.Center:
                    writer.WriteAttributeString("algn", "ctr");
                    break;
                case TextAlignment.Right:
                    writer.WriteAttributeString("algn", "r");
                    break;
                case TextAlignment.Justify:
                    writer.WriteAttributeString("algn", "just");
                    break;
                default:
                    writer.WriteAttributeString("algn", "l");
                    break;
            }
            writer.WriteEndElement();
            
            // 文本运行
            for (int i = 0; i < para.Runs.Count; i++)
            {
                var run = para.Runs[i];
                writer.WriteStartElement("a", "r", NS_A);
                
                // 运行属性
                writer.WriteStartElement("a", "rPr", NS_A);
                writer.WriteAttributeString("lang", "en-US");
                writer.WriteAttributeString("dirty", "0");
                
                if (run.FontSize > 0)
                    writer.WriteAttributeString("sz", run.FontSize.ToString());
                if (run.Bold)
                    writer.WriteAttributeString("b", "1");
                if (run.Italic)
                    writer.WriteAttributeString("i", "1");
                if (run.Underline)
                    writer.WriteAttributeString("u", "sng");
                    
                // 字体颜色
                if (!string.IsNullOrEmpty(run.Color))
                {
                    writer.WriteStartElement("a", "solidFill", NS_A);
                    writer.WriteStartElement("a", "srgbClr", NS_A);
                    writer.WriteAttributeString("val", run.Color);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                
                // 字体
                if (!string.IsNullOrEmpty(run.FontName))
                {
                    writer.WriteStartElement("a", "latin", NS_A);
                    writer.WriteAttributeString("typeface", run.FontName);
                    writer.WriteEndElement();
                    writer.WriteStartElement("a", "ea", NS_A);
                    writer.WriteAttributeString("typeface", run.FontName);
                    writer.WriteEndElement();
                }
                
                // Hyperlink
                if (!string.IsNullOrEmpty(run.Hyperlink))
                {
                    int hId = 2000 + (slideNum * 1000) + (shapeIndex * 50) + runIndexOffset + i;
                    writer.WriteStartElement("a", "hlinkClick", NS_A);
                    writer.WriteAttributeString("r", "id", NS_R, $"hId{hId}");
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement(); // rPr
                
                writer.WriteStartElement("a", "t", NS_A);
                writer.WriteString(run.Text ?? "");
                writer.WriteEndElement();
                
                writer.WriteEndElement(); // r
            }
            
            writer.WriteEndElement(); // p
        }
        
        #endregion
        
        #region Charts
        
        private void WriteChartFrame(XmlWriter writer, Shape shape, int shapeId, int chartId)
        {
            // 图表使用 graphicFrame 而不是 sp
            writer.WriteStartElement("p", "graphicFrame", NS_P);
            
            // nvGraphicFramePr
            writer.WriteStartElement("p", "nvGraphicFramePr", NS_P);
            
            writer.WriteStartElement("p", "cNvPr", NS_P);
            writer.WriteAttributeString("id", shapeId.ToString());
            writer.WriteAttributeString("name", $"Chart {chartId}");
            writer.WriteEndElement();
            
            writer.WriteStartElement("p", "cNvGraphicFramePr", NS_P);
            writer.WriteEndElement();
            
            writer.WriteStartElement("p", "nvPr", NS_P);
            writer.WriteEndElement();
            
            writer.WriteEndElement(); // nvGraphicFramePr
            
            // xfrm
            writer.WriteStartElement("p", "xfrm", NS_P);
            writer.WriteStartElement("a", "off", NS_A);
            writer.WriteAttributeString("x", Math.Max(0, shape.Left).ToString());
            writer.WriteAttributeString("y", Math.Max(0, shape.Top).ToString());
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ext", NS_A);
            writer.WriteAttributeString("cx", Math.Max(914400, shape.Width).ToString());
            writer.WriteAttributeString("cy", Math.Max(914400, shape.Height).ToString());
            writer.WriteEndElement();
            writer.WriteEndElement(); // xfrm
            
            // graphic
            writer.WriteStartElement("a", "graphic", NS_A);
            writer.WriteStartElement("a", "graphicData", NS_A);
            writer.WriteAttributeString("uri", NS_C);
            
            writer.WriteStartElement("c", "chart", NS_C);
            writer.WriteAttributeString("xmlns", "c", null, NS_C);
            writer.WriteAttributeString("r", "id", NS_R, $"rId{chartId}");
            writer.WriteEndElement();
            
            writer.WriteEndElement(); // graphicData
            writer.WriteEndElement(); // graphic
            
            writer.WriteEndElement(); // graphicFrame
        }
        
        private void WriteChartXml(string baseDir, Chart chart, int chartId)
        {
            var chartDir = Path.Combine(baseDir, "ppt", "charts");
            Directory.CreateDirectory(chartDir);
            
            var path = Path.Combine(chartDir, $"chart{chartId}.xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("c", "chartSpace", NS_C);
            writer.WriteAttributeString("xmlns", "a", null, NS_A);
            writer.WriteAttributeString("xmlns", "r", null, NS_R);
            
            writer.WriteStartElement("c", "chart", NS_C);
            writer.WriteStartElement("c", "plotArea", NS_C);
            
            // Layout
            writer.WriteStartElement("c", "layout", NS_C);
            writer.WriteEndElement();
            
            // Legend
            if (chart.ShowLegend)
            {
                writer.WriteStartElement("c", "legend", NS_C);
                writer.WriteStartElement("c", "legendPos", NS_C);
                writer.WriteAttributeString("val", chart.LegendPosition ?? "r");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "overlay", NS_C);
                writer.WriteAttributeString("val", "0");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            
            if (!string.IsNullOrEmpty(chart.Title))
            {
                writer.WriteStartElement("c", "title", NS_C);
                writer.WriteStartElement("c", "tx", NS_C);
                writer.WriteStartElement("c", "rich", NS_C);
                writer.WriteStartElement("a", "bodyPr", NS_A);
                writer.WriteEndElement();
                writer.WriteStartElement("a", "lstStyle", NS_A);
                writer.WriteEndElement();
                writer.WriteStartElement("a", "p", NS_A);
                writer.WriteStartElement("a", "pPr", NS_A);
                writer.WriteEndElement();
                writer.WriteStartElement("a", "r", NS_A);
                writer.WriteStartElement("a", "rPr", NS_A);
                writer.WriteAttributeString("lang", "en-US");
                writer.WriteEndElement();
                writer.WriteStartElement("a", "t", NS_A);
                writer.WriteString(chart.Title);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteStartElement("c", "layout", NS_C);
                writer.WriteEndElement();
                writer.WriteStartElement("c", "overlay", NS_C);
                writer.WriteAttributeString("val", "0");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            string chartTag = "barChart";
            bool isBar = false;
            bool isLine = false;
            bool isPie = false;
            bool isArea = false;
            bool isScatter = false;
            bool isRadar = false;

            if (chart.Type == "bar") { chartTag = "barChart"; isBar = true; }
            else if (chart.Type == "line") { chartTag = "lineChart"; isLine = true; }
            else if (chart.Type == "pie") { chartTag = "pieChart"; isPie = true; }
            else if (chart.Type == "area") { chartTag = "areaChart"; isArea = true; }
            else if (chart.Type == "scatter") { chartTag = "scatterChart"; isScatter = true; }
            else if (chart.Type == "radar") { chartTag = "radarChart"; isRadar = true; }
            else { chartTag = "barChart"; isBar = true; } // Fallback

            writer.WriteStartElement("c", chartTag, NS_C);

            if (isBar)
            {
                writer.WriteStartElement("c", "barDir", NS_C);
                writer.WriteAttributeString("val", "col"); // Default to column
                writer.WriteEndElement();
                writer.WriteStartElement("c", "grouping", NS_C);
                writer.WriteAttributeString("val", "clustered");
                writer.WriteEndElement();
            }
            else if (isLine)
            {
                writer.WriteStartElement("c", "grouping", NS_C);
                writer.WriteAttributeString("val", "standard");
                writer.WriteEndElement();
            }
            else if (isPie)
            {
                writer.WriteStartElement("c", "varyColors", NS_C);
                writer.WriteAttributeString("val", "1");
                writer.WriteEndElement();
            }
            
            int serIdx = 0;
            foreach (var series in chart.Series)
            {
                // Simple column letter mapping: Series 0 -> B, Series 1 -> C, etc.
                // Assuming we don't have more than 25 series (Z). Category is typically A.
                char colLetter = (char)('B' + serIdx);
                if (serIdx > 24) colLetter = 'Z'; // Fallback for too many series
                    writer.WriteStartElement("c", "ser", NS_C);
                    
                    writer.WriteStartElement("c", "idx", NS_C);
                    writer.WriteAttributeString("val", serIdx.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("c", "order", NS_C);
                    writer.WriteAttributeString("val", serIdx.ToString());
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("c", "tx", NS_C);
                writer.WriteStartElement("c", "strRef", NS_C);
                writer.WriteStartElement("c", "f", NS_C);
                writer.WriteString($"Sheet1!${colLetter}$1");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "strCache", NS_C);
                writer.WriteStartElement("c", "ptCount", NS_C);
                writer.WriteAttributeString("val", "1");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "pt", NS_C);
                writer.WriteAttributeString("idx", "0");
                writer.WriteStartElement("c", "v", NS_C);
                writer.WriteString(series.Name ?? $"Series {serIdx + 1}");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                // Series formatting (Color and Markers)
                if (!string.IsNullOrEmpty(series.Color))
                {
                    writer.WriteStartElement("c", "spPr", NS_C);
                    writer.WriteStartElement("a", "solidFill", NS_A);
                    writer.WriteStartElement("a", "srgbClr", NS_A);
                    writer.WriteAttributeString("val", series.Color);
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    if (isLine || isScatter || isRadar)
                    {
                        writer.WriteStartElement("a", "ln", NS_A);
                        writer.WriteStartElement("a", "solidFill", NS_A);
                        writer.WriteStartElement("a", "srgbClr", NS_A);
                        writer.WriteAttributeString("val", series.Color);
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // spPr
                }

                if ((isLine || isScatter) && !string.IsNullOrEmpty(series.MarkerType) && series.MarkerType != "none")
                {
                    writer.WriteStartElement("c", "marker", NS_C);
                    writer.WriteStartElement("c", "symbol", NS_C);
                    writer.WriteAttributeString("val", series.MarkerType);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                
                // Categories
                if (series.Categories != null && series.Categories.Count > 0)
                {
                    writer.WriteStartElement("c", "cat", NS_C);
                    writer.WriteStartElement("c", "strRef", NS_C);
                    writer.WriteStartElement("c", "f", NS_C);
                    writer.WriteString($"Sheet1!$A$2:$A${1 + series.Categories.Count}");
                    writer.WriteEndElement();
                    writer.WriteStartElement("c", "strCache", NS_C);
                    writer.WriteStartElement("c", "ptCount", NS_C);
                    writer.WriteAttributeString("val", series.Categories.Count.ToString());
                    writer.WriteEndElement();
                    for (int j = 0; j < series.Categories.Count; j++)
                    {
                        writer.WriteStartElement("c", "pt", NS_C);
                        writer.WriteAttributeString("idx", j.ToString());
                        writer.WriteStartElement("c", "v", NS_C);
                        writer.WriteString(series.Categories[j] ?? "");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // strCache
                    writer.WriteEndElement(); // strRef
                    writer.WriteEndElement(); // cat
                }
                
                // Values
                writer.WriteStartElement("c", "val", NS_C);
                writer.WriteStartElement("c", "numRef", NS_C);
                writer.WriteStartElement("c", "f", NS_C);
                writer.WriteString($"Sheet1!${colLetter}$2:${colLetter}${1 + series.Values.Count}");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "numCache", NS_C);
                writer.WriteStartElement("c", "formatCode", NS_C);
                writer.WriteString("General");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "ptCount", NS_C);
                writer.WriteAttributeString("val", series.Values.Count.ToString());
                writer.WriteEndElement();
                for (int j = 0; j < series.Values.Count; j++)
                {
                    writer.WriteStartElement("c", "pt", NS_C);
                    writer.WriteAttributeString("idx", j.ToString());
                    writer.WriteStartElement("c", "v", NS_C);
                    writer.WriteString(series.Values[j].ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                writer.WriteEndElement(); // ser
                serIdx++;
            }
            
            writer.WriteStartElement("c", "axId", NS_C);
            writer.WriteAttributeString("val", "1");
            writer.WriteEndElement();
            writer.WriteStartElement("c", "axId", NS_C);
            writer.WriteAttributeString("val", "2");
            writer.WriteEndElement();
            
            writer.WriteEndElement(); // chartTag
                
                // Category axis
                writer.WriteStartElement("c", "catAx", NS_C);
                writer.WriteStartElement("c", "axId", NS_C);
                writer.WriteAttributeString("val", "1");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "scaling", NS_C);
                writer.WriteStartElement("c", "orientation", NS_C);
                writer.WriteAttributeString("val", "minMax");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteStartElement("c", "delete", NS_C);
                writer.WriteAttributeString("val", "0");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "axPos", NS_C);
                writer.WriteAttributeString("val", "b");
                writer.WriteEndElement();

                if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
                {
                    writer.WriteStartElement("c", "title", NS_C);
                    writer.WriteStartElement("c", "tx", NS_C);
                    writer.WriteStartElement("c", "rich", NS_C);
                    writer.WriteStartElement("a", "bodyPr", NS_A); writer.WriteEndElement();
                    writer.WriteStartElement("a", "lstStyle", NS_A); writer.WriteEndElement();
                    writer.WriteStartElement("a", "p", NS_A);
                    writer.WriteStartElement("a", "r", NS_A);
                    writer.WriteStartElement("a", "t", NS_A);
                    writer.WriteString(chart.CategoryAxisTitle);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("c", "overlay", NS_C);
                    writer.WriteAttributeString("val", "0");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }

                writer.WriteStartElement("c", "crossAx", NS_C);
                writer.WriteAttributeString("val", "2");
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                // Value axis
                writer.WriteStartElement("c", "valAx", NS_C);
                writer.WriteStartElement("c", "axId", NS_C);
                writer.WriteAttributeString("val", "2");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "scaling", NS_C);
                writer.WriteStartElement("c", "orientation", NS_C);
                writer.WriteAttributeString("val", "minMax");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteStartElement("c", "delete", NS_C);
                writer.WriteAttributeString("val", "0");
                writer.WriteEndElement();
                writer.WriteStartElement("c", "axPos", NS_C);
                writer.WriteAttributeString("val", "l");
                writer.WriteEndElement();

                if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
                {
                    writer.WriteStartElement("c", "title", NS_C);
                    writer.WriteStartElement("c", "tx", NS_C);
                    writer.WriteStartElement("c", "rich", NS_C);
                    writer.WriteStartElement("a", "bodyPr", NS_A); writer.WriteEndElement();
                    writer.WriteStartElement("a", "lstStyle", NS_A); writer.WriteEndElement();
                    writer.WriteStartElement("a", "p", NS_A);
                    writer.WriteStartElement("a", "r", NS_A);
                    writer.WriteStartElement("a", "t", NS_A);
                    writer.WriteString(chart.ValueAxisTitle);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("c", "overlay", NS_C);
                    writer.WriteAttributeString("val", "0");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteStartElement("c", "crossAx", NS_C);
                writer.WriteAttributeString("val", "1");
                writer.WriteEndElement();
                writer.WriteEndElement();
            
            
            writer.WriteEndElement(); // plotArea
            writer.WriteEndElement(); // chart
            writer.WriteEndElement(); // chartSpace
            writer.WriteEndDocument();
        }
        
        #endregion
        
        #region Slide Relationships
        
        private void WriteSlideRelationship(string baseDir, Presentation presentation, Slide slide, int slideNum)
        {
            var path = Path.Combine(baseDir, "ppt", "slides", "_rels", $"slide{slideNum}.xml.rels");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", NS_RELS);
            
            // 每个 slide 必须指向 slideLayout
            WriteRelationship(writer, "rId1", REL_SLIDE_LAYOUT, "../slideLayouts/slideLayout1.xml");
            
            // Charts
            int chartRid = 2;
            int imageRid = 100;
            foreach (var shape in slide.Shapes)
            {
                if (shape.Type == "Chart" && shape.Chart != null)
                {
                    WriteRelationship(writer, $"rId{chartRid}", REL_CHART, $"../charts/chart{chartRid - 1}.xml");
                    chartRid++;
                }
                else if (shape.Type == "Picture" && shape.ImageId != null)
                {
                    var imgInfo = presentation.Images.FirstOrDefault(img => img.Id == shape.ImageId.Value);
                    if (imgInfo != null)
                    {
                        string ext = imgInfo.Extension ?? "png";
                        WriteRelationship(writer, $"rId{imageRid}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", $"../media/image{shape.ImageId}.{ext}");
                        imageRid++;
                    }
                }
                
                // Hyperlinks for shapes
                if (!string.IsNullOrEmpty(shape.Hyperlink))
                {
                    // Use a unique ID range for hyperlinks to avoid collision
                    int hId = 1000 + (slideNum * 100) + slide.Shapes.IndexOf(shape);
                    WriteRelationship(writer, $"hId{hId}", REL_HYPERLINK, shape.Hyperlink, "External");
                }
                
                // Hyperlinks for text runs
                int runIndex = 0;
                foreach (var para in shape.Paragraphs)
                {
                    foreach (var run in para.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.Hyperlink))
                        {
                            int hId = 2000 + (slideNum * 1000) + (slide.Shapes.IndexOf(shape) * 50) + runIndex;
                            WriteRelationship(writer, $"hId{hId}", REL_HYPERLINK, run.Hyperlink, "External");
                        }
                        runIndex++;
                    }
                }
            }

            if (!string.IsNullOrEmpty(slide.Notes))
            {
                WriteRelationship(writer, "rIdNotes", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide", $"../notesSlides/notesSlide{slideNum}.xml");
            }
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        #endregion

        private void WriteNotesSlideXml(string baseDir, Slide slide, int slideNum)
        {
            string notesPath = Path.Combine(baseDir, "ppt", "notesSlides", $"notesSlide{slideNum}.xml");
            using var writer = XmlWriter.Create(notesPath, new XmlWriterSettings { Indent = true });

            writer.WriteStartDocument();
            writer.WriteStartElement("p", "notes", NS_P);
            writer.WriteAttributeString("xmlns", "a", null, NS_A);
            writer.WriteAttributeString("xmlns", "r", null, NS_R);
            writer.WriteAttributeString("xmlns", "p", null, NS_P);

            writer.WriteStartElement("p", "cSld", NS_P);
            writer.WriteStartElement("p", "spTree", NS_P);
            
            // nvGrpSpPr and grpSpPr via helper
            WriteGroupShapeProperties(writer);

            // Notes Textbox
            writer.WriteStartElement("p", "sp", NS_P);
            writer.WriteStartElement("p", "nvSpPr", NS_P);
            writer.WriteStartElement("p", "cNvPr", NS_P);
            writer.WriteAttributeString("id", "2");
            writer.WriteAttributeString("name", "Notes Placeholder");
            writer.WriteEndElement();
            writer.WriteStartElement("p", "cNvSpPr", NS_P);
            writer.WriteStartElement("a", "spLocks", NS_A);
            writer.WriteAttributeString("noGrp", "1");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("p", "nvPr", NS_P);
            writer.WriteStartElement("p", "ph", NS_P);
            writer.WriteAttributeString("type", "body");
            writer.WriteAttributeString("idx", "1");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // nvSpPr

            writer.WriteStartElement("p", "spPr", NS_P);
            writer.WriteEndElement();

            writer.WriteStartElement("p", "txBody", NS_P);
            writer.WriteStartElement("a", "bodyPr", NS_A);
            writer.WriteEndElement();
            writer.WriteStartElement("a", "lstStyle", NS_A);
            writer.WriteEndElement();
            writer.WriteStartElement("a", "p", NS_A);
            writer.WriteStartElement("a", "r", NS_A);
            writer.WriteStartElement("a", "t", NS_A);
            writer.WriteString(slide.Notes);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // txBody
            writer.WriteEndElement(); // sp

            writer.WriteEndElement(); // spTree
            writer.WriteEndElement(); // cSld

            writer.WriteEndElement(); // notes
        }

        private void WriteNotesSlideRelationships(string baseDir, int slideNum)
        {
            string relsPath = Path.Combine(baseDir, "ppt", "notesSlides", "_rels", $"notesSlide{slideNum}.xml.rels");
            using var writer = XmlWriter.Create(relsPath, new XmlWriterSettings { Indent = true });

            writer.WriteStartDocument();
            writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

            writer.WriteStartElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships");
            writer.WriteAttributeString("Id", "rId1");
            writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/package/2006/relationships/slide");
            writer.WriteAttributeString("Target", $"../slides/slide{slideNum}.xml");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }
        
        #region SlideLayout / SlideMaster / Theme
        
        private void WriteSlideLayouts(string baseDir, Presentation presentation)
        {
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                int layoutNum = i + 1;
                var path = Path.Combine(baseDir, "ppt", "slideLayouts", $"slideLayout{layoutNum}.xml");
                using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
                
                writer.WriteStartDocument(true);
                writer.WriteStartElement("p", "sldLayout", NS_P);
                writer.WriteAttributeString("xmlns", "a", null, NS_A);
                writer.WriteAttributeString("xmlns", "r", null, NS_R);
                writer.WriteAttributeString("type", "blank");
                writer.WriteAttributeString("preserve", "1");
                
                writer.WriteStartElement("p", "cSld", NS_P);
                writer.WriteStartElement("p", "spTree", NS_P);
                WriteGroupShapeProperties(writer);
                writer.WriteEndElement(); // spTree
                writer.WriteEndElement(); // cSld
                
                writer.WriteEndElement(); // sldLayout
                writer.WriteEndDocument();
            }
        }
        
        private void WriteSlideLayoutRelationships(string baseDir, Presentation presentation)
        {
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                int layoutNum = i + 1;
                var path = Path.Combine(baseDir, "ppt", "slideLayouts", "_rels", $"slideLayout{layoutNum}.xml.rels");
                using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
                
                writer.WriteStartDocument(true);
                writer.WriteStartElement("Relationships", NS_RELS);
                WriteRelationship(writer, "rId1", REL_SLIDE_MASTER, $"../slideMasters/slideMaster{layoutNum}.xml");
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private void WriteSlideMasters(string baseDir, Presentation presentation)
        {
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                var masterSlide = presentation.Masters.Count > i ? presentation.Masters[i] : new Slide();
                int masterNum = i + 1;
                var path = Path.Combine(baseDir, "ppt", "slideMasters", $"slideMaster{masterNum}.xml");
                using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
                
                writer.WriteStartDocument(true);
                writer.WriteStartElement("p", "sldMaster", NS_P);
                writer.WriteAttributeString("xmlns", "a", null, NS_A);
                writer.WriteAttributeString("xmlns", "r", null, NS_R);
                
                writer.WriteStartElement("p", "cSld", NS_P);
                writer.WriteStartElement("p", "spTree", NS_P);
                WriteGroupShapeProperties(writer);
                
                int shapeId = 2;
                int chartId = 1;
                int imageRid = 100;
                
                foreach (var shape in masterSlide.Shapes)
                {
                    if (shape.Type == "Chart" && shape.Chart != null)
                    {
                        WriteChartFrame(writer, shape, shapeId, chartId);
                        WriteChartXml(baseDir, shape.Chart, chartId);
                        chartId++;
                    }
                    else if (shape.Type == "Table" && shape.Table != null)
                    {
                        WriteGraphicFrame(writer, shape, shapeId);
                    }
                    else if (shape.Type == "Picture" && shape.ImageId != null)
                    {
                        if (presentation.Images.Any(img => img.Id == shape.ImageId.Value))
                        {
                            WritePictureShape(writer, shape, shapeId, imageRid, 1000 + masterNum, masterSlide.Shapes.IndexOf(shape));
                            imageRid++;
                        }
                    }
                    else
                    {
                        WriteTextBoxShape(writer, shape, shapeId, 1000 + masterNum, masterSlide.Shapes.IndexOf(shape));
                    }
                    shapeId++;
                }
                
                writer.WriteEndElement(); // spTree
                writer.WriteEndElement(); // cSld
                
                // clrMap (required by schema before sldLayoutIdLst)
                writer.WriteStartElement("p", "clrMap", NS_P);
                writer.WriteAttributeString("bg1", "lt1");
                writer.WriteAttributeString("tx1", "dk1");
                writer.WriteAttributeString("bg2", "lt2");
                writer.WriteAttributeString("tx2", "dk2");
                writer.WriteAttributeString("accent1", "accent1");
                writer.WriteAttributeString("accent2", "accent2");
                writer.WriteAttributeString("accent3", "accent3");
                writer.WriteAttributeString("accent4", "accent4");
                writer.WriteAttributeString("accent5", "accent5");
                writer.WriteAttributeString("accent6", "accent6");
                writer.WriteAttributeString("hlink", "hlink");
                writer.WriteAttributeString("folHlink", "folHlink");
                writer.WriteEndElement(); // clrMap
                
                // sldLayoutIdLst
                writer.WriteStartElement("p", "sldLayoutIdLst", NS_P);
                writer.WriteStartElement("p", "sldLayoutId", NS_P);
                writer.WriteAttributeString("id", "2147483649");
                writer.WriteAttributeString("r", "id", NS_R, "rId1");
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                writer.WriteEndElement(); // sldMaster
                writer.WriteEndDocument();
            }
        }
        
        private void WriteSlideMasterRelationships(string baseDir, Presentation presentation)
        {
            int masterCount = Math.Max(1, presentation.Masters.Count);
            for (int i = 0; i < masterCount; i++)
            {
                var masterSlide = presentation.Masters.Count > i ? presentation.Masters[i] : new Slide();
                int masterNum = i + 1;
                var path = Path.Combine(baseDir, "ppt", "slideMasters", "_rels", $"slideMaster{masterNum}.xml.rels");
                using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
                
                writer.WriteStartDocument(true);
                writer.WriteStartElement("Relationships", NS_RELS);
                WriteRelationship(writer, "rId1", REL_SLIDE_LAYOUT, $"../slideLayouts/slideLayout{masterNum}.xml");
                WriteRelationship(writer, "rId2", REL_THEME, "../theme/theme1.xml");
                
                int chartRid = 3;
                int imageRid = 100;
                foreach (var shape in masterSlide.Shapes)
                {
                    if (shape.Type == "Chart" && shape.Chart != null)
                    {
                        WriteRelationship(writer, $"rId{chartRid}", REL_CHART, $"../charts/chart{chartRid - 1}.xml");
                        chartRid++;
                    }
                    else if (shape.Type == "Picture" && shape.ImageId != null)
                    {
                        var imgInfo = presentation.Images.FirstOrDefault(img => img.Id == shape.ImageId.Value);
                        if (imgInfo != null)
                        {
                            string ext = imgInfo.Extension ?? "png";
                            WriteRelationship(writer, $"rId{imageRid}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", $"../media/image{shape.ImageId}.{ext}");
                            imageRid++;
                        }
                    }
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private void WriteTheme(string baseDir, Presentation presentation)
        {
            var path = Path.Combine(baseDir, "ppt", "theme", "theme1.xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            // Look for a color scheme in the first master slide, or fallback to default Office Colors
            ColorScheme scheme = null;
            if (presentation.Masters.Count > 0 && presentation.Masters[0].ColorScheme != null)
            {
                scheme = presentation.Masters[0].ColorScheme;
            }
            // Use defaults if none found
            scheme ??= new ColorScheme()
            {
                Background = "FFFFFF",
                TextAndLines = "000000",
                Shadows = "44546A", // Actually dk2 in PPTX usually
                TitleText = "E7E6E6", // Actually lt2 in PPTX usually
                Fills = "4472C4",
                Accent = "ED7D31",
                AccentAndHyperlink = "0563C1",
                AccentAndFollowingHyperlink = "954F72"
            };
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("a", "theme", NS_A);
            writer.WriteAttributeString("name", "Office Theme");
            
            writer.WriteStartElement("a", "themeElements", NS_A);
            
            // clrScheme — 完整的 12 色方案
            writer.WriteStartElement("a", "clrScheme", NS_A);
            writer.WriteAttributeString("name", "Office");
            
            // Map legacy color scheme to PPTX theme colors
            WriteColorElement(writer, "dk1", scheme.TextAndLines);
            WriteColorElement(writer, "lt1", scheme.Background);
            WriteColorElement(writer, "dk2", scheme.TitleText);
            WriteColorElement(writer, "lt2", scheme.Shadows);
            WriteColorElement(writer, "accent1", scheme.Fills);
            WriteColorElement(writer, "accent2", scheme.Accent);
            WriteColorElement(writer, "accent3", "A5A5A5"); // Fallbacks for extra PPTX accents
            WriteColorElement(writer, "accent4", "FFC000");
            WriteColorElement(writer, "accent5", "5B9BD5");
            WriteColorElement(writer, "accent6", "70AD47");
            WriteColorElement(writer, "hlink", scheme.AccentAndHyperlink);
            WriteColorElement(writer, "folHlink", scheme.AccentAndFollowingHyperlink);
            
            writer.WriteEndElement(); // clrScheme
            
            // fontScheme
            writer.WriteStartElement("a", "fontScheme", NS_A);
            writer.WriteAttributeString("name", "Office");
            
            writer.WriteStartElement("a", "majorFont", NS_A);
            writer.WriteStartElement("a", "latin", NS_A);
            writer.WriteAttributeString("typeface", "Calibri Light");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ea", NS_A);
            writer.WriteAttributeString("typeface", "");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "cs", NS_A);
            writer.WriteAttributeString("typeface", "");
            writer.WriteEndElement();
            writer.WriteEndElement(); // majorFont
            
            writer.WriteStartElement("a", "minorFont", NS_A);
            writer.WriteStartElement("a", "latin", NS_A);
            writer.WriteAttributeString("typeface", "Calibri");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ea", NS_A);
            writer.WriteAttributeString("typeface", "");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "cs", NS_A);
            writer.WriteAttributeString("typeface", "");
            writer.WriteEndElement();
            writer.WriteEndElement(); // minorFont
            
            writer.WriteEndElement(); // fontScheme
            
            // fmtScheme
            writer.WriteStartElement("a", "fmtScheme", NS_A);
            writer.WriteAttributeString("name", "Office");
            
            // fillStyleLst (需要至少 3 个)
            writer.WriteStartElement("a", "fillStyleLst", NS_A);
            writer.WriteStartElement("a", "solidFill", NS_A);
            writer.WriteStartElement("a", "schemeClr", NS_A);
            writer.WriteAttributeString("val", "phClr");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("a", "gradFill", NS_A);
            writer.WriteAttributeString("rotWithShape", "1");
            writer.WriteStartElement("a", "gsLst", NS_A);
            writer.WriteStartElement("a", "gs", NS_A);
            writer.WriteAttributeString("pos", "0");
            writer.WriteStartElement("a", "schemeClr", NS_A);
            writer.WriteAttributeString("val", "phClr");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("a", "gs", NS_A);
            writer.WriteAttributeString("pos", "100000");
            writer.WriteStartElement("a", "schemeClr", NS_A);
            writer.WriteAttributeString("val", "phClr");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // gsLst
            writer.WriteStartElement("a", "lin", NS_A);
            writer.WriteAttributeString("ang", "5400000");
            writer.WriteAttributeString("scaled", "0");
            writer.WriteEndElement();
            writer.WriteEndElement(); // gradFill
            writer.WriteStartElement("a", "gradFill", NS_A);
            writer.WriteAttributeString("rotWithShape", "1");
            writer.WriteStartElement("a", "gsLst", NS_A);
            writer.WriteStartElement("a", "gs", NS_A);
            writer.WriteAttributeString("pos", "0");
            writer.WriteStartElement("a", "schemeClr", NS_A);
            writer.WriteAttributeString("val", "phClr");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("a", "gs", NS_A);
            writer.WriteAttributeString("pos", "100000");
            writer.WriteStartElement("a", "schemeClr", NS_A);
            writer.WriteAttributeString("val", "phClr");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // gsLst
            writer.WriteStartElement("a", "lin", NS_A);
            writer.WriteAttributeString("ang", "5400000");
            writer.WriteAttributeString("scaled", "0");
            writer.WriteEndElement();
            writer.WriteEndElement(); // gradFill
            writer.WriteEndElement(); // fillStyleLst
            
            // lnStyleLst (至少 3 个)
            writer.WriteStartElement("a", "lnStyleLst", NS_A);
            for (int i = 0; i < 3; i++)
            {
                writer.WriteStartElement("a", "ln", NS_A);
                writer.WriteAttributeString("w", ((i + 1) * 6350).ToString());
                writer.WriteAttributeString("cap", "flat");
                writer.WriteAttributeString("cmpd", "sng");
                writer.WriteAttributeString("algn", "ctr");
                writer.WriteStartElement("a", "solidFill", NS_A);
                writer.WriteStartElement("a", "schemeClr", NS_A);
                writer.WriteAttributeString("val", "phClr");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteStartElement("a", "prstDash", NS_A);
                writer.WriteAttributeString("val", "solid");
                writer.WriteEndElement();
                writer.WriteEndElement(); // ln
            }
            writer.WriteEndElement(); // lnStyleLst
            
            // effectStyleLst (至少 3 个)
            writer.WriteStartElement("a", "effectStyleLst", NS_A);
            for (int i = 0; i < 3; i++)
            {
                writer.WriteStartElement("a", "effectStyle", NS_A);
                writer.WriteStartElement("a", "effectLst", NS_A);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); // effectStyleLst
            
            // bgFillStyleLst (至少 3 个)
            writer.WriteStartElement("a", "bgFillStyleLst", NS_A);
            for (int i = 0; i < 3; i++)
            {
                writer.WriteStartElement("a", "solidFill", NS_A);
                writer.WriteStartElement("a", "schemeClr", NS_A);
                writer.WriteAttributeString("val", "phClr");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); // bgFillStyleLst
            
            writer.WriteEndElement(); // fmtScheme
            
            writer.WriteEndElement(); // themeElements
            
            writer.WriteStartElement("a", "objectDefaults", NS_A);
            writer.WriteEndElement();
            writer.WriteStartElement("a", "extraClrSchemeLst", NS_A);
            writer.WriteEndElement();
            
            writer.WriteEndElement(); // theme
            writer.WriteEndDocument();
        }
        
        private void WriteColorElement(XmlWriter writer, string elementName, string colorHex)
        {
            writer.WriteStartElement("a", elementName, NS_A);
            writer.WriteStartElement("a", "srgbClr", NS_A);
            writer.WriteAttributeString("val", colorHex);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        
        #endregion
        
        #region Properties
        
        private void WriteCoreProperties(string baseDir)
        {
            var path = Path.Combine(baseDir, "docProps", "core.xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("cp", "coreProperties", NS_CP);
            writer.WriteAttributeString("xmlns", "dc", null, NS_DC);
            writer.WriteAttributeString("xmlns", "dcterms", null, NS_DCTERMS);
            writer.WriteAttributeString("xmlns", "dcmitype", null, NS_DCMITYPE);
            writer.WriteAttributeString("xmlns", "xsi", null, NS_XSI);
            
            writer.WriteStartElement("dc", "title", NS_DC);
            writer.WriteString("Presentation");
            writer.WriteEndElement();
            
            writer.WriteStartElement("dc", "creator", NS_DC);
            writer.WriteString("NPptToPptx Converter");
            writer.WriteEndElement();
            
            writer.WriteStartElement("dcterms", "created", NS_DCTERMS);
            writer.WriteAttributeString("xsi", "type", NS_XSI, "dcterms:W3CDTF");
            writer.WriteString(DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
            writer.WriteEndElement();
            
            writer.WriteStartElement("dcterms", "modified", NS_DCTERMS);
            writer.WriteAttributeString("xsi", "type", NS_XSI, "dcterms:W3CDTF");
            writer.WriteString(DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
            writer.WriteEndElement();
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        private void WriteExtendedProperties(string baseDir)
        {
            var path = Path.Combine(baseDir, "docProps", "app.xml");
            using var writer = XmlWriter.Create(path, new XmlWriterSettings { Indent = true });
            
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Properties", NS_EP);
            writer.WriteAttributeString("xmlns", "vt", null, "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            
            writer.WriteElementString("Application", NS_EP, "Microsoft Office PowerPoint");
            writer.WriteElementString("AppVersion", NS_EP, "16.0000");
            writer.WriteElementString("Slides", NS_EP, "1");
            writer.WriteElementString("HiddenSlides", NS_EP, "0");
            writer.WriteElementString("ScaleCrop", NS_EP, "false");
            writer.WriteElementString("LinksUpToDate", NS_EP, "false");
            writer.WriteElementString("SharedDoc", NS_EP, "false");
            writer.WriteElementString("HyperlinksChanged", NS_EP, "false");
            
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        
        #endregion
        
        #region VBA
        
        private void WriteVbaProject(string baseDir, VbaProject vbaProject)
        {
            if (vbaProject?.ProjectData != null)
            {
                var vbaDir = Path.Combine(baseDir, "ppt", "vba");
                Directory.CreateDirectory(vbaDir);
                File.WriteAllBytes(Path.Combine(vbaDir, "vbaProject.bin"), vbaProject.ProjectData);
            }
        }
        
        #endregion
        
        #region Packaging
        
        private void PackageAsPptx(string sourceDir, string targetPath)
        {
            if (File.Exists(targetPath))
                File.Delete(targetPath);
            
            using var zipArchive = ZipFile.Open(targetPath, ZipArchiveMode.Create);
            var files = Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories);
            
            foreach (var file in files)
            {
                // 使用正斜杠作为 ZIP 路径分隔符
                string relativePath = Path.GetRelativePath(sourceDir, file).Replace('\\', '/');
                zipArchive.CreateEntryFromFile(file, relativePath, CompressionLevel.Optimal);
            }
        }
        
        #endregion
        
        #region Helpers
        
        private void WriteMediaFiles(string baseDir, Presentation presentation)
        {
            if (presentation.Images.Count == 0) return;
            
            var mediaDir = Path.Combine(baseDir, "ppt", "media");
            Directory.CreateDirectory(mediaDir);
            
            foreach (var img in presentation.Images)
            {
                string ext = img.Extension ?? "png";
                string path = Path.Combine(mediaDir, $"image{img.Id}.{ext}");
                File.WriteAllBytes(path, img.Data);
            }
        }
        
        private void WriteRelationship(XmlWriter writer, string id, string type, string target, string targetMode = null)
        {
            writer.WriteStartElement("Relationship", NS_RELS);
            writer.WriteAttributeString("Id", id);
            writer.WriteAttributeString("Type", type);
            writer.WriteAttributeString("Target", target);
            if (!string.IsNullOrEmpty(targetMode))
            {
                writer.WriteAttributeString("TargetMode", targetMode);
            }
            writer.WriteEndElement();
        }
        
        private int CountCharts(Presentation presentation)
        {
            int count = 0;
            foreach (var slide in presentation.Slides)
            {
                foreach (var shape in slide.Shapes)
                {
                    if (shape.Type == "Chart" && shape.Chart != null)
                        count++;
                }
            }
            return count;
        }
        
        public void Dispose()
        {
        }
        
        #endregion
    }
}
