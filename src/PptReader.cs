using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Globalization;

namespace Nefdev.PptToPptx
{
    public class PptReader : IDisposable
    {
        private readonly Stream _stream;
        private readonly ConversionOptions? _options;
        private readonly Action<string>? _log;
        private byte[]? _picturesData;
        private OleCompoundFile? _oleFile;
        
        private struct OleObjectInfo
        {
            public string StorageName;
            public string ProgId;
        }

        // exObjId → OLE storage name / progId (e.g., "_1326458456")
        private Dictionary<int, OleObjectInfo> _exOleObjMap = new Dictionary<int, OleObjectInfo>();
        
        // PPT Record type constants
        private const ushort RT_Document = 1000;
        private const ushort RT_Slide = 1006;
        private const ushort RT_SlideListWithText = 1008;
        private const ushort RT_MainMaster = 1016;
        private const ushort RT_SlideMasterAtom = 1017;
        private const ushort RT_SlidePersistAtom = 1011;
        private const ushort RT_TextCharsAtom = 4000;  // 0x0FA0 — Unicode text
        private const ushort RT_TextBytesAtom = 4008;  // 0x0FA8 — ANSI text
        private const ushort RT_StyleTextPropAtom = 4001;
        private const ushort RT_TextHeaderAtom = 3999;  // 0x0F9F
        private const ushort RT_UserEditAtom = 4085;  // 0x0FF5
        private const ushort RT_PersistDirectoryAtom = 6002;  // 0x1772
        private const ushort RT_CurrentUserAtom = 4086;
        private const ushort RT_SlideAtom = 1007;
        private const ushort RT_Notes = 1008;
        private const ushort RT_NotesAtom = 1009;
        private const ushort RT_Environment = 1010;
        private const ushort RT_SlideShowSlideInfoAtom = 1012;
        private const ushort RT_DocumentAtom = 1001;
        private const ushort RT_ColorSchemeAtom = 2032;
        private const ushort RT_FontCollection = 2005;
        private const ushort RT_FontEntityAtom = 4023;
        
        // Hyperlink records
        private const ushort RT_ExObjList = 1033;
        private const ushort RT_ExHyperlink = 4055;
        private const ushort RT_ExHyperlinkAtom = 4051;
        private const ushort RT_InteractiveInfo = 4082;
        private const ushort RT_InteractiveInfoAtom = 4083;
        private const ushort RT_TextInteractiveInfoAtom = 4084;
        private const ushort RT_CString = 4056;
        private const ushort RT_AnimationInfoContainer = 4072;
        private const ushort RT_AnimationInfoAtom = 4073;
        
        // OLE / ExObj records
        private const ushort RT_ExObjRefAtom = 3009;   // Links shape to exObjId
        private const ushort RT_ExOleObjStg = 4113;    // OLE Object storage reference
        private const ushort RT_ExOleObjAtom = 4035;   // OLE Object atom with exObjId and storage name
        private const ushort RT_ExEmbed = 4044;        // ExEmbed container
        private const ushort RT_ExOleEmbed = 4034;     // ExOleEmbed container
        private const ushort RT_ExOleLink = 4036;      // ExOleLink container
        private const ushort RT_ExObjListAtom = 1034;  // ExObjList atom
        
        // Programmable Tags
        private const ushort RT_ProgTags = 5000;
        private const ushort RT_ProgStringTag = 5001;
        private const ushort RT_ProgBinaryTag = 5002;
        private const ushort RT_BinaryTagData = 5003;
        
        // Escher record types
        private const ushort ESCHER_DggContainer = 0xF000;
        private const ushort ESCHER_BStoreContainer = 0xF001;
        private const ushort ESCHER_DgContainer = 0xF002;
        private const ushort ESCHER_SpgrContainer = 0xF003;
        private const ushort ESCHER_SpContainer = 0xF004;
        private const ushort ESCHER_Sp = 0xF00A;
        private const ushort ESCHER_ClientTextbox = 0xF00D;
        private const ushort ESCHER_ClientData = 0xF011;
        private const ushort ESCHER_ClientAnchor = 0xF010;
        private const ushort ESCHER_ChildAnchor = 0xF00F;
        private const ushort ESCHER_Opt = 0xF00B;
        private const ushort ESCHER_BlipFirst = 0xF018;
        private const ushort ESCHER_BlipLast = 0xF117;
        
        private Dictionary<int, ImageResource> _blipMap = new Dictionary<int, ImageResource>();
        private List<string> _fontTable = new List<string>();
        private Dictionary<int, string> _hyperlinkMap = new Dictionary<int, string>();
        private Dictionary<int, string> _notesIdMap = new Dictionary<int, string>(); // slideIdRef -> notes text

        public PptReader(string path, ConversionOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Input .ppt path must be provided.", nameof(path));
            if (!File.Exists(path))
                throw new FileNotFoundException("Input .ppt file not found.", path);

            _stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            _options = options;
            _log = options?.Log;
        }
        
        public Presentation ReadPresentation()
        {
            EncodingRegistration.EnsureCodePages();

            _picturesData = null;
            _blipMap.Clear();
            _fontTable.Clear();
            _hyperlinkMap.Clear();
            _notesIdMap.Clear();
            _exOleObjMap.Clear();
            
            var presentation = new Presentation();
            
            // 使用 OLE Compound File 解析器
            var oleFile = new OleCompoundFile(_stream);
            _oleFile = oleFile;
            oleFile.Parse();
            
            // 读取 PowerPoint 文档流
            var pptStream = oleFile.GetStream("PowerPoint Document");
            if (pptStream == null)
                throw new InvalidDataException("OLE container does not contain a 'PowerPoint Document' stream.");

            using (pptStream)
            {
                byte[] pptData = ReadAllBytes(pptStream);
                    
                    // 先尝试通过 Current User 找到 UserEdit 链
                    int userEditOffset = -1;
                    var currentUserStream = oleFile.GetStream("Current User");
                    if (currentUserStream != null)
                    {
                        using (currentUserStream)
                        {
                            userEditOffset = ReadCurrentUser(currentUserStream);
                        }
                    }
                    
                    // 读取 Persist 映射表
                    var persistDirectory = BuildPersistDirectory(pptData, userEditOffset);
                    
                    // 读取 Pictures 流
                    var picturesStream = oleFile.GetStream("Pictures");
                    if (picturesStream != null)
                    {
                        using (picturesStream)
                        {
                            _picturesData = ReadAllBytes(picturesStream);
                            _log?.Invoke($"Read Pictures stream, size: {_picturesData.Length}");
                        }
                    }

                    // 提取所有全局超链接
                    GlobalScanForHyperlinks(pptData);

                    // 解析 Document 和 Slides
                    ParsePptData(pptData, presentation, persistDirectory);
                    
                    // 关联 Notes 到 Slide
                    foreach (var slide in presentation.Slides)
                    {
                        if (_notesIdMap.TryGetValue(slide.SlideId, out string notesText))
                        {
                            slide.Notes = notesText;
                        }
                    }

                    // 将 blipMap 中的图片添加到 Presentation
                    presentation.Images.AddRange(_blipMap.Values);
            }
            
            // 如果没有解析到任何幻灯片，尝试直接扫描
            if (presentation.Slides.Count == 0)
            {
                var pptStream2 = oleFile.GetStream("PowerPoint Document");
                if (pptStream2 != null)
                {
                    using (pptStream2)
                    {
                        byte[] pptData = ReadAllBytes(pptStream2);
                        DirectScanForSlides(pptData, presentation);
                    }
                }
            }
            
            // 读取 VBA 项目
            presentation.VbaProject = ReadVbaProject(oleFile);

            // Post-process: attach embedded resources (from OleObject shapes) and mark footer placeholders
            AttachEmbeddedResources(presentation);
            DetectFooterPlaceholders(presentation);

            _log?.Invoke($"ReadPresentation complete. Total slides: {presentation.Slides.Count}, Total images: {presentation.Images.Count}");
            return presentation;
        }

        private void AttachEmbeddedResources(Presentation presentation)
        {
            if (presentation == null) return;
            int nextId = presentation.EmbeddedResources.Count + 1;

            IEnumerable<Slide> allSlides = presentation.Slides.Concat(presentation.Masters);
            foreach (var slide in allSlides)
            {
                foreach (var shape in slide.Shapes)
                {
                    if (shape.Type == "OleObject" && shape.ImageData != null && shape.ImageData.Length > 0)
                    {
                        var res = new EmbeddedResource
                        {
                            Id = nextId++,
                            Kind = (shape.ImageContentType != null && (shape.ImageContentType.StartsWith("audio/") || shape.ImageContentType.StartsWith("video/"))) ? "media" : "ole",
                            ProgId = null,
                            FileName = $"object{nextId - 1}",
                            Extension = GuessExtFromContentType(shape.ImageContentType),
                            ContentType = shape.ImageContentType,
                            Data = shape.ImageData
                        };
                        presentation.EmbeddedResources.Add(res);

                        shape.EmbeddedResourceId = res.Id;

                        // Clear temporary stash fields to avoid bloating other code paths
                        shape.ImageData = null;
                        shape.ImageContentType = null;
                    }
                }
            }
        }

        private string GuessExtFromContentType(string? contentType)
        {
            if (string.IsNullOrEmpty(contentType)) return "bin";
            return contentType switch
            {
                "audio/wav" => "wav",
                "audio/mpeg" => "mp3",
                "video/mp4" => "mp4",
                "application/vnd.openxmlformats-officedocument.oleObject" => "bin",
                _ => "bin"
            };
        }

        private void DetectFooterPlaceholders(Presentation presentation)
        {
            if (presentation == null) return;
            long h = presentation.SlideHeight > 0 ? presentation.SlideHeight : 6858000;
            long footerBandTop = h - 600000; // ~0.66 inch from bottom

            foreach (var slide in presentation.Slides)
            {
                foreach (var shape in slide.Shapes)
                {
                    if (shape.Type != "TextBox" && shape.Type != "Rectangle") continue;
                    if (shape.Top < footerBandTop) continue;

                    string txt = shape.Text ?? string.Join("", shape.Paragraphs.ConvertAll(p => p.GetPlainText()));
                    txt = txt?.Trim();
                    if (string.IsNullOrEmpty(txt)) continue;

                    // Slide number: small numeric token
                    if (txt.Length <= 4 && int.TryParse(txt, out _))
                    {
                        shape.PlaceholderType = "sldNum";
                        shape.Type = "Placeholder";
                        continue;
                    }
                    // Date-ish: contains '/', '-' and digits
                    bool hasDigit = txt.Any(char.IsDigit);
                    bool hasSep = txt.Contains('/') || txt.Contains('-') || txt.Contains('.');
                    if (hasDigit && hasSep && txt.Length <= 20)
                    {
                        shape.PlaceholderType = "dt";
                        shape.Type = "Placeholder";
                        continue;
                    }
                    // Otherwise treat as footer text
                    if (txt.Length <= 80)
                    {
                        shape.PlaceholderType = "ftr";
                        shape.Type = "Placeholder";
                    }
                }
            }
        }

        private int ReadCurrentUser(Stream stream)
        {
            byte[] data = ReadAllBytes(stream);
            if (data.Length < 20) return -1;
            // CurrentUserAtom header is 8 bytes. offsetToCurrentEdit is at byte 16 (8+8)
            return BitConverter.ToInt32(data, 16);
        }

        private Dictionary<int, int> BuildPersistDirectory(byte[] data, int lastUserEditOffset)
        {
            var persistMap = new Dictionary<int, int>();
            int currentOffset = lastUserEditOffset;

            while (currentOffset >= 0 && currentOffset < data.Length - 8)
            {
                var header = ReadRecordHeader(data, currentOffset);
                if (header.RecType == RT_UserEditAtom)
                {
                    int atomStart = currentOffset + 8;
                    // PersistPointers offset is at byte 8 of atom
                    int persistOffset = BitConverter.ToInt32(data, atomStart + 8);
                    int prevEditOffset = BitConverter.ToInt32(data, atomStart + 4);

                    if (persistOffset >= 0 && persistOffset < data.Length - 8)
                    {
                        var pDirHeader = ReadRecordHeader(data, persistOffset);
                        if (pDirHeader.RecType == RT_PersistDirectoryAtom)
                        {
                            ParsePersistDirectory(data, persistOffset + 8, (int)pDirHeader.RecLen, persistMap);
                        }
                    }
                    currentOffset = prevEditOffset;
                }
                else
                {
                    break;
                }
            }
            return persistMap;
        }

        private void ParsePersistDirectory(byte[] data, int start, int length, Dictionary<int, int> persistMap)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            while (pos + 4 <= end)
            {
                uint entry = BitConverter.ToUInt32(data, pos);
                pos += 4;
                int count = (int)(entry >> 20);
                int baseId = (int)(entry & 0xFFFFF);

                for (int i = 0; i < count; i++)
                {
                    if (pos + 4 <= end)
                    {
                        int offset = BitConverter.ToInt32(data, pos);
                        persistMap[baseId + i] = offset;
                        pos += 4;
                    }
                }
            }
        }

        private void ParsePptData(byte[] data, Presentation presentation, Dictionary<int, int> persistMap)
        {
            // Find Document record using persistMap or scan
            // Typically Document is at persist ID 1
            if (persistMap.TryGetValue(1, out int docOffset))
            {
                var header = ReadRecordHeader(data, docOffset);
                if (header.RecType == RT_Document)
                {
                    ParseDocumentContainer(data, docOffset + 8, (int)header.RecLen, presentation, persistMap);
                }
            }
            else
            {
                // Fallback: scan for Document
                int offset = ScanForRecord(data, RT_Document);
                if (offset >= 0)
                {
                    var header = ReadRecordHeader(data, offset);
                    ParseDocumentContainer(data, offset + 8, (int)header.RecLen, presentation, persistMap);
                }
            }

            // Also scan for Notes containers (RT_Notes = 1008) in the persist directory
            foreach (var kvp in persistMap)
            {
                int offset = kvp.Value;
                if (offset >= 0 && offset < data.Length - 8)
                {
                    var header = ReadRecordHeader(data, offset);
                    if (header.RecType == RT_Notes)
                    {
                        ParseNotesContainer(data, offset + 8, (int)header.RecLen);
                    }
                }
            }
        }

        private void ParseDocumentContainer(byte[] data, int start, int length, Presentation presentation, Dictionary<int, int> persistMap)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            // 收集 SlidePersistAtom 信息
            var slidePersistIds = new List<int>();
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == RT_DocumentAtom)
                {
                    // DocumentAtom: slideSize(8 bytes) + notesSize(8 bytes) + ...
                    // slideSize: width(4) + height(4) in master units (1/576 inch)
                    if (header.RecLen >= 8)
                    {
                        int slideW = BitConverter.ToInt32(data, pos + 8);
                        int slideH = BitConverter.ToInt32(data, pos + 8 + 4);
                        // Convert master units to EMU: 1 master unit = 12700/8 EMU = 1587.5 EMU
                        // Actually, DocumentAtom stores in "master units" where 576 units = 1 inch
                        // EMU: 1 inch = 914400 EMU, so 1 master unit = 914400/576 = 1587.5 EMU
                        if (slideW > 0 && slideH > 0)
                        {
                            presentation.SlideWidth = (int)(slideW * 914400L / 576);
                            presentation.SlideHeight = (int)(slideH * 914400L / 576);
                            _log?.Invoke($"Slide size from DocumentAtom: {slideW}x{slideH} master units => {presentation.SlideWidth}x{presentation.SlideHeight} EMU");
                        }
                    }
                }
                else if (header.RecType == RT_Environment)
                {
                    ParseEnvironment(data, pos + 8, (int)header.RecLen);
                }
                else if (header.RecType == ESCHER_DggContainer)
                {
                    _log?.Invoke($"Found DggContainer at {pos}, len {header.RecLen}");
                    ParseDrawingGroup(data, pos + 8, (int)header.RecLen);
                }
                else if (header.RecType == RT_SlideListWithText)
                {
                    ParseSlideListWithText(data, pos + 8, (int)header.RecLen, presentation, persistMap, slidePersistIds, header.RecInstance);
                }
                else if (header.RecType == RT_ExObjList)
                {
                    ParseExObjList(data, pos + 8, (int)header.RecLen);
                }
                
                pos = recordEnd;
                if (pos <= start) break; // 防止无限循环
            }
        }

        private void ParseEnvironment(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == RT_FontCollection)
                {
                    ParseFontCollection(data, pos + 8, (int)header.RecLen);
                }
                else if (header.RecType == RT_ColorSchemeAtom)
                {
                    // For the environment, this defines the default presentation color scheme.
                    // We can store it on a pseudo-slide or just the first master if needed later.
                    // Since it's global, we just parse it as a proof of concept and 
                    // will rely on the slide/master specific ones.
                    var globalScheme = ParseColorSchemeAtom(data, pos + 8, (int)header.RecLen);
                    _log?.Invoke("Parsed Environment Global ColorScheme");
                }
                else if (header.IsContainer)
                {
                    ParseEnvironment(data, pos + 8, (int)header.RecLen);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
        }
        
        private void ParseFontCollection(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == RT_FontEntityAtom && header.RecLen >= 64)
                {
                    // FontEntityAtom: 64 bytes of WCHAR font face name (null-padded)
                    string fontName = Encoding.Unicode.GetString(data, atomStart, 64).TrimEnd('\0');
                    _fontTable.Add(fontName);
                }
                else if (header.IsContainer)
                {
                    ParseFontCollection(data, pos + 8, (int)header.RecLen);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
            
            if (_fontTable.Count > 0)
            {
                _log?.Invoke($"Parsed {_fontTable.Count} fonts: {string.Join(", ", _fontTable)}");
            }
        }

        private void GlobalScanForHyperlinks(byte[] data)
        {
            for (int p = 0; p < data.Length - 8; p += 8)  // Advance by 8 to check every potential record header
            {
                ushort recType = BitConverter.ToUInt16(data, p + 2);
                if (recType == RT_ExHyperlink)
                {
                    uint recLen = BitConverter.ToUInt32(data, p + 4);
                    ParseExHyperlink(data, p + 8, (int)recLen);
                }
            }
        }

        private void ParseExObjList(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == RT_ExOleEmbed || header.RecType == RT_ExEmbed || header.RecType == RT_ExOleLink)
                {
                    // Parse the container to find ExOleObjAtom
                    ParseExOleEmbedContainer(data, atomStart, (int)header.RecLen);
                }
                else if (header.IsContainer)
                {
                    ParseExObjList(data, atomStart, (int)header.RecLen);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
        }
        
        private void ParseExOleEmbedContainer(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            int? exObjId = null;
            string? storageName = null;
            string progId = null;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == RT_ExOleObjAtom && header.RecLen >= 24)
                {
                    // ExOleObjAtom layout:
                    // 0: drawAspect(4), 1: (don't care)(4), 2: exObjId(4), 3: subType(4), ...
                    exObjId = BitConverter.ToInt32(data, atomStart + 8);
                }
                else if (header.RecType == RT_ExOleObjStg && header.RecLen >= 4)
                {
                    // ExOleObjStg contains a 4-byte index used to construct the storage name
                    // The storage name pattern is "_" + (10-digit decimal number)
                    // But the actual mapping often uses a persistence directory reference
                    int stgIndex = BitConverter.ToInt32(data, atomStart);
                    // Storage name in OLE compound file — try common patterns
                    storageName = $"_{stgIndex}";
                }
                else if (header.RecType == RT_CString && header.RecInstance == 0x01)
                {
                    // Instance 0x10 or 0x01 in ExOleEmbed is the ProgId (e.g., "MSGraph.Chart.8")
                    progId = ParseCString(data, atomStart, (int)header.RecLen);
                }
                else if (header.RecType == RT_CString && header.RecInstance == 0x10)
                {
                    progId = ParseCString(data, atomStart, (int)header.RecLen);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
            
            if (exObjId.HasValue && storageName != null)
            {
                _exOleObjMap[exObjId.Value] = new OleObjectInfo { StorageName = storageName, ProgId = progId };
                _log?.Invoke($"OLE ExObj mapping: exObjId={exObjId.Value} -> storage='{storageName}' progId='{progId}'");
            }
        }

        private void ParseExHyperlink(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            int? hyperlinkId = null;
            string url = null;

            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;

                if (header.RecType == RT_ExHyperlinkAtom)
                {
                    if (header.RecLen >= 4)
                    {
                        hyperlinkId = BitConverter.ToInt32(data, atomStart);
                    }
                }
                else if (header.RecType == RT_CString)
                {
                    url = ParseCString(data, atomStart, (int)header.RecLen);
                }

                pos = recordEnd;
                if (pos <= start) break;
            }

            if (hyperlinkId.HasValue && !string.IsNullOrEmpty(url))
            {
                _hyperlinkMap[hyperlinkId.Value] = url;
                _log?.Invoke($"Mapped hyperlink {hyperlinkId.Value} -> {url}");
            }
        }

        private string ParseCString(byte[] data, int start, int length)
        {
            if (length <= 0) return "";
            // PPT RT_CString is Unicode
            return Encoding.Unicode.GetString(data, start, length).TrimEnd('\0');
        }

        private void ParseInteractiveInfo(byte[] data, int start, int length, Shape shape)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;

            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;

                if (header.RecType == RT_InteractiveInfoAtom)
                {
                    if (header.RecLen >= 13)
                    {
                        // Layout per Apache POI / MS-PPT:
                        // 0..3 soundRef, 4..7 hyperlinkId, 8 action, 9 oleVerb, 10 jump, 11 flags, 12 hyperlinkType
                        int hyperlinkId = BitConverter.ToInt32(data, atomStart + 4);
                        byte action = data[atomStart + 8];
                        byte jump = data[atomStart + 10];
                        byte hyperlinkType = data[atomStart + 12];

                        // External hyperlink address (if exists)
                        if (_hyperlinkMap.TryGetValue(hyperlinkId, out string url) && !string.IsNullOrEmpty(url))
                        {
                            shape.Hyperlink = url;
                            _log?.Invoke($"Associated hyperlink {hyperlinkId} ({url}) with shape");
                        }

                        // Internal jump / action mapping (best-effort)
                        // action: 3=JUMP, 4=HYPERLINK
                        if (action == 3) // ACTION_JUMP
                        {
                            shape.ClickAction = jump switch
                            {
                                1 => "ppaction://hlinkshowjump?jump=nextslide",
                                2 => "ppaction://hlinkshowjump?jump=previousslide",
                                3 => "ppaction://hlinkshowjump?jump=firstslide",
                                4 => "ppaction://hlinkshowjump?jump=lastslide",
                                5 => "ppaction://hlinkshowjump?jump=lastslideviewed",
                                6 => "ppaction://hlinkshowjump?jump=endshow",
                                _ => shape.ClickAction
                            };
                        }
                        else if (action == 4 && shape.Hyperlink == null) // ACTION_HYPERLINK but no URL decoded
                        {
                            // Try to map by hyperlink type if possible
                            shape.ClickAction = hyperlinkType switch
                            {
                                0x00 => "ppaction://hlinkshowjump?jump=nextslide",
                                0x01 => "ppaction://hlinkshowjump?jump=previousslide",
                                0x02 => "ppaction://hlinkshowjump?jump=firstslide",
                                0x03 => "ppaction://hlinkshowjump?jump=lastslide",
                                _ => shape.ClickAction
                            };
                        }
                    }
                }
                else if (header.IsContainer)
                {
                    ParseInteractiveInfo(data, atomStart, (int)header.RecLen, shape);
                }

                pos = recordEnd;
                if (pos <= start) break;
            }
        }

        private int ParseInteractiveInfoId(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;

            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;

                if (header.RecType == RT_InteractiveInfoAtom)
                {
                    if (header.RecLen >= 8)
                    {
                        return BitConverter.ToInt32(data, atomStart + 4);
                    }
                }
                else if (header.IsContainer)
                {
                    int res = ParseInteractiveInfoId(data, atomStart, (int)header.RecLen);
                    if (res > 0) return res;
                }

                pos = recordEnd;
                if (pos <= start) break;
            }
            return 0;
        }

        private void ApplyTextHyperlinks(TextParagraph paragraph, List<(int start, int end, int hyperlinkId)> hyperlinks)
        {
            if (hyperlinks.Count == 0 || paragraph.Runs.Count == 0) return;

            int currentPos = 0;
            foreach (var run in paragraph.Runs)
            {
                int runLen = run.Text.Length;
                int runStart = currentPos;
                int runEnd = currentPos + runLen;

                // Check if any hyperlink range covers this run
                // Note: PPT text hyperlinks are at character level. 
                // We'll apply it to the whole run if the run is within the range.
                foreach (var (hStart, hEnd, hId) in hyperlinks)
                {
                    if (hStart <= runStart && hEnd >= runEnd)
                    {
                        if (_hyperlinkMap.TryGetValue(hId, out string url))
                        {
                            run.Hyperlink = url;
                            _log?.Invoke($"Assigned hyperlink '{url}' to TextRun: '{TruncateText(run.Text, 60)}'");
                            break;
                        }
                    }
                }
                currentPos += runLen;
            }
        }

        private int ParseExObjRefAtom(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;

            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;

                if (header.RecType == RT_ExObjRefAtom)
                {
                    if (header.RecLen >= 4)
                    {
                        // The atom contains the exObjId (4 bytes)
                        return BitConverter.ToInt32(data, atomStart);
                    }
                }
                else if (header.IsContainer)
                {
                    int res = ParseExObjRefAtom(data, atomStart, (int)header.RecLen);
                    if (res > 0) return res;
                }

                pos = recordEnd;
                if (pos <= start) break;
            }
            return 0;
        }

        private Chart? TryParseChartFromExObjId(int exObjId)
        {
            var ole = _oleFile;
            if (ole == null) return null;

            string storageName = null;
            if (_exOleObjMap.TryGetValue(exObjId, out OleObjectInfo info))
            {
                storageName = info.StorageName;
            }
            else
            {
                // Fallback: there might not be a mapping, but maybe there's only one storage
                // Or maybe the storage name is just $"_{exObjId}" 
                var storages = ole.GetStoragesByPrefix("_");
                if (storages.Count > 0)
                {
                    // For now just try the first one if we can't map it properly
                    // A better approach would be to parse the OEPlaceholderAtom properly.
                    storageName = storages[0].Name;
                    _log?.Invoke($"Warning: exObjId {exObjId} not in map, trying fallback storage {storageName}");
                }
            }

            if (storageName != null)
            {
                // Most MS Graph/Excel charts store data in a "Workbook" stream inside the OLE storage
                var stream = ole.GetChildStream(storageName, "Workbook");
                
                // Sometimes it's called "Graph Data" or just "Book" or "\x01CompObj"
                if (stream == null)
                    stream = ole.GetChildStream(storageName, "Book");
                    
                if (stream != null)
                {
                    using (stream)
                    {
                        byte[] biffData = ReadAllBytes(stream);
                        var parser = new PptChartParser();
                        try
                        {
                            return parser.ParseChart(biffData, _options);
                        }
                        catch (Exception ex)
                        {
                            _log?.Invoke($"Error parsing chart from {storageName}: {ex.Message}");
                        }
                    }
                }
                else
                {
                    _log?.Invoke($"Warning: No Workbook stream found in OLE storage '{storageName}'. Available entries: {string.Join(", ", ole.GetAllStreamNames())}");
                }
            }
            
            return null;
        }

        private EmbeddedResource? TryExtractEmbeddedResourceFromExObjId(int exObjId)
        {
            var ole = _oleFile;
            if (ole == null) return null;

            if (!_exOleObjMap.TryGetValue(exObjId, out OleObjectInfo info) || string.IsNullOrEmpty(info.StorageName))
                return null;

            // Prefer common payload stream names for embedded packages/media
            string[] candidateNames = new[] { "CONTENTS", "Contents", "\u0001Ole10Native", "Package", "Data", "Media", "CONTENTS1" };
            OleCompoundFile.DirectoryEntry? bestEntry = null;
            foreach (var name in candidateNames)
            {
                var entry = GetChildStreamEntry(info.StorageName, name);
                if (entry != null && entry.Size > 0)
                {
                    bestEntry = entry;
                    break;
                }
            }

            // Fallback: pick the largest child stream (excluding Workbook/Book which are chart data)
            if (bestEntry == null)
            {
                var entries = GetChildStreamEntries(info.StorageName);
                foreach (var e in entries)
                {
                    if (e.Type != OleCompoundFile.DirectoryEntryType.UserStream) continue;
                    if (string.Equals(e.Name, "Workbook", StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(e.Name, "Book", StringComparison.OrdinalIgnoreCase)) continue;
                    if (bestEntry == null || e.Size > bestEntry.Size) bestEntry = e;
                }
            }

            if (bestEntry == null) return null;

            using var stream = ole.GetStream(bestEntry);
            if (stream == null) return null;
            byte[] payload = ReadAllBytes(stream);
            if (payload.Length == 0) return null;

            var res = new EmbeddedResource();
            res.Kind = GuessKindFromProgId(info.ProgId);
            res.ProgId = info.ProgId;
            res.FileName = $"{info.StorageName}_{bestEntry.Name}";
            GuessExtensionAndContentType(payload, out string ext, out string ct);
            res.Extension = ext;
            res.ContentType = ct;
            res.Data = payload;
            return res;
        }

        private string GuessKindFromProgId(string? progId)
        {
            if (string.IsNullOrEmpty(progId)) return "unknown";
            string p = progId.ToLowerInvariant();
            if (p.Contains("sound") || p.Contains("media") || p.Contains("video") || p.Contains("wmplayer"))
                return "media";
            return "ole";
        }

        private void GuessExtensionAndContentType(byte[] payload, out string extension, out string contentType)
        {
            extension = "bin";
            contentType = "application/octet-stream";
            if (payload.Length >= 12)
            {
                // RIFF....WAVE
                if (payload[0] == (byte)'R' && payload[1] == (byte)'I' && payload[2] == (byte)'F' && payload[3] == (byte)'F'
                    && payload[8] == (byte)'W' && payload[9] == (byte)'A' && payload[10] == (byte)'V' && payload[11] == (byte)'E')
                {
                    extension = "wav";
                    contentType = "audio/wav";
                    return;
                }
                // ID3 (mp3)
                if (payload[0] == (byte)'I' && payload[1] == (byte)'D' && payload[2] == (byte)'3')
                {
                    extension = "mp3";
                    contentType = "audio/mpeg";
                    return;
                }
                // MP4 ftyp
                if (payload[4] == (byte)'f' && payload[5] == (byte)'t' && payload[6] == (byte)'y' && payload[7] == (byte)'p')
                {
                    extension = "mp4";
                    contentType = "video/mp4";
                    return;
                }
                // OLE Compound File header
                if (payload[0] == 0xD0 && payload[1] == 0xCF && payload[2] == 0x11 && payload[3] == 0xE0)
                {
                    extension = "bin";
                    contentType = "application/vnd.openxmlformats-officedocument.oleObject";
                    return;
                }
            }
        }

        private List<OleCompoundFile.DirectoryEntry> GetChildStreamEntries(string storageName)
        {
            if (_oleFile == null) return new List<OleCompoundFile.DirectoryEntry>();
            return _oleFile.GetChildEntries(storageName);
        }

        private OleCompoundFile.DirectoryEntry? GetChildStreamEntry(string storageName, string streamName)
        {
            foreach (var e in GetChildStreamEntries(storageName))
            {
                if (e.Type == OleCompoundFile.DirectoryEntryType.UserStream && e.Name == streamName) return e;
            }
            return null;
        }

        private void ParseDrawingGroup(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == ESCHER_BStoreContainer)
                {
                    _log?.Invoke($"Found BStoreContainer at {pos}, len {header.RecLen}");
                    ParseBStore(data, pos + 8, (int)header.RecLen);
                }
                else if (header.IsContainer)
                {
                    ParseDrawingGroup(data, pos + 8, (int)header.RecLen);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
        }

        private void ParseBStore(byte[] data, int start, int length)
        {
            // Count FBSE entries for inline blips, but primarily scan Pictures stream sequentially
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            int blipIndex = 1;
            
            // First, try to extract any inline blips embedded directly in FBSE records
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == 0xF007 && header.RecLen > 36) // FBSE with embedded blip
                {
                    int blipOffset = atomStart + 36;
                    if (blipOffset + 8 <= recordEnd)
                    {
                        var blipHeader = ReadRecordHeader(data, blipOffset);
                        if (blipHeader.RecType >= ESCHER_BlipFirst && blipHeader.RecType <= ESCHER_BlipLast)
                        {
                            ExtractBlip(data, blipOffset, (int)blipHeader.RecLen + 8, blipIndex);
                        }
                    }
                    blipIndex++;
                }
                else if (header.RecType == 0xF007)
                {
                    blipIndex++; // Count the FBSE entry even if no inline blip
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
            
            // Now scan the Pictures stream sequentially — this is the reliable method
            // The Pictures stream contains blip records back-to-back with no container wrapper
            if (_picturesData != null && _picturesData.Length > 8)
            {
                int picPos = 0;
                int picBlipIndex = 1;
                while (picPos + 8 <= _picturesData.Length)
                {
                    var picHeader = ReadRecordHeader(_picturesData, picPos);
                    if (picHeader.RecType >= ESCHER_BlipFirst && picHeader.RecType <= ESCHER_BlipLast 
                        && picHeader.RecLen > 0 && picHeader.RecLen < (uint)(_picturesData.Length - picPos))
                    {
                        if (!_blipMap.ContainsKey(picBlipIndex)) // Don't overwrite inline blips
                        {
                            ExtractBlip(_picturesData, picPos, (int)picHeader.RecLen + 8, picBlipIndex);
                        }
                        picBlipIndex++;
                        picPos += 8 + (int)picHeader.RecLen;
                    }
                    else
                    {
                        // Skip unknown data — try next byte alignment (shouldn't happen in well-formed streams)
                        picPos++;
                    }
                }
            _log?.Invoke($"Pictures stream scan: found {_blipMap.Count} blips");
            }
        }

        private void ExtractBlip(byte[] data, int start, int length, int id)
        {
            var header = ReadRecordHeader(data, start);
            int dataOffset = start + 8;
            
            byte[] imgData = null;
            string ext = "png";
            string contentType = "image/png";
            
            // Blip RecType values:
            //   0xF01A = EMF, 0xF01B = WMF, 0xF01C = PICT
            //   0xF01D = JPEG (instance 0x46A), 0xF01E = PNG (instance 0x6E0)
            //   0xF01F = DIB, 0xF029 = TIFF
            //   0xF01D also used for JPEG with instance 0x46B (secondary UID)
            //   0xF01E also used for PNG with instance 0x6E1 (secondary UID)
            
            if (header.RecType >= 0xF018 && header.RecType <= 0xF117)
            {
                bool isMetafile = (header.RecType >= 0xF01A && header.RecType <= 0xF01C);
                
                // Detect secondary UID: for JPEG/PNG/DIB, instance odd = has secondary UID
                // For metafiles (EMF/WMF/PICT), instance odd = has secondary UID  
                bool hasSecondaryUid = (header.RecInstance & 1) == 1;
                
                int headerSize;
                if (isMetafile)
                {
                    // Metafile blips: rgbUid(16) [+ rgbUid2(16)] + cbSave(4) + rcBounds(8) + ptSize(8) + cbSave(4) + compression(1) + filter(1)
                    // = 16 + 4 + 8 + 8 + 4 + 1 + 1 = 42 bytes, or 58 with secondary UID
                    // Simpler: the fixed header part after UID is 26 bytes
                    headerSize = hasSecondaryUid ? (16 + 16 + 26) : (16 + 26);
                }
                else
                {
                    // Bitmap blips (JPEG/PNG/DIB/TIFF): rgbUid(16) [+ rgbUid2(16)] + marker(1)
                    headerSize = hasSecondaryUid ? (16 + 16 + 1) : (16 + 1);
                }
                
                switch (header.RecType)
                {
                    case 0xF01A: // EMF
                        ext = "emf";
                        contentType = "image/x-emf";
                        break;
                    case 0xF01B: // WMF
                        ext = "wmf";
                        contentType = "image/x-wmf";
                        break;
                    case 0xF01C: // PICT
                        ext = "pict";
                        contentType = "image/x-pict";
                        break;
                    case 0xF01D: // JPEG
                        ext = "jpg";
                        contentType = "image/jpeg";
                        break;
                    case 0xF01E: // PNG
                        ext = "png";
                        contentType = "image/png";
                        break;
                    case 0xF01F: // DIB
                        ext = "bmp";
                        contentType = "image/bmp";
                        break;
                    case 0xF029: // TIFF
                        ext = "tiff";
                        contentType = "image/tiff";
                        break;
                    default:
                        break;
                }
                
                int imgLen = (int)header.RecLen - headerSize;
                if (imgLen > 0 && dataOffset + headerSize + imgLen <= data.Length)
                {
                    imgData = new byte[imgLen];
                    Array.Copy(data, dataOffset + headerSize, imgData, 0, imgLen);
                    
                    // Metafile blips may be zlib-compressed; check compression byte
                    if (isMetafile && headerSize >= 42)
                    {
                        int compressionOffset = dataOffset + headerSize - 2; // compression byte is 2nd-to-last in header
                        if (compressionOffset >= 0 && compressionOffset < data.Length)
                        {
                            byte compression = data[compressionOffset];
                            if (compression == 0x00) // 0 = deflate compressed
                            {
                                try
                                {
                                    // Read uncompressed size from cbSave-preceding field (cbSize at offset UID+4)
                                    using var compStream = new System.IO.MemoryStream(imgData);
                                    using var deflate = new System.IO.Compression.DeflateStream(compStream, System.IO.Compression.CompressionMode.Decompress);
                                    using var outStream = new System.IO.MemoryStream();
                                    deflate.CopyTo(outStream);
                                    imgData = outStream.ToArray();
                                }
                                catch
                                {
                                    // If decompression fails, keep original data
                                }
                            }
                        }
                    }
                }
            }
            
            if (imgData != null && imgData.Length > 0)
            {
                _blipMap[id] = new ImageResource { Id = id, Data = imgData, Extension = ext, ContentType = contentType };
            }
        }

        private void ParseSlideListWithText(byte[] data, int start, int length, Presentation presentation, Dictionary<int, int> persistMap, List<int> slidePersistIds, int listInstance)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            Slide currentSlide = null;
            string lastText = null;
            List<(int start, int end, int hyperlinkId)> pendingHyperlinks = new List<(int, int, int)>();
            int? lastTextRangeStart = null;
            int? lastTextRangeEnd = null;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                switch (header.RecType)
                {
                    case RT_SlidePersistAtom:
                        if (header.RecLen >= 20)
                        {
                            int persistRef = BitConverter.ToInt32(data, atomStart);
                            int slideId = BitConverter.ToInt32(data, atomStart + 8);
                            slidePersistIds.Add(persistRef);
                            
                            currentSlide = new Slide();
                            if (listInstance == 0) // Masters
                            {
                                currentSlide.Index = presentation.Masters.Count + 1;
                                presentation.Masters.Add(currentSlide);
                                _log?.Invoke($"Adding Master Slide Index={currentSlide.Index}, persistRef={persistRef}");
                            }
                            else
                            {
                                currentSlide.Index = presentation.Slides.Count + 1;
                                presentation.Slides.Add(currentSlide);
                                _log?.Invoke($"Adding Regular Slide Index={currentSlide.Index}, persistRef={persistRef}");
                            }
                            
                            if (persistMap.TryGetValue(persistRef, out int slideOffset))
                            {
                                if (slideOffset >= 0 && slideOffset < data.Length - 8)
                                {
                                    var slideHeader = ReadRecordHeader(data, slideOffset);
                                    if (slideHeader.RecType == RT_Slide || slideHeader.RecType == RT_MainMaster)
                                    {
                                        ParseSlideContainer(data, slideOffset + 8, (int)slideHeader.RecLen, currentSlide);
                                    }
                                }
                            }
                        }
                        break;
                        
                    case RT_TextCharsAtom:
                        lastText = ReadUnicodeString(data, atomStart, (int)header.RecLen);
                        if (currentSlide != null && !string.IsNullOrWhiteSpace(lastText))
                        {
                            var paragraph = new TextParagraph();
                            paragraph.Runs.Add(new TextRun { Text = lastText });
                            currentSlide.TextContent.Add(paragraph);
                        }
                        break;
                        
                    case RT_TextBytesAtom:
                        lastText = ReadAnsiString(data, atomStart, (int)header.RecLen);
                        if (currentSlide != null && !string.IsNullOrWhiteSpace(lastText))
                        {
                            var paragraph = new TextParagraph();
                            paragraph.Runs.Add(new TextRun { Text = lastText });
                            currentSlide.TextContent.Add(paragraph);
                        }
                        break;

                    case RT_StyleTextPropAtom:
                        if (currentSlide != null && currentSlide.TextContent.Count > 0 && !string.IsNullOrEmpty(lastText))
                        {
                            var lastPara = currentSlide.TextContent.Last();
                            ParseStyleTextPropAtom(data, atomStart, (int)header.RecLen, lastPara, lastText.Length);
                            ApplyTextHyperlinks(lastPara, pendingHyperlinks);
                            pendingHyperlinks.Clear();
                        }
                        break;

                    case RT_TextInteractiveInfoAtom:
                        if (header.RecLen >= 8)
                        {
                            lastTextRangeStart = BitConverter.ToInt32(data, atomStart);
                            lastTextRangeEnd = BitConverter.ToInt32(data, atomStart + 4);
                        }
                        break;

                    case RT_InteractiveInfo:
                        if (lastTextRangeStart.HasValue && lastTextRangeEnd.HasValue)
                        {
                            // In text sequence, InteractiveInfo container usually follows TextInteractiveInfoAtom
                            // We need to look inside it for InteractiveInfoAtom
                            int iiId = ParseInteractiveInfoId(data, atomStart, (int)header.RecLen);
                            if (iiId > 0)
                            {
                                pendingHyperlinks.Add((lastTextRangeStart.Value, lastTextRangeEnd.Value, iiId));
                            }
                            lastTextRangeStart = null;
                            lastTextRangeEnd = null;
                        }
                        break;
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
        }

        private void ParseStyleTextPropAtom(byte[] data, int start, int length, TextParagraph paragraph, int textLength)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            string fullText = paragraph.GetPlainText();
            int textCharCount = textLength + 1; // +1 for the CR at the end in PPT text runs

            // === Phase 1: Parse paragraph-level runs ===
            var paraStyles = new List<(int count, TextAlignment align, ushort indentLevel, uint mask,
                                       ushort bulletFlags, ushort? bulletChar, ushort? bulletFontRef,
                                       short? lineSpacing, short? spaceBefore, short? spaceAfter,
                                       short? leftMargin, short? indent)>();
            int paraTextConsumed = 0;
            bool isFirstParaRun = true;
            while (pos < end && paraTextConsumed < textCharCount)
            {
                if (pos + 4 > end) break;
                int paraRunChars = (int)BitConverter.ToUInt32(data, pos);
                pos += 4;
                paraTextConsumed += paraRunChars;

                if (pos + 2 > end) break;
                ushort paraIndentLevel = BitConverter.ToUInt16(data, pos);
                pos += 2;

                if (pos + 4 > end) break;
                uint paraMask = BitConverter.ToUInt32(data, pos);
                pos += 4;

                TextAlignment align = TextAlignment.Left;
                ushort bulletFlags = 0;
                ushort? bulletChar = null;
                ushort? bulletFontRef = null;
                short? lineSpacing = null;
                short? spaceBefore = null;
                short? spaceAfter = null;
                short? leftMargin = null;
                short? indent = null;

                // Parse paragraph properties based on mask bits
                if ((paraMask & 0x000F) != 0) // bulletFlags
                {
                    if (pos + 2 <= end)
                    {
                        bulletFlags = BitConverter.ToUInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x0080) != 0) // bulletChar
                {
                    if (pos + 2 <= end)
                    {
                        bulletChar = BitConverter.ToUInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x0010) != 0) // bulletFontRef
                {
                    if (pos + 2 <= end)
                    {
                        bulletFontRef = BitConverter.ToUInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x0040) != 0) // bulletSize
                    pos = Math.Min(pos + 2, end);
                if ((paraMask & 0x0020) != 0) // bulletColor
                    pos = Math.Min(pos + 4, end);
                if ((paraMask & 0x0800) != 0) // textAlignment
                {
                    if (pos + 2 <= end)
                    {
                        ushort alignVal = BitConverter.ToUInt16(data, pos);
                        align = alignVal switch { 1 => TextAlignment.Center, 2 => TextAlignment.Right, 3 => TextAlignment.Justify, _ => TextAlignment.Left };
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x1000) != 0) // lineSpacing
                {
                    if (pos + 2 <= end)
                    {
                        lineSpacing = BitConverter.ToInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x2000) != 0) // spaceBefore
                {
                    if (pos + 2 <= end)
                    {
                        spaceBefore = BitConverter.ToInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x4000) != 0) // spaceAfter
                {
                    if (pos + 2 <= end)
                    {
                        spaceAfter = BitConverter.ToInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x8000) != 0) // leftMargin
                {
                    if (pos + 2 <= end)
                    {
                        leftMargin = BitConverter.ToInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x10000) != 0) // indent
                {
                    if (pos + 2 <= end)
                    {
                        indent = BitConverter.ToInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((paraMask & 0x0100) != 0) // defaultTabSize
                    pos = Math.Min(pos + 2, end);
                if ((paraMask & 0x0400) != 0) // wrapFlags
                    pos = Math.Min(pos + 2, end);
                if ((paraMask & 0x200000) != 0) // fontAlign
                    pos = Math.Min(pos + 2, end);
                // textDirection, reserved, etc. — skip any other known bits
                paraStyles.Add((paraRunChars, align, paraIndentLevel, paraMask,
                                bulletFlags, bulletChar, bulletFontRef,
                                lineSpacing, spaceBefore, spaceAfter,
                                leftMargin, indent));

                // 只把第一段样式映射到 Paragraph，对后续 runs 暂不细分
                if (isFirstParaRun)
                {
                    isFirstParaRun = false;
                    paragraph.IndentLevel = paraIndentLevel;
                    // 简单判断是否有项目符号
                    paragraph.HasBullet = bulletFlags != 0 || bulletChar.HasValue;
                    if (bulletChar.HasValue)
                    {
                        paragraph.BulletChar = (char)bulletChar.Value;
                    }
                    if (bulletFontRef.HasValue)
                    {
                        int idx = bulletFontRef.Value;
                        if (idx >= 0 && idx < _fontTable.Count)
                        {
                            paragraph.BulletFont = _fontTable[idx];
                        }
                    }
                    paragraph.LineSpacing = lineSpacing;
                    paragraph.SpaceBefore = spaceBefore;
                    paragraph.SpaceAfter = spaceAfter;
                    paragraph.LeftMargin = leftMargin;
                    paragraph.Indent = indent;
                }
            }

            // Apply first paragraph alignment
            if (paraStyles.Count > 0)
            {
                paragraph.Alignment = paraStyles[0].align;
            }

            // === Phase 2: Parse character-level runs ===
            var runs = new List<TextRun>();
            int charTextConsumed = 0;
            int textPos = 0;
            while (pos < end && charTextConsumed < textCharCount)
            {
                if (pos + 4 > end) break;
                int charRunChars = (int)BitConverter.ToUInt32(data, pos);
                pos += 4;
                charTextConsumed += charRunChars;

                if (pos + 4 > end) break;
                uint charMask = BitConverter.ToUInt32(data, pos);
                pos += 4;

                var run = new TextRun();

                // Parse character properties based on mask bits
                if ((charMask & 0xFFFF) != 0) // charFlags (bold, italic, underline, etc.)
                {
                    if (pos + 2 <= end)
                    {
                        ushort flags = BitConverter.ToUInt16(data, pos);
                        run.Bold = (flags & 0x01) != 0;
                        run.Italic = (flags & 0x02) != 0;
                        run.Underline = (flags & 0x04) != 0;
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((charMask & 0x10000) != 0) // fontRef
                {
                    if (pos + 2 <= end)
                    {
                        int fontIdx = BitConverter.ToUInt16(data, pos);
                        if (fontIdx >= 0 && fontIdx < _fontTable.Count)
                        {
                            run.FontName = _fontTable[fontIdx];
                        }
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((charMask & 0x200000) != 0) // oldEAFontRef
                    pos = Math.Min(pos + 2, end);
                if ((charMask & 0x400000) != 0) // ansiFontRef
                    pos = Math.Min(pos + 2, end);
                if ((charMask & 0x800000) != 0) // symbolFontRef
                    pos = Math.Min(pos + 2, end);
                if ((charMask & 0x20000) != 0) // fontSize
                {
                    if (pos + 2 <= end)
                    {
                        // PPT stores font size in hundredths of a point (same as OOXML a:rPr@sz).
                        // The previous implementation multiplied by 100 again, producing oversized text.
                        run.FontSize = BitConverter.ToUInt16(data, pos);
                    }
                    pos = Math.Min(pos + 2, end);
                }
                if ((charMask & 0x40000) != 0) // color
                {
                    if (pos + 4 <= end)
                    {
                        byte b = data[pos];
                        byte g = data[pos + 1];
                        byte r = data[pos + 2];
                        run.Color = $"{r:X2}{g:X2}{b:X2}";
                    }
                    pos = Math.Min(pos + 4, end);
                }
                if ((charMask & 0x80000) != 0) // position (superscript/subscript)
                    pos = Math.Min(pos + 2, end);

                // Assign text for this run
                int runTextLen = Math.Min(charRunChars, fullText.Length - textPos);
                if (runTextLen > 0)
                {
                    run.Text = fullText.Substring(textPos, runTextLen);
                    textPos += runTextLen;
                    runs.Add(run);
                }
            }

            if (runs.Count > 0)
            {
                paragraph.Runs = runs;
            }
        }

        private Shape? ParseSpContainer(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            var shape = new Shape { Type = "Rectangle" };
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                switch (header.RecType)
                {
                    case ESCHER_ChildAnchor:
                        if (header.RecLen >= 16)
                        {
                            int left = BitConverter.ToInt32(data, atomStart);
                            int top = BitConverter.ToInt32(data, atomStart + 4);
                            int right = BitConverter.ToInt32(data, atomStart + 8);
                            int bottom = BitConverter.ToInt32(data, atomStart + 12);
                            
                            // ChildAnchors are relative to the parent group.
                            // In PPT, flat shapes use master coordinates (where 1 unit = 1/8 point = 1587.5 EMUs)
                            // We do a rough conversion here assuming no scaling group.
                            shape.Left = (long)(left * 12700) / 8;
                            shape.Top = (long)(top * 12700) / 8;
                            shape.Width = (long)((right - left) * 12700) / 8;
                            shape.Height = (long)((bottom - top) * 12700) / 8;
                        }
                        break;
                        
                    case ESCHER_ClientAnchor:
                        if (header.RecLen == 8)
                        {
                            // PowerPoint specific ClientAnchor is 8 bytes (Top, Left, Right, Bottom as Int16 in 1/8 points)
                            short top = BitConverter.ToInt16(data, atomStart);
                            short left = BitConverter.ToInt16(data, atomStart + 2);
                            short right = BitConverter.ToInt16(data, atomStart + 4);
                            short bottom = BitConverter.ToInt16(data, atomStart + 6);
                            
                            shape.Left = (long)(left * 12700) / 8;
                            shape.Top = (long)(top * 12700) / 8;
                            shape.Width = (long)((right - left) * 12700) / 8;
                            shape.Height = (long)((bottom - top) * 12700) / 8;
                        }
                        else if (header.RecLen >= 16)
                        {
                            int top = BitConverter.ToInt32(data, atomStart);
                            int left = BitConverter.ToInt32(data, atomStart + 4);
                            int right = BitConverter.ToInt32(data, atomStart + 8);
                            int bottom = BitConverter.ToInt32(data, atomStart + 12);
                            
                            shape.Left = (long)(left * 12700) / 8;
                            shape.Top = (long)(top * 12700) / 8;
                            shape.Width = (long)((right - left) * 12700) / 8;
                            shape.Height = (long)((bottom - top) * 12700) / 8;
                        }
                        break;
                        
                    case ESCHER_ClientTextbox:
                        shape.Type = "TextBox";
                        ParseClientTextbox(data, atomStart, (int)header.RecLen, shape);
                        break;
                        
                    case ESCHER_Sp:
                        if (header.RecLen >= 8)
                        {
                            uint flags = BitConverter.ToUInt32(data, atomStart + 4);
                            if ((flags & 0x01) != 0) shape.Type = "Group";
                        }
                        break;

                    case ESCHER_Opt:
                        // Find blip id if this is a picture
                        ParseEscherOpt(data, atomStart, (int)header.RecLen, shape, (ushort)header.RecInstance);
                        break;
                        
                    case ESCHER_ClientData:
                        ParseInteractiveInfo(data, atomStart, (int)header.RecLen, shape);
                        // Check for OLE object reference (chart detection)
                        int exObjId = ParseExObjRefAtom(data, atomStart, (int)header.RecLen);
                        if (exObjId > 0)
                        {
                            var chart = TryParseChartFromExObjId(exObjId);
                            if (chart != null)
                            {
                                shape.Type = "Chart";
                                shape.Chart = chart;
                                _log?.Invoke($"Parsed chart from exObjId={exObjId}: {chart.Series.Count} series");
                            }
                            else
                            {
                                var emb = TryExtractEmbeddedResourceFromExObjId(exObjId);
                                if (emb != null)
                                {
                                    // Assign an ID later in a post-pass; store on shape for now via placeholder fields
                                    shape.Type = "OleObject";
                                    shape.Text = shape.Text ?? $"[Embedded object] {emb.ProgId ?? ""}".Trim();
                                    // Temporarily stash in ImageData fields to avoid further model changes here
                                    shape.ImageData = emb.Data;
                                    shape.ImageContentType = emb.ContentType;
                                }
                            }
                        }
                        
                        // Check for Programmable Tags (Table detection)
                        ParseProgTags(data, atomStart, (int)header.RecLen, shape);

                        // Check for Animation Info
                        ParseAnimationInfo(data, atomStart, (int)header.RecLen, shape);
                        break;
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
            
            if (shape.Type == "Group" && string.IsNullOrEmpty(shape.Text) && shape.Paragraphs.Count == 0 && shape.ImageId == null)
                return null;
                
            return shape;
        }

        private void ParseSlideContainer(byte[] data, int start, int length, Slide slide)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                int atomStart = pos + 8;
                
                if (header.RecType == RT_SlideAtom || header.RecType == RT_SlideMasterAtom)
                {
                    if (header.RecLen >= 12)
                    {
                        slide.SlideId = BitConverter.ToInt32(data, atomStart + 8);
                    }
                }
                else if (header.RecType == ESCHER_DgContainer)
                {
                    ParseDrawingContainer(data, pos + 8, (int)header.RecLen, slide);
                }
                else if (header.RecType == RT_SlideShowSlideInfoAtom)
                {
                    ParseSlideShowSlideInfoAtom(data, atomStart, (int)header.RecLen, slide);
                }
                else if (header.RecType == RT_ColorSchemeAtom)
                {
                    slide.ColorScheme = ParseColorSchemeAtom(data, atomStart, (int)header.RecLen);
                }
                else if (header.IsContainer)
                {
                    ParseSlideContainer(data, pos + 8, (int)header.RecLen, slide);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
        }
        
        private void ParseSlideShowSlideInfoAtom(byte[] data, int start, int length, Slide slide)
        {
            if (length < 24) return;
            
            // 0..3:  slideTime (ticks, 1 tick = 1/256s or ms depending on flags)
            // 4..7:  soundIdRef
            // 8..9:  effectDir
            // 10..11: effectType
            // 12..13: action
            // 14..15: autoAdvanceTime (ticks)
            // 16..19: flags
            // 20..21: speed
            
            int effectType = BitConverter.ToInt16(data, start + 10);
            int flags = BitConverter.ToInt32(data, start + 16);
            int speedVal = BitConverter.ToInt16(data, start + 20);
            int advanceTimeTicks = BitConverter.ToInt32(data, start);
            
            var transition = new SlideTransition();
            
            // 1. Map Speed: 0=Slow, 1=Medium, 2=Fast
            transition.Speed = speedVal switch
            {
                0 => "slow",
                1 => "med",
                _ => "fast"
            };
            
            // 2. Map Effect Type (basic PPTX mappings)
            transition.Type = effectType switch
            {
                0 => "none",
                1 => "cut",
                2 => "cutThroughBlack",
                513 => "blinds",
                769 => "checker",
                1025 => "cover",
                1281 => "dissolve",
                1537 => "fade",
                1793 => "pull",
                2049 => "randomBar",
                2305 => "strips",
                2561 => "wipe",
                2817 => "box",
                3073 => "wedge",
                3329 => "split",
                _ => (effectType >= 256 && effectType <= 512) ? "random" : "none"
            };
            
            // 3. Flags mapping
            // Bit 0x04: Has sound
            // Bit 0x10: Has auto advance
            bool hasAutoAdvance = (flags & 0x10) != 0;
            transition.HasAutoAdvance = hasAutoAdvance;
            
            if (hasAutoAdvance)
            {
                // Typically slideTime is stored as "ticks" where 1 tick = 1 millisecond in modern files / BIFF8, 
                // or 1/256 sec in older formats. For MS-PPT it's generally either milliseconds or ticks.
                // However, autoAdvanceTime in PPT is often simply Milliseconds.
                transition.AdvanceTime = advanceTimeTicks;
            }
            
            slide.Transition = transition;
            _log?.Invoke($"Parsed Slide Transition: {transition.Type}, speed={transition.Speed}, autoAdvance={transition.HasAutoAdvance} ({transition.AdvanceTime}ms)");
        }

        private void ParseAnimationInfo(byte[] data, int start, int length, Shape shape)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;

            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                int atomStart = pos + 8;

                if (header.RecType == RT_AnimationInfoContainer)
                {
                    ParseAnimationInfo(data, atomStart, (int)header.RecLen, shape);
                }
                else if (header.RecType == RT_AnimationInfoAtom && header.RecLen >= 28)
                {
                    var anim = new ShapeAnimation();
                    
                    // Flags at offset 4 (after dimColor 4 bytes)
                    // Byte 4: bits for fReverse, fAutomatic, fSound, fStopSound
                    byte flags1 = data[atomStart + 4];

                    // fAutomatic is a 2-bit field at bit index 2. When it is 0, the effect is click-triggered.
                    // Non-zero values indicate automatic sequencing (with/after previous).
                    byte fAutomatic = (byte)((flags1 >> 2) & 0x03);
                    anim.TriggerOnClick = fAutomatic == 0;

                    // orderID is at offset 8 (2 bytes)
                    anim.Order = BitConverter.ToInt16(data, atomStart + 8);

                    // animEffect is at at offset 11 (1 byte)
                    byte effect = data[atomStart + 11];
                    // animEffectDirection is at offset 12 (1 byte)
                    byte direction = data[atomStart + 12];

                    anim.Type = effect switch
                    {
                        0x01 => "fly", // Fly
                        0x0D => "fade", // Appear/Fade? 0x0D is often Appear
                        0x02 => "fly", // Fly
                        0x03 => "fly", // Fly
                        0x04 => "fly", // Fly
                        0x08 => "wipe", // Wipe
                        0x0C => "dissolve", // Dissolve
                        0x10 => "zoom", // Zoom
                        _ => "fade"
                    };

                    anim.Direction = direction switch
                    {
                        0x00 => "l", // Left
                        0x01 => "t", // Top
                        0x02 => "r", // Right
                        0x03 => "b", // Bottom
                        _ => "none"
                    };

                    shape.Animation = anim;
                    _log?.Invoke($"Parsed Animation for shape: type={anim.Type}, order={anim.Order}, triggerOnClick={anim.TriggerOnClick}");
                }

                pos = recordEnd;
                if (pos <= start) break;
            }
        }

        private void ParseProgTags(byte[] data, int start, int length, Shape shape)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;

            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                int atomStart = pos + 8;

                if (header.RecType == RT_ProgTags)
                {
                    ParseProgTags(data, atomStart, (int)header.RecLen, shape);
                }
                else if (header.RecType == RT_ProgBinaryTag)
                {
                    // ProgBinaryTag contains a CString for tagName and a BinaryTagData for data
                    string tagName = "";
                    int subPos = atomStart;
                    int subEnd = recordEnd;
                    while (subPos + 8 <= subEnd)
                    {
                        var subHeader = ReadRecordHeader(data, subPos);
                        int subAtomStart = subPos + 8;
                        if (subHeader.RecType == RT_CString)
                        {
                            tagName = Encoding.Unicode.GetString(data, subAtomStart, (int)subHeader.RecLen).TrimEnd('\0');
                        }
                        else if (subHeader.RecType == RT_BinaryTagData)
                        {
                            if (tagName == "___PPT10" || tagName == "___PPT12")
                            {
                                // This is a native table indicator
                                shape.IsNativeTable = true;
                                _log?.Invoke($"Detected Native Table via {tagName}");
                            }
                        }
                        subPos += 8 + (int)subHeader.RecLen;
                    }
                }

                pos = recordEnd;
                if (pos <= start) break;
            }
        }

        private ColorScheme? ParseColorSchemeAtom(byte[] data, int start, int length)
        {
            if (length < 32) return null;

            var scheme = new ColorScheme();
            scheme.Background = GetColorHexStr(data, start);
            scheme.TextAndLines = GetColorHexStr(data, start + 4);
            scheme.Shadows = GetColorHexStr(data, start + 8);
            scheme.TitleText = GetColorHexStr(data, start + 12);
            scheme.Fills = GetColorHexStr(data, start + 16);
            scheme.Accent = GetColorHexStr(data, start + 20);
            scheme.AccentAndHyperlink = GetColorHexStr(data, start + 24);
            scheme.AccentAndFollowingHyperlink = GetColorHexStr(data, start + 28);
            
            return scheme;
        }

        private string GetColorHexStr(byte[] data, int offset)
        {
            // PPT stores colors as an array of 4 bytes: R, G, B, _ (usually).
            // This is effectively a standard COLORREF but reversed because COLORREF is usually 0x00bbggrr,
            // while in memory here it's literally R G B _.
            if (offset + 2 < data.Length)
            {
                byte r = data[offset];
                byte g = data[offset + 1];
                byte b = data[offset + 2];
                return $"{r:X2}{g:X2}{b:X2}";
            }
            return "000000";
        }

        private void ParseNotesContainer(byte[] data, int start, int length)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            int slideIdRef = -1;
            var tempSlide = new Slide();
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                int atomStart = pos + 8;
                
                if (header.RecType == RT_NotesAtom)
                {
                    if (header.RecLen >= 4)
                    {
                        slideIdRef = BitConverter.ToInt32(data, atomStart);
                    }
                }
                else if (header.RecType == ESCHER_DgContainer)
                {
                    ParseDrawingContainer(data, pos + 8, (int)header.RecLen, tempSlide);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }
            
            if (slideIdRef > 0 && tempSlide.TextContent.Count > 0)
            {
                string notesText = string.Join("\n", tempSlide.TextContent.ConvertAll(p => p.GetPlainText()));
                if (!string.IsNullOrWhiteSpace(notesText))
                {
                    _notesIdMap[slideIdRef] = notesText;
                }
            }
        }

        private void ParseDrawingContainer(byte[] data, int start, int length, Slide slide)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;
            
            var shapesInThisContainer = new List<Shape>();
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == ESCHER_SpContainer)
                {
                    var shape = ParseSpContainer(data, pos + 8, (int)header.RecLen);
                    if (shape != null)
                    {
                        shapesInThisContainer.Add(shape);
                    }
                }
                else if (header.IsContainer)
                {
                    // For nested group containers (SpgrContainer)
                    ParseDrawingContainer(data, pos + 8, (int)header.RecLen, slide);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }

            // Post-process shapes in this container to detect tables
            var finalShapes = TryDetectTable(shapesInThisContainer);
            foreach (var shape in finalShapes)
            {
                slide.Shapes.Add(shape);
                if (!string.IsNullOrEmpty(shape.Text))
                {
                    var para = new TextParagraph();
                    para.Runs.Add(new TextRun { Text = shape.Text });
                    slide.TextContent.Add(para);
                }
            }
        }

        private List<Shape> TryDetectTable(List<Shape> shapes)
        {
            if (shapes.Count == 0) return shapes;

            // Simple table detection: If a group has many rectangles (cells)
            // that are aligned in rows and columns.
            var groupShape = shapes.FirstOrDefault(s => s.Type == "Group");
            if (groupShape == null) return shapes;

            var childShapes = shapes.Where(s => s != groupShape).ToList();
            if (childShapes.Count < 2) return shapes;

            // Group by Top position to find rows (with 1/2 point tolerance = 6350 EMUs approx)
            // Actually, PPT coordinates are often 1/8 points. 1/8 point = 1587.5 EMUs.
            // Let's use a tolerance of 2000 EMUs.
            long tolerance = 2000;
            
            var rows = childShapes
                .GroupBy(s => (long)(Math.Round((double)s.Top / tolerance) * tolerance))
                .OrderBy(g => g.Key)
                .ToList();

            var cols = childShapes
                .GroupBy(s => (long)(Math.Round((double)s.Left / tolerance) * tolerance))
                .OrderBy(g => g.Key)
                .ToList();

            // Threshold: If it's a grid (Rows * Cols approximately matches child count)
            // Or if it was explicitly marked as a native table via ProgTags
            bool isHeuristicTable = rows.Count > 1 && cols.Count > 1 && Math.Abs(rows.Count * cols.Count - childShapes.Count) <= 2;
            
            if (groupShape.IsNativeTable || isHeuristicTable)
            {
                var table = new Table();
                foreach (var rowGroup in rows)
                {
                    var row = new TableRow();
                    var sortedCells = rowGroup.OrderBy(s => s.Left).ToList();
                    
                    // Row height is the max height of cells in this row
                    row.Height = sortedCells.Count > 0 ? sortedCells.Max(c => c.Height) : 370840;

                    foreach (var cellShape in sortedCells)
                    {
                        var cell = new TableCell();
                        cell.TextContent = cellShape.Paragraphs;
                        cell.FillColor = cellShape.FillColor;
                        
                        // Accurate Layout extraction
                        cell.VerticalAlignment = cellShape.VerticalAlignment;
                        if (cellShape.MarginLeft.HasValue) cell.MarginLeft = cellShape.MarginLeft.Value;
                        if (cellShape.MarginTop.HasValue) cell.MarginTop = cellShape.MarginTop.Value;
                        if (cellShape.MarginRight.HasValue) cell.MarginRight = cellShape.MarginRight.Value;
                        if (cellShape.MarginBottom.HasValue) cell.MarginBottom = cellShape.MarginBottom.Value;
                        
                        row.Cells.Add(cell);
                    }
                    table.Rows.Add(row);
                }

                // Populate ColumnWidths based on the columns we found
                foreach (var colGroup in cols)
                {
                    var firstCellInCol = colGroup.OrderBy(c => c.Top).First();
                    table.ColumnWidths.Add(firstCellInCol.Width);
                }

                groupShape.Type = "Table";
                groupShape.Table = table;
                // Only return the Table (group) shape, discard individual cell shapes
                return new List<Shape> { groupShape };
            }

            return shapes;
        }

        private void ParseEscherOpt(byte[] data, int start, int length, Shape shape, ushort propCount)
        {
            int propTableSize = propCount * 6;
            int complexDataStart = start + propTableSize;
            int currentComplexOffset = 0;

            byte[] vData = null; int vOffset = 0, vLen = 0;
            byte[] sData = null; int sOffset = 0, sLen = 0;

            for (int i = 0; i < propCount; i++)
            {
                int entryPos = start + (i * 6);
                if (entryPos + 6 > data.Length) break;

                ushort propId = BitConverter.ToUInt16(data, entryPos);
                uint propValue = BitConverter.ToUInt32(data, entryPos + 2);

                int pid = propId & 0x3FFF;
                bool isComplex = (propId & 0x8000) != 0;

                if (isComplex)
                {
                    int dataLen = (int)propValue;
                    int dataOffset = complexDataStart + currentComplexOffset;
                    if (dataOffset + dataLen <= data.Length)
                    {
                        if (pid == 325) { vData = data; vOffset = dataOffset; vLen = dataLen; }
                        else if (pid == 326) { sData = data; sOffset = dataOffset; sLen = dataLen; }
                        else HandleComplexProperty(pid, data, dataOffset, dataLen, shape);
                    }
                    currentComplexOffset += dataLen;
                }
                else
                {
                    HandleSimpleProperty(pid, propValue, shape);
                }
            }
            
            if (vData != null || sData != null)
            {
                BuildCustomGeometry(vData, vOffset, vLen, sData, sOffset, sLen, shape);
            }
        }

        private void HandleSimpleProperty(int pid, uint propValue, Shape shape)
        {
            switch (pid)
            {
                case 0x0104: // pib (Picture Blip ID)
                    shape.ImageId = (int)propValue;
                    shape.Type = "Picture";
                    break;
                case 0x0181: // fillColor
                    byte fb = (byte)(propValue & 0xFF);
                    byte fg = (byte)((propValue >> 8) & 0xFF);
                    byte fr = (byte)((propValue >> 16) & 0xFF);
                    shape.FillColor = $"{fr:X2}{fg:X2}{fb:X2}";
                    break;
                case 0x01BF: // fNoFillHitTest
                    if ((propValue & 0x10) == 0) shape.FillColor = null;
                    break;
                case 0x01C0: // lineColor
                    byte lb = (byte)(propValue & 0xFF);
                    byte lg = (byte)((propValue >> 8) & 0xFF);
                    byte lr = (byte)((propValue >> 16) & 0xFF);
                    shape.LineColor = $"{lr:X2}{lg:X2}{lb:X2}";
                    break;
                case 0x01CB: // lineWidth
                    shape.LineWidth = propValue;
                    break;
                case 0x01CE: // lineDashing (very rough)
                    shape.LineDash = propValue switch
                    {
                        0 => "solid",
                        1 => "dot",
                        2 => "dash",
                        3 => "dashDot",
                        4 => "lgDash",
                        _ => "solid"
                    };
                    break;
                case 0x01C3: // fNoLine
                    if ((propValue & 0x08) == 0) shape.LineColor = null;
                    break;
                case 0x0180: // fillType (rough: non-zero -> treat as gradient)
                    shape.HasGradientFill = propValue != 0;
                    break;
                case 0x0182: // fillBackColor
                    byte bb = (byte)(propValue & 0xFF);
                    byte bg = (byte)((propValue >> 8) & 0xFF);
                    byte br = (byte)((propValue >> 16) & 0xFF);
                    shape.FillBackColor = $"{br:X2}{bg:X2}{bb:X2}";
                    break;
                case 0x0201: // shadowColor (best-effort)
                    byte sb = (byte)(propValue & 0xFF);
                    byte sg = (byte)((propValue >> 8) & 0xFF);
                    byte sr = (byte)((propValue >> 16) & 0xFF);
                    shape.ShadowColor = $"{sr:X2}{sg:X2}{sb:X2}";
                    shape.HasShadow = true;
                    break;
                case 0x0081: // dyTextTop
                    shape.MarginTop = (long)propValue;
                    break;
                case 0x0082: // dyTextBottom
                    shape.MarginBottom = (long)propValue;
                    break;
                case 0x0083: // dxTextLeft
                    shape.MarginLeft = (long)propValue;
                    break;
                case 0x0084: // dxTextRight
                    shape.MarginRight = (long)propValue;
                    break;
                case 0x0085: // anchorText
                    shape.VerticalAlignment = propValue switch
                    {
                        0 or 3 => "t",
                        1 or 4 => "ctr",
                        2 or 5 => "b",
                        _ => "t"
                    };
                    break;
                
                // Geometry simple props
                case 320: if (shape.Geometry == null) shape.Geometry = new ShapeGeometry(); shape.Geometry.GeoLeft = (int)propValue; break;
                case 321: if (shape.Geometry == null) shape.Geometry = new ShapeGeometry(); shape.Geometry.GeoTop = (int)propValue; break;
                case 322: if (shape.Geometry == null) shape.Geometry = new ShapeGeometry(); shape.Geometry.GeoRight = (int)propValue; break;
                case 323: if (shape.Geometry == null) shape.Geometry = new ShapeGeometry(); shape.Geometry.GeoBottom = (int)propValue; break;
            }
        }

        private void HandleComplexProperty(int pid, byte[] data, int offset, int length, Shape shape)
        {
            // Placeholder for other complex properties like pSegmentInfo if needed elsewhere
        }

        private List<byte[]> ParseIMsoArray(byte[] data, int offset, int length, int cbItemSize)
        {
            var items = new List<byte[]>();
            if (length < 6) return items;

            ushort nItems = BitConverter.ToUInt16(data, offset);
            ushort nItemsMax = BitConverter.ToUInt16(data, offset + 2);
            ushort cbItem = BitConverter.ToUInt16(data, offset + 4);

            int pos = offset + 6;
            for (int i = 0; i < nItems; i++)
            {
                if (pos + cbItem > data.Length) break;
                byte[] item = new byte[cbItem];
                Array.Copy(data, pos, item, 0, cbItem);
                items.Add(item);
                pos += cbItem;
            }
            return items;
        }

        private void BuildCustomGeometry(byte[]? vData, int vOffset, int vLen, byte[]? sData, int sOffset, int sLen, Shape shape)
        {
            if (shape.Geometry == null) shape.Geometry = new ShapeGeometry();
            
            var vertices = vData != null ? ParseIMsoArray(vData, vOffset, vLen, 8) : new List<byte[]>();
            var segments = sData != null ? ParseIMsoArray(sData, sOffset, sLen, 2) : new List<byte[]>();

            var path = new GeometryPath();
            int vIndex = 0;

            foreach (var sBytes in segments)
            {
                ushort sVal = BitConverter.ToUInt16(sBytes, 0);
                int cmd = sVal >> 13;
                int count = sVal & 0x1FFF;

                var command = new GeometryCommand();
                switch (cmd)
                {
                    case 0: command.Type = "moveTo"; break;
                    case 1: command.Type = "lnTo"; break;
                    case 2: command.Type = "cubicBezTo"; break;
                    case 3: command.Type = "close"; break;
                    case 4: command.Type = "none"; break; // End
                    default: command.Type = "lnTo"; break;
                }

                if (command.Type != "none")
                {
                    for (int j = 0; j < count; j++)
                    {
                        if (vIndex < vertices.Count)
                        {
                            var vBytes = vertices[vIndex++];
                            command.Points.Add(new GeometryPoint
                            {
                                X = BitConverter.ToInt32(vBytes, 0),
                                Y = BitConverter.ToInt32(vBytes, 4)
                            });
                        }
                    }
                    path.Commands.Add(command);
                }
            }
            
            // If segments is empty, it's a simple polygon
            if (segments.Count == 0 && vertices.Count > 0)
            {
                for (int i = 0; i < vertices.Count; i++)
                {
                    var vBytes = vertices[i];
                    path.Commands.Add(new GeometryCommand {
                        Type = (i == 0) ? "moveTo" : "lnTo",
                        Points = new List<GeometryPoint> {
                            new GeometryPoint { X = BitConverter.ToInt32(vBytes, 0), Y = BitConverter.ToInt32(vBytes, 4) }
                        }
                    });
                }
            }

            shape.Geometry.Paths.Add(path);
        }
        
        private void ParseClientTextbox(byte[] data, int start, int length, Shape shape)
        {
            int end = Math.Min(start + length, data.Length);
            int pos = start;

            string lastText = null;
            TextParagraph lastPara = null;
            List<(int start, int end, int hyperlinkId)> pendingHyperlinks = new List<(int, int, int)>();
            int? lastTextRangeStart = null;
            int? lastTextRangeEnd = null;
            
            while (pos + 8 <= end)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                int recordEnd = pos + 8 + (int)header.RecLen;
                
                if (header.RecType == RT_TextHeaderAtom)
                {
                    // We don't currently use TextHeaderAtom (text type) but keep it in the scan
                }
                else if (header.RecType == RT_TextCharsAtom)
                {
                    string text = ReadUnicodeString(data, atomStart, (int)header.RecLen);
                    if (!string.IsNullOrEmpty(text))
                    {
                        shape.Text = (shape.Text ?? "") + text;
                        var paragraph = new TextParagraph();
                        paragraph.Runs.Add(new TextRun { Text = text });
                        shape.Paragraphs.Add(paragraph);

                        lastText = text;
                        lastPara = paragraph;
                    }
                }
                else if (header.RecType == RT_TextBytesAtom)
                {
                    string text = ReadAnsiString(data, atomStart, (int)header.RecLen);
                    if (!string.IsNullOrEmpty(text))
                    {
                        shape.Text = (shape.Text ?? "") + text;
                        var paragraph = new TextParagraph();
                        paragraph.Runs.Add(new TextRun { Text = text });
                        shape.Paragraphs.Add(paragraph);

                        lastText = text;
                        lastPara = paragraph;
                    }
                }
                else if (header.RecType == RT_StyleTextPropAtom)
                {
                    if (lastPara != null && !string.IsNullOrEmpty(lastText))
                    {
                        ParseStyleTextPropAtom(data, atomStart, (int)header.RecLen, lastPara, lastText.Length);
                        ApplyTextHyperlinks(lastPara, pendingHyperlinks);
                        pendingHyperlinks.Clear();
                    }
                }
                else if (header.RecType == RT_TextInteractiveInfoAtom)
                {
                    if (header.RecLen >= 8)
                    {
                        lastTextRangeStart = BitConverter.ToInt32(data, atomStart);
                        lastTextRangeEnd = BitConverter.ToInt32(data, atomStart + 4);
                    }
                }
                else if (header.RecType == RT_InteractiveInfo)
                {
                    if (lastTextRangeStart.HasValue && lastTextRangeEnd.HasValue)
                    {
                        int iiId = ParseInteractiveInfoId(data, atomStart, (int)header.RecLen);
                        if (iiId > 0)
                        {
                            pendingHyperlinks.Add((lastTextRangeStart.Value, lastTextRangeEnd.Value, iiId));
                        }
                        lastTextRangeStart = null;
                        lastTextRangeEnd = null;
                    }
                }
                else if (header.IsContainer)
                {
                    // Recurse into any nested containers inside ClientTextbox.
                    // Note: we don't thread state into nested calls; most PPT files keep the text atoms flat.
                    ParseClientTextbox(data, atomStart, (int)header.RecLen, shape);
                }
                
                pos = recordEnd;
                if (pos <= start) break;
            }

            // If the textbox had hyperlink ranges but no StyleTextPropAtom, still apply best-effort to the last paragraph.
            if (pendingHyperlinks.Count > 0 && lastPara != null)
            {
                ApplyTextHyperlinks(lastPara, pendingHyperlinks);
            }
        }
        
        /// <summary>
        /// 直接扫描整个 PPT 流提取所有文本（备用方案）
        /// </summary>
        private void DirectScanForSlides(byte[] data, Presentation presentation)
        {
            // 第一遍: 找到所有的 SlideContainer
            var slideOffsets = new List<int>();
            int pos = 0;
            
            while (pos + 8 <= data.Length)
            {
                var header = ReadRecordHeader(data, pos);
                
                if (header.RecType == RT_Slide && header.IsContainer)
                {
                    slideOffsets.Add(pos);
                    pos += 8 + (int)header.RecLen;
                }
                else if (header.IsContainer)
                {
                    pos += 8; // 进入容器继续扫描
                }
                else if (header.RecLen > 0 && header.RecLen < data.Length)
                {
                    pos += 8 + (int)header.RecLen;
                }
                else
                {
                    pos += 8;
                }
                
                if (pos <= 0) break;
            }
            
            // 解析每个 SlideContainer
            foreach (int offset in slideOffsets)
            {
                var header = ReadRecordHeader(data, offset);
                var slide = new Slide { Index = presentation.Slides.Count + 1 };
                ParseSlideContainer(data, offset + 8, (int)header.RecLen, slide);
                presentation.Slides.Add(slide);
            }
            
            // 如果还是没有找到，做最后的文本扫描
            if (presentation.Slides.Count == 0)
            {
                var defaultSlide = new Slide { Index = 1 };
                ScanAllText(data, defaultSlide);
                if (defaultSlide.TextContent.Count > 0 || defaultSlide.Shapes.Count > 0)
                {
                    presentation.Slides.Add(defaultSlide);
                }
            }
        }
        
        /// <summary>
        /// 扫描整个 PPT 流中的所有文本记录
        /// </summary>
        private void ScanAllText(byte[] data, Slide slide)
        {
            int pos = 0;
            while (pos + 8 <= data.Length)
            {
                var header = ReadRecordHeader(data, pos);
                int atomStart = pos + 8;
                
                if (header.RecType == RT_TextCharsAtom && !header.IsContainer && header.RecLen > 0 && header.RecLen < 100000)
                {
                    if (atomStart + header.RecLen <= data.Length)
                    {
                        string text = ReadUnicodeString(data, atomStart, (int)header.RecLen);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var paragraph = new TextParagraph();
                            paragraph.Runs.Add(new TextRun { Text = text });
                            slide.TextContent.Add(paragraph);
                            
                            var shape = new Shape { Type = "TextBox", Text = text };
                            shape.Left = 457200;  // 0.5 inch
                            shape.Top = (long)(457200 + slide.Shapes.Count * 914400);  // stack vertically
                            shape.Width = 8229600;  // 9 inches
                            shape.Height = 457200;  // 0.5 inch
                            shape.Paragraphs.Add(paragraph);
                            slide.Shapes.Add(shape);
                        }
                    }
                }
                else if (header.RecType == RT_TextBytesAtom && !header.IsContainer && header.RecLen > 0 && header.RecLen < 100000)
                {
                    if (atomStart + header.RecLen <= data.Length)
                    {
                        string text = ReadAnsiString(data, atomStart, (int)header.RecLen);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var paragraph = new TextParagraph();
                            paragraph.Runs.Add(new TextRun { Text = text });
                            slide.TextContent.Add(paragraph);
                            
                            var shape = new Shape { Type = "TextBox", Text = text };
                            shape.Left = 457200;
                            shape.Top = (long)(457200 + slide.Shapes.Count * 914400);
                            shape.Width = 8229600;
                            shape.Height = 457200;
                            shape.Paragraphs.Add(paragraph);
                            slide.Shapes.Add(shape);
                        }
                    }
                }
                
                // 移到下一个记录
                if (header.RecLen > 0 && header.RecLen < data.Length && !header.IsContainer)
                {
                    pos = atomStart + (int)header.RecLen;
                }
                else if (header.IsContainer)
                {
                    pos += 8; // 进入容器内部
                }
                else
                {
                    pos += 8;
                }
                
                if (pos <= 0) break;
            }
        }
        
        #region PPT Record Header
        
        private struct RecordHeader
        {
            public int RecVer;
            public int RecInstance;
            public ushort RecType;
            public uint RecLen;
            
            public bool IsContainer => RecVer == 0x0F;
        }
        
        private RecordHeader ReadRecordHeader(byte[] data, int offset)
        {
            if (offset + 8 > data.Length)
                return new RecordHeader { RecType = 0, RecLen = 0 };
                
            ushort verInst = BitConverter.ToUInt16(data, offset);
            return new RecordHeader
            {
                RecVer = verInst & 0x0F,
                RecInstance = (verInst >> 4) & 0x0FFF,
                RecType = BitConverter.ToUInt16(data, offset + 2),
                RecLen = BitConverter.ToUInt32(data, offset + 4)
            };
        }
        
        private int ScanForRecord(byte[] data, ushort recordType)
        {
            int pos = 0;
            while (pos + 8 <= data.Length)
            {
                var header = ReadRecordHeader(data, pos);
                if (header.RecType == recordType)
                    return pos;
                    
                if (header.IsContainer)
                {
                    pos += 8; // 进入容器
                }
                else if (header.RecLen > 0 && pos + 8 + header.RecLen <= data.Length)
                {
                    pos += 8 + (int)header.RecLen;
                }
                else
                {
                    pos += 8;
                }
            }
            return -1;
        }
        
        #endregion
        
        #region Helpers
        
        private string ReadUnicodeString(byte[] data, int offset, int byteCount)
        {
            if (offset + byteCount > data.Length)
                byteCount = data.Length - offset;
            if (byteCount <= 0) return "";
            
            try
            {
                return Encoding.Unicode.GetString(data, offset, byteCount).TrimEnd('\0');
            }
            catch
            {
                return "";
            }
        }
        
        private string ReadAnsiString(byte[] data, int offset, int byteCount)
        {
            if (offset + byteCount > data.Length)
                byteCount = data.Length - offset;
            if (byteCount <= 0) return "";
            
            try
            {
                // PPT stores ANSI text using the system ANSI code page.
                // On Windows this matches the current culture's ANSI code page (e.g., 936 on zh-CN).
                int ansiCp = _options?.PptAnsiCodePageOverride ?? CultureInfo.CurrentCulture.TextInfo.ANSICodePage;
                var encoding = Encoding.GetEncoding(ansiCp);
                return encoding.GetString(data, offset, byteCount).TrimEnd('\0');
            }
            catch
            {
                // Fallbacks
                try { return Encoding.GetEncoding(1252).GetString(data, offset, byteCount).TrimEnd('\0'); }
                catch { return ""; }
            }
        }
        
        private static string TruncateText(string text, int maxLen)
        {
            if (text == null) return "";
            text = text.Replace("\r", "\\r").Replace("\n", "\\n");
            return text.Length > maxLen ? text.Substring(0, maxLen) + "..." : text;
        }
        
        private static byte[] ReadAllBytes(Stream stream)
        {
            if (stream is MemoryStream ms)
            {
                return ms.ToArray();
            }
            
            using var result = new MemoryStream();
            stream.CopyTo(result);
            return result.ToArray();
        }
        
        #endregion
        
        private VbaProject? ReadVbaProject(OleCompoundFile oleFile)
        {
            var vbaStream = oleFile.GetStream("VBA Project");
            if (vbaStream != null)
            {
                using (vbaStream)
                {
                    var vbaProject = new VbaProject();
                    vbaProject.ProjectData = ReadAllBytes(vbaStream);
                    return vbaProject;
                }
            }
            return null;
        }
        
        public void Dispose()
        {
            _stream.Dispose();
        }
    }
}
