using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace Nefdev.PptToPptx
{
    public class PptChartParser
    {
        private const ushort BIFF_EOF = 0x000A;
        private const ushort BIFF_BOF = 0x0809;
        // Chart specific records
        private const ushort CH_CHART = 0x1002;
        private const ushort CH_SERIES = 0x1003;
        private const ushort CH_DATAFORMAT = 0x1006;
        private const ushort CH_CHARTFORMAT = 0x1014;
        private const ushort CH_SERIESTEXT = 0x100D;
        
        // Data records
        private const ushort NUMBER = 0x0203;
        private const ushort LABEL = 0x0204;
        private const ushort LABELSST = 0x00FD;
        private const ushort SST = 0x00FC;

        // Common format records
        private const ushort FORMAT = 0x041E;

        // Chart sub-stream records for types and titles
        private const ushort CH_BAR = 0x1017;
        private const ushort CH_LINE = 0x1018;
        private const ushort CH_PIE = 0x1019;
        private const ushort CH_AREA = 0x101A;
        private const ushort CH_SCATTER = 0x101B;
        private const ushort CH_RADAR = 0x101C;
        private const ushort CH_TEXT = 0x1025;
        private const ushort CH_BEGIN = 0x1033;
        private const ushort CH_END = 0x1034;
        private const ushort CH_DEFAULTTEXT = 0x1024;
        private const ushort CH_DATATABLE = 0x1032;
        private const ushort CH_FRAME = 0x1033;
        private const ushort CH_AXIS = 0x101D;
        private const ushort CH_TICK = 0x101E;
        private const ushort CH_VALUERANGE = 0x101F;
        private const ushort CH_LEGEND = 0x1015;
        private const ushort CH_AREAFORMAT = 0x100A;
        private const ushort CH_LINEFORMAT = 0x1007;
        private const ushort CH_MARKERFORMAT = 0x1022;
        
        public Chart ParseChart(byte[] biffData)
        {
            var chart = new Chart();
            chart.Type = "bar"; // Default

            using var stream = new MemoryStream(biffData);
            using var reader = new BinaryReader(stream);

            var sstStrings = new List<string>();
            var sstOffsets = new List<uint>();

            ChartSeries currentSeries = null;
            
            // To align categories and values from Sheet data (often in _123456 Workbook streams)
            // Legacy MS Graph dumps data as simply:
            // Column 0 = Categories
            // Column 1 = Series 1 Values
            // Column 2 = Series 2 Values
            // Row 0 = Series Names
            // Row 1..N = Data values
            
            var cells = new Dictionary<(int row, int col), string>();
            var numbers = new Dictionary<(int row, int col), double>();
            var seriesColors = new Dictionary<int, string>();
            var seriesMarkers = new Dictionary<int, string>();

            var records = new List<BiffRecord>();
            BiffRecord lastRecord = null;
            byte lastTextType = 0; // 1=Title, 2=Category, 3=Value
            int currentSeriesFormattingIndex = -1;
            
            while (stream.Position < stream.Length)
            {
                if (stream.Position + 4 > stream.Length) break;
                
                var record = BiffRecord.Read(reader);
                if (record.Id == 0x003C && lastRecord != null) // CONTINUE
                {
                    lastRecord.Continues.Add(record.Data);
                }
                else
                {
                    records.Add(record);
                    lastRecord = record;
                }
            }

            foreach (var record in records)
            {
                try
                {
                    switch (record.Id)
                    {
                        case FORMAT:
                            // We could parse formats if needed
                            break;
                            
                        case SST:
                            ParseSstInfo(record, sstStrings, sstOffsets);
                            break;
                            
                        case NUMBER:
                            {
                                using var recStream = new MemoryStream(record.Data);
                                using var recReader = new BinaryReader(recStream);
                                ushort row = recReader.ReadUInt16();
                                ushort col = recReader.ReadUInt16();
                                ushort xf = recReader.ReadUInt16();
                                double val = recReader.ReadDouble();
                                numbers[(row, col)] = val;
                                cells[(row, col)] = val.ToString();
                            }
                            break;
                            
                        case LABEL:
                            {
                                var stringReader = new BiffStringReader(record, 6); // Skip row, col, xf
                                using var recStream = new MemoryStream(record.Data);
                                using var recReader = new BinaryReader(recStream);
                                ushort row = recReader.ReadUInt16();
                                ushort col = recReader.ReadUInt16();
                                
                                string text = stringReader.ReadString();
                                cells[(row, col)] = text;
                            }
                            break;
                            
                        case LABELSST:
                            {
                                using var recStream = new MemoryStream(record.Data);
                                using var recReader = new BinaryReader(recStream);
                                ushort row = recReader.ReadUInt16();
                                ushort col = recReader.ReadUInt16();
                                ushort xf = recReader.ReadUInt16();
                                uint sstIndex = recReader.ReadUInt32();
                                if (sstIndex < sstStrings.Count)
                                {
                                    cells[(row, col)] = sstStrings[(int)sstIndex];
                                }
                            }
                            break;
                            
                        case CH_SERIES:
                            currentSeriesFormattingIndex++;
                            break;

                        case CH_AREAFORMAT:
                            // Area format for the current series (bars, areas)
                            if (record.Data.Length >= 16 && currentSeriesFormattingIndex >= 0)
                            {
                                int rgbFore = BitConverter.ToInt32(record.Data, 0);
                                byte r = (byte)(rgbFore & 0xFF);
                                byte g = (byte)((rgbFore >> 8) & 0xFF);
                                byte b = (byte)((rgbFore >> 16) & 0xFF);
                                seriesColors[currentSeriesFormattingIndex] = $"{r:X2}{g:X2}{b:X2}";
                            }
                            break;

                        case CH_LINEFORMAT:
                            // Line format for lines or borders
                            if (record.Data.Length >= 12 && currentSeriesFormattingIndex >= 0)
                            {
                                // rgb is at offset 0 (4 bytes)
                                int rgb = BitConverter.ToInt32(record.Data, 0);
                                byte r = (byte)(rgb & 0xFF);
                                byte g = (byte)((rgb >> 8) & 0xFF);
                                byte b = (byte)((rgb >> 16) & 0xFF);
                                // For line charts, this is the primary color
                                if (!seriesColors.ContainsKey(currentSeriesFormattingIndex))
                                    seriesColors[currentSeriesFormattingIndex] = $"{r:X2}{g:X2}{b:X2}";
                            }
                            break;

                        case CH_MARKERFORMAT:
                            if (record.Data.Length >= 20 && currentSeriesFormattingIndex >= 0)
                            {
                                ushort markerType = BitConverter.ToUInt16(record.Data, 12);
                                seriesMarkers[currentSeriesFormattingIndex] = markerType switch {
                                    1 => "square",
                                    2 => "diamond",
                                    3 => "triangle",
                                    4 => "x",
                                    5 => "star",
                                    6 => "dot",
                                    7 => "dash",
                                    8 => "circle",
                                    9 => "plus",
                                    _ => "none"
                                };
                            }
                            break;

                        case CH_CHARTFORMAT:
                            // Not fully detailed, but can extract type hints or rely on default
                            break;

                        case CH_BAR:
                            chart.Type = "bar";
                            break;
                        case CH_LINE:
                            chart.Type = "line";
                            break;
                        case CH_PIE:
                            chart.Type = "pie";
                            break;
                        case CH_AREA:
                            chart.Type = "area";
                            break;
                        case CH_SCATTER:
                            chart.Type = "scatter";
                            break;
                        case CH_RADAR:
                            chart.Type = "radar";
                            break;

                        case CH_TEXT:
                            // Text object header. if iType == 1 (Title), iType == 2 (Category), iType == 3 (Value)
                            if (record.Data.Length > 0)
                                lastTextType = record.Data[0];
                            break;

                        case CH_LEGEND:
                            chart.ShowLegend = true;
                            if (record.Data.Length >= 18)
                            {
                                ushort wCheat = BitConverter.ToUInt16(record.Data, 16);
                                chart.LegendPosition = wCheat switch {
                                    0 => "b",
                                    1 => "tr",
                                    2 => "t",
                                    3 => "r",
                                    4 => "l",
                                    _ => "r"
                                };
                            }
                            break;

                        case CH_SERIESTEXT:
                            {
                                // If this appears without a preceding SERIES record, it's often the chart title
                                // In MS-GRL, if Text(1025h).iType == 1, then this is the title.
                                using var recStream = new MemoryStream(record.Data);
                                using var recReader = new BinaryReader(recStream);
                                
                                if (record.Data.Length > 2)
                                {
                                    byte cch = record.Data[2];
                                    if (record.Data.Length > 3)
                                    {
                                        byte flags = record.Data[3];
                                        bool isUni = (flags & 0x01) != 0;
                                        string text = "";
                                        if (isUni)
                                            text = Encoding.Unicode.GetString(record.Data, 4, Math.Min(cch * 2, record.Data.Length - 4));
                                        else
                                            text = Encoding.GetEncoding(1252).GetString(record.Data, 4, Math.Min(cch, record.Data.Length - 4));
                                        
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            if (lastTextType == 1 || (lastTextType == 0 && string.IsNullOrEmpty(chart.Title)))
                                                chart.Title = text;
                                            else if (lastTextType == 2)
                                                chart.CategoryAxisTitle = text;
                                            else if (lastTextType == 3)
                                                chart.ValueAxisTitle = text;
                                        }
                                    }
                                }
                            }
                            break;
                    }
                }
                catch
                {
                    // Ignore malformed records
                }
            }

            // After parsing all cells, assemble the Chart
            // Row 0, Col 1..MaxCol -> Series Names
            // Row 1..MaxRow, Col 0 -> Categories
            // Row 1..MaxRow, Col 1..MaxCol -> Values
            
            int maxRow = -1;
            int maxCol = -1;
            foreach (var key in cells.Keys)
            {
                maxRow = Math.Max(maxRow, key.row);
                maxCol = Math.Max(maxCol, key.col);
            }

            if (maxRow >= 0 && maxCol >= 0)
            {
                // Find categories
                var categoryList = new List<string>();
                for (int r = 1; r <= maxRow; r++)
                {
                    if (cells.TryGetValue((r, 0), out string cat))
                        categoryList.Add(cat);
                    else
                        categoryList.Add("");
                }

                // Build series
                for (int c = 1; c <= maxCol; c++)
                {
                    var series = new ChartSeries();
                    
                    if (cells.TryGetValue((0, c), out string sName))
                        series.Name = sName;
                    else
                        series.Name = $"Series {c}";
                        
                    series.Categories = new List<string>(categoryList);
                    
                    for (int r = 1; r <= maxRow; r++)
                    {
                        if (numbers.TryGetValue((r, c), out double num))
                            series.Values.Add(num);
                        else
                            series.Values.Add(0.0);
                    }
                    
                    if (series.Values.Count > 0)
                    {
                        // Apply formatting (series index starts at 0)
                        int seriesIdx = c - 1;
                        if (seriesColors.TryGetValue(seriesIdx, out string color))
                            series.Color = color;
                        if (seriesMarkers.TryGetValue(seriesIdx, out string marker))
                            series.MarkerType = marker;

                        chart.Series.Add(series);
                    }
                }
            }

            // Fallback if empty (shouldn't happen for valid MS Graph)
            if (chart.Series.Count == 0)
            {
                var dummySeries = new ChartSeries { Name = "Series 1" };
                dummySeries.Categories.Add("Category 1");
                dummySeries.Values.Add(1.0);
                chart.Series.Add(dummySeries);
            }

            return chart;
        }

        private void ParseSstInfo(BiffRecord record, List<string> strings, List<uint> offsets)
        {
            if (record.Data == null || record.Data.Length < 8) return;
            
            using var stream = new MemoryStream(record.Data);
            using var reader = new BinaryReader(stream);
            
            uint totalStrings = reader.ReadUInt32();
            uint uniqueStrings = reader.ReadUInt32();
            
            var stringReader = new BiffStringReader(record, 8); // SST Header size is 8 bytes
            
            for (int i = 0; i < uniqueStrings; i++)
            {
                try
                {
                    string text = stringReader.ReadString();
                    strings.Add(text);
                }
                catch
                {
                    // If parsing a string fails, try to salvage
                    if (strings.Count == i) strings.Add("");
                }
            }
        }
    }
}
