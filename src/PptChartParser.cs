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

            var records = new List<BiffRecord>();
            BiffRecord lastRecord = null;
            
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
                            
                        case CH_CHARTFORMAT:
                            // Not fully detailed, but can extract type hints or rely on default
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
