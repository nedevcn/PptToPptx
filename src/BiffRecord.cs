using System;
using System.IO;

namespace Nefdev.PptToPptx
{
    public class BiffRecord
    {
        public ushort Id { get; set; }
        public ushort Length { get; set; }
        public byte[] Data { get; set; }
        
        /// <summary>
        /// Store all subsequent CONTINUE (0x003C) records
        /// </summary>
        public System.Collections.Generic.List<byte[]> Continues { get; set; } = new System.Collections.Generic.List<byte[]>();

        public byte[] GetAllData()
        {
            if (Data == null) return Array.Empty<byte>();
            if (Continues.Count == 0) return Data;

            int totalLength = Data.Length;
            foreach (var chunk in Continues)
            {
                totalLength += chunk.Length;
            }

            byte[] fullData = new byte[totalLength];
            Array.Copy(Data, 0, fullData, 0, Data.Length);
            
            int offset = Data.Length;
            foreach (var chunk in Continues)
            {
                Array.Copy(chunk, 0, fullData, offset, chunk.Length);
                offset += chunk.Length;
            }
            
            return fullData;
        }

        public static BiffRecord Read(BinaryReader reader)
        {
            var record = new BiffRecord();
            record.Id = reader.ReadUInt16();
            record.Length = reader.ReadUInt16();
            record.Data = reader.ReadBytes(record.Length);
            return record;
        }
    }
}
