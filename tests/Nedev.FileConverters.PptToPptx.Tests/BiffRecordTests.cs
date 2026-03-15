using System.IO;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class BiffRecordTests
    {
        [Fact]
        public void Read_ValidRecord_ParsesCorrectly()
        {
            using var stream = new MemoryStream();
            using var writer = new BinaryWriter(stream);

            writer.Write((ushort)0x0001);
            writer.Write((ushort)4);
            writer.Write(new byte[] { 0x01, 0x02, 0x03, 0x04 });
            writer.Flush();
            stream.Position = 0;

            using var reader = new BinaryReader(stream);
            var record = BiffRecord.Read(reader);

            Assert.Equal((ushort)0x0001, record.Id);
            Assert.Equal((ushort)4, record.Length);
            Assert.Equal(4, record.Data.Length);
            Assert.Equal(new byte[] { 0x01, 0x02, 0x03, 0x04 }, record.Data);
        }

        [Fact]
        public void GetAllData_NoContinues_ReturnsDataOnly()
        {
            var record = new BiffRecord
            {
                Id = 1,
                Length = 4,
                Data = new byte[] { 0x01, 0x02, 0x03, 0x04 }
            };

            var result = record.GetAllData();

            Assert.Equal(record.Data, result);
        }

        [Fact]
        public void GetAllData_WithContinues_ReturnsConcatenatedData()
        {
            var record = new BiffRecord
            {
                Id = 1,
                Length = 4,
                Data = new byte[] { 0x01, 0x02 }
            };
            record.Continues.Add(new byte[] { 0x03, 0x04 });
            record.Continues.Add(new byte[] { 0x05, 0x06 });

            var result = record.GetAllData();

            Assert.Equal(new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05, 0x06 }, result);
        }

        [Fact]
        public void GetAllData_NullData_ReturnsEmptyArray()
        {
            var record = new BiffRecord
            {
                Id = 1,
                Data = null
            };

            var result = record.GetAllData();

            Assert.NotNull(result);
            Assert.Empty(result);
        }
    }
}
