using System;
using System.IO;
using System.Linq;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class OleCompoundFileTests
    {
        [Fact]
        public void Constructor_NullStream_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => new OleCompoundFile(null));
        }

        [Fact]
        public void Constructor_InvalidHeader_ThrowsInvalidDataException()
        {
            var invalidData = new byte[512];
            new Random().NextBytes(invalidData);
            using var stream = new MemoryStream(invalidData);

            Assert.Throws<InvalidDataException>(() => new OleCompoundFile(stream));
        }

        [Fact]
        public void Constructor_ValidHeader_DoesNotThrow()
        {
            var validHeader = new byte[512];
            byte[] magic = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            Array.Copy(magic, validHeader, magic.Length);

            using var stream = new MemoryStream(validHeader);
            var ex = Record.Exception(() => new OleCompoundFile(stream));
            Assert.Null(ex);
        }
    }
}
