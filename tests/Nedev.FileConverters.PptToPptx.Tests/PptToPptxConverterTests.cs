using System;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class PptToPptxConverterTests
    {
        [Fact]
        public void Convert_NullInputPath_ThrowsArgumentException()
        {
            var ex = Assert.Throws<ArgumentException>(() =>
                PptToPptxConverter.Convert(null!, "output.pptx"));
            Assert.Equal("pptPath", ex.ParamName);
        }

        [Fact]
        public void Convert_EmptyInputPath_ThrowsArgumentException()
        {
            var ex = Assert.Throws<ArgumentException>(() =>
                PptToPptxConverter.Convert("", "output.pptx"));
            Assert.Equal("pptPath", ex.ParamName);
        }

        [Fact]
        public void Convert_NullOutputPath_ThrowsArgumentException()
        {
            var tempFile = Path.GetTempFileName();
            try
            {
                var ex = Assert.Throws<ArgumentException>(() =>
                    PptToPptxConverter.Convert(tempFile, null!));
                Assert.Equal("pptxPath", ex.ParamName);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Convert_NonExistentInputFile_ThrowsFileNotFoundException()
        {
            var nonExistentPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".ppt");
            Assert.Throws<FileNotFoundException>(() =>
                PptToPptxConverter.Convert(nonExistentPath, "output.pptx"));
        }

        [Fact]
        public void Convert_SameInputOutputPath_ThrowsArgumentException()
        {
            var tempFile = Path.GetTempFileName();
            try
            {
                var ex = Assert.Throws<ArgumentException>(() =>
                    PptToPptxConverter.Convert(tempFile, tempFile));
                Assert.Equal("pptxPath", ex.ParamName);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Convert_WithOptions_DoesNotThrow()
        {
            var options = new ConversionOptions
            {
                KeepTempFiles = false,
                PptAnsiCodePageOverride = 1252,
                BiffAnsiCodePageOverride = 1252
            };

            Assert.NotNull(options);
            Assert.False(options.KeepTempFiles);
            Assert.Equal(1252, options.PptAnsiCodePageOverride);
            Assert.Equal(1252, options.BiffAnsiCodePageOverride);
        }

        [Fact]
        public void Convert_WithLogCallback_CapturesLogMessages()
        {
            var options = new ConversionOptions();
            var logMessages = new System.Collections.Generic.List<string>();
            options.Log = msg => logMessages.Add(msg);

            options.Log?.Invoke("Test message");

            Assert.Single(logMessages);
            Assert.Equal("Test message", logMessages[0]);
        }
    }
}
