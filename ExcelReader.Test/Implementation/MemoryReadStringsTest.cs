using ExcelReader.src.Implementation;
using System.IO.Compression;

namespace ExcelReader.Test.Implementation
{

    public class MemoryReadStringsTest
    {
        [Test]
        public async Task MemoryReadStrings_ReadSimpleFileAsync()
        {
            var fileStream = new FileStream("Files/simple.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var memoryReadStrings = new MemoryReadStrings();
            await memoryReadStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = memoryReadStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("Hello"));
        }

        [Test]
        public async Task MemoryReadStrings_ReadSimpleWithSpecialCharactersFileAsync()
        {
            var fileStream = new FileStream("Files/simpleSpecial.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var memoryReadStrings = new MemoryReadStrings();
            await memoryReadStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = memoryReadStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("'Hello"));
        }

        [Test]
        public async Task MemoryReadStrings_ReadEmptyAsync()
        {
            var fileStream = new FileStream("Files/empty.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var memoryReadStrings = new MemoryReadStrings();
            await memoryReadStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => memoryReadStrings.GetStringByIndex(0));
        }

        [Test]
        public async Task MemoryReadStrings_ReadOnlyNumbersAsync()
        {
            var fileStream = new FileStream("Files/numbers.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var memoryReadStrings = new MemoryReadStrings();
            await memoryReadStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => memoryReadStrings.GetStringByIndex(0));
        }

        [Test]
        public async Task MemoryReadStrings_BigStringAsync()
        {
            var fileStream = new FileStream("Files/bigstring.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var memoryReadStrings = new MemoryReadStrings();
            await memoryReadStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = memoryReadStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("Hello"));
        }

        [Test]
        public async Task MemoryReadStrings_BigFileAsync()
        {
            var fileStream = new FileStream("Files/bigfile.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var memoryReadStrings = new MemoryReadStrings();
            await memoryReadStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = memoryReadStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("id"));
        }
    }
}
