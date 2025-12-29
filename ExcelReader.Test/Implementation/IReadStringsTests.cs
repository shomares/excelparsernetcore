

using ExcelReader.src.Implementation;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Xml;

namespace ExcelReader.Test.Implementation
{
    internal class IReadStringsTests
    {
        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
        public async Task MemoryReadStrings_ReadSimpleFileAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/simple.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;
            await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = readStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("Hello"));
        }

        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
        public async Task MemoryReadStrings_ReadSimpleWithSpecialCharactersFileAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/simpleSpecial.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;

            await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = readStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("'Hello"));
        }


        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
       
        public async Task MemoryReadStrings_ReadEmptyAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/empty.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;
            await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => readStrings.GetStringByIndex(0));
        }

        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
        public async Task MemoryReadStrings_ReadOnlyNumbersAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/numbers.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;
            await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => readStrings.GetStringByIndex(0));
        }

        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
        public async Task MemoryReadStrings_BigStringAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/bigstring.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;
            await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = readStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("Hello"));
        }

        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
        public async Task MemoryReadStrings_BigFileAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/bigfile.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;
            await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive);

            // Assert
            var value = readStrings.GetStringByIndex(0);
            Assert.That(value, Is.EqualTo("id"));
        }

        [TestCase(typeof(MemoryReadStrings))]
        [TestCase(typeof(StreamReadStrings))]
        public void MemoryReadStrings_BadAsync(Type readStringsType)
        {
            var fileStream = new FileStream("Files/corrupted.xlsx", FileMode.Open, FileAccess.Read);
            var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);
            // Arrange & Act
            using var readStrings = (IReadStrings)Activator.CreateInstance(readStringsType)!;

            Assert.ThrowsAsync<XmlException>(async () => await readStrings.LoadInfoAsync(new src.Entity.FileInfoExcel
            {
                PartName = "xl/sharedStrings.xml",

            }, zipArchive));
        }
    }
}
