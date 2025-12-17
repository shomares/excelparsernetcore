using ExcelReader.src.Interfaces;
using System.IO.Compression;


namespace ExcelReader.src.Implementation
{
    public class ReaderExcelSimple : IReaderRow
    {

        private readonly ReaderFileSimple readerFileZip = new ReaderFileSimple();
        private readonly ReadStrings readStrings = new ReadStrings();
        private readonly ReaderSheet readerSheet = new ReaderSheet();


        public async Task<IDictionary<string, object>> GetNextRowAsync(string fileName)
        {
            await using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            using var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);

            var indexValues = await readerFileZip.GetFileInfosAsync(zipArchive);

            var stringsFile = indexValues.FirstOrDefault(it => it.PartName == "/xl/sharedStrings.xml") ?? throw new Exception("Shared strings file not found in the Excel archive.");
            var strings = await readStrings.ReadStringsAsync(stringsFile, zipArchive);


            var firstSheetPage = indexValues.FirstOrDefault(it => it.PartName == "/xl/worksheets/sheet1.xml") ?? throw new Exception("Shared strings file not found in the Excel archive.");

           

            foreach (var row in await readerSheet.ReadSheetAsync(firstSheetPage.PartName, strings, zipArchive))
            {
                Console.WriteLine(row.R);
            }


            return new Dictionary<string, object>
            {
                { "FileInfos", indexValues }
             };
        }
    }
}
