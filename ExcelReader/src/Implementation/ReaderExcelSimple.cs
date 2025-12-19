using ExcelReader.src.Interfaces;
using System.IO.Compression;


namespace ExcelReader.src.Implementation
{
    public class ReaderExcelSimple : IReaderRow
    {

        private readonly ReaderFileSimple readerFileZip = new();
        private readonly ReadStrings readStrings = new();


        public async Task<IDictionary<string, object>> GetNextRowAsync(string fileName, string sheetName)
        {
            await using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            using var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);

            var indexValues = await readerFileZip.GetFileInfosAsync(zipArchive);

            var stringsFile = indexValues.FirstOrDefault(it => it.PartName == "/xl/sharedStrings.xml") ?? throw new Exception("Shared strings file not found in the Excel archive.");
            var strings = await readStrings.ReadStringsAsync(stringsFile, zipArchive);
            using ReaderSheet readerSheet = new();
            var firstSheetPage = indexValues.FirstOrDefault(it => it.PartName == $"/xl/worksheets/{sheetName}.xml") ?? throw new Exception("Shared strings file not found in the Excel archive.");

            if (firstSheetPage == null)
            {
                throw new Exception($"Sheet {sheetName} not found in the Excel archive.");
            }

            foreach (var row in readerSheet.ReadSheet(firstSheetPage.PartName, zipArchive))
            {
                Console.WriteLine(row.Parameters);
            }


            return new Dictionary<string, object>
            {
                { "FileInfos", indexValues }
             };
        }
    }
}
