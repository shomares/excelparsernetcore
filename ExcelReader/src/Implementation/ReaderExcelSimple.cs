using ExcelReader.src.Interfaces;
using System.IO.Compression;


namespace ExcelReader.src.Implementation
{
    public class ReaderExcelSimple : IReaderRow
    {

        private readonly ReaderFileSimple readerFileZip = new();
        private readonly StreamReadStrings readStrings = new();
        private readonly StreamReaderSheet readerSheet = new();
        private bool disposedValue;

        public async Task<IDictionary<string, object>> GetNextRowAsync(string fileName, string sheetName)
        {
            await using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            using var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);

            var indexValues = await readerFileZip.GetFileInfosAsync(zipArchive);

            var stringsFile = indexValues.FirstOrDefault(it => it.PartName == "/xl/sharedStrings.xml") ?? throw new Exception("Shared strings file not found in the Excel archive.");
            var strings = readStrings.GetStrings(stringsFile, zipArchive).ToArray();
            var firstSheetPage = indexValues.FirstOrDefault(it => it.PartName == $"/xl/worksheets/{sheetName}.xml") ?? throw new Exception("Shared strings file not found in the Excel archive.");

            if (firstSheetPage == null)
            {
                throw new Exception($"Sheet {sheetName} not found in the Excel archive.");
            }

            long index = 0;
            foreach (var row in readerSheet.ReadSheet(firstSheetPage.PartName, zipArchive))
            {
                index++;
            }

            Console.WriteLine(index);
            return new Dictionary<string, object>
            {
                { "FileInfos", indexValues }
             };
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    readStrings.Dispose();
                    readerSheet.Dispose();
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
