using ExcelReader.src.Config;
using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;


namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Reader Excel implementation, for big files, this process the file in a streaming way
    /// </summary>
    public class ReaderExcelSimple : IReaderRow
    {

        private readonly ReadStringsFactory readStringsFactory = new();
        private IReadStrings? readStrings;
        private readonly StreamReaderSheet readerSheet = new();
        private bool disposedValue;
        private FileInfoExcel? firstSheetPage;
        private ZipArchive? zipArchive;
        private FileStream? fileStream;
        private ReaderStyles? readerStyles;
        private ConfigurationReader? _configuration = null;

        public IEnumerable<dynamic> GetNextRow(string sheetName)
        {
            if (_configuration == null)
            {
                throw new InvalidOperationException("Configuration is not initialized. You must call ReadFileAsync before reading rows.");
            }

            if (zipArchive == null)
            {
                throw new InvalidOperationException("You must call ReadFileAsync before reading rows.");
            }

            if (readStrings == null)
            {
                throw new InvalidOperationException("ReadStrings is not initialized. You must call ReadFileAsync before reading rows.");
            }

            if (readerStyles == null)
            {
                throw new InvalidOperationException("ReaderStyles is not initialized. You must call ReadFileAsync before reading rows.");
            }


            var indexValues = (zipArchive.Entries.First(it => it.FullName == $"xl/worksheets/{sheetName}.xml") ?? throw new Exception($"Sheet {sheetName} not found in the Excel archive.")) ?? throw new Exception($"Sheet {sheetName} not found in the Excel archive.");
            firstSheetPage = new FileInfoExcel
            {
                PartName = indexValues.FullName
            };

            long index = 0;

            var readerRow = new ReaderLineSimple(readStrings, readerStyles);
            var values = readerSheet.ReadSheet(firstSheetPage.PartName, zipArchive);
            IDictionary<string, string> columns = new Dictionary<string, string>();
            foreach (var row in values)
            {
                if (index == 0 && _configuration.HasHeaders)
                {
                    columns = readerRow.ReadColumns(row);
                    index++;
                    continue;
                }


                var toYield = readerRow.ReadRow(row, columns);
                yield return toYield;
                index++;
            }

        }

        public async Task ReadFileAsync(string fileName, ConfigurationReader? configuration = null)
        {
            fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read);

            readerStyles = new ReaderStyles();
            _configuration = configuration ?? new ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = false
            };

            readStrings = readStringsFactory.CreateReadStrings(fileName, configuration);
            await readStrings.LoadInfoAsync(new FileInfoExcel
            {
                PartName = "/xl/sharedStrings.xml"
            }, zipArchive);


            readerStyles.LoadInfo(zipArchive);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {

                    readerStyles?.Dispose();
                    readStrings?.Dispose();
                    readerSheet.Dispose();
                    zipArchive?.Dispose();
                    fileStream?.Dispose();
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
