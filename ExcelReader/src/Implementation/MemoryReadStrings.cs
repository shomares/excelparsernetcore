using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Text;

namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Summary read string implementation
    /// </summary>
    internal class MemoryReadStrings : IReadStrings
    {
        private Stream? entryStream;
        private StreamReader? streamReader;
        private bool disposedValue;
        private List<string>? values = [];



        /// <summary>
        /// Dispose pattern implementation
        /// </summary>
        /// <param name="disposing">Indicate if resources is disposings</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    entryStream?.Dispose();
                    streamReader?.Dispose();
                    values?.Clear();
                    values = null;
                }

                disposedValue = true;
            }
        }

        /// <summary>
        /// Release pattern implementation
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public void LoadInfo(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == fileInfoExcel.PartName?.TrimStart('/'));

            if (entry == null)
            {
                return;
            }

            var response = new List<string>();
            entryStream = entry.Open();
            streamReader = new StreamReader(entryStream);
            values = [];
            var currentElement = new StringBuilder();

            while (!streamReader.EndOfStream)
            {
                var c = (char)streamReader.Read();

                if (c == '<')
                {
                    while (!streamReader.EndOfStream && c != ' ' && c != '>')
                    {
                        c = (char)streamReader.Read();

                        if (c != '>' && c != ' ')
                        {
                            currentElement.Append(c);
                        }
                    }

                    var currentElementStr = currentElement.ToString();

                    while (!streamReader.EndOfStream && c != '>')
                    {
                        c = (char)streamReader.Read();
                    }

                    //Read t properties
                    if (currentElementStr == "t")
                    {
                        var value = ReaderValues.ReadValue(streamReader);
                        values?.Add(value);
                    }


                    currentElement.Clear();
                }
            }
        }

        public string GetStringByIndex(int index)
        {
            return values != null ? values[index] : string.Empty;
        }
    }
}
