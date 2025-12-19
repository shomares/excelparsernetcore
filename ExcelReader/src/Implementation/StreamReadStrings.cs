using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Text;

namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Summary read string implementation
    /// </summary>
    internal class StreamReadStrings : IReadStrings
    {
        private Stream? entryStream;
        private StreamReader? streamReader;
        private bool disposedValue;

        /// <summary>
        /// Read string values from sharedStrings.xml file, using streaming in order to reduce memory consumption
        /// </summary>
        /// <param name="fileInfoExcel">File info excel</param>
        /// <param name="zipArchive">Zip file archive</param>
        /// <returns>Collection string, string by string</returns>
        public IEnumerable<string> GetStrings(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == fileInfoExcel.PartName?.TrimStart('/'));

            if (entry == null)
            {
                return [];
            }

            var response = new List<string>();
            entryStream = entry.Open();
            streamReader = new StreamReader(entryStream);
            return ParseContent(streamReader);
        }


        /// <summary>
        /// Parse content of sharedStrings.xml file, reading it as a stream
        /// </summary>
        /// <param name="bytes">Reading it as stream</param>
        /// <returns>A collection string, strings by string</returns>
        private static IEnumerable<string> ParseContent(StreamReader bytes)
        {
            var currentElement = new StringBuilder();

            while (!bytes.EndOfStream)
            {
                var c = (char)bytes.Read();

                if (c == '<')
                {
                    while (!bytes.EndOfStream && c != ' ' && c != '>')
                    {
                        c = (char)bytes.Read();

                        if (c != '>' && c != ' ')
                        {
                            currentElement.Append(c);
                        }
                    }

                    var currentElementStr = currentElement.ToString();

                    //Read t properties
                    if (currentElementStr == "t")
                    {
                        var value = ReaderValues.ReadValue(bytes);
                        yield return value;
                    }
                    else
                    {
                        while (!bytes.EndOfStream && c != '>')
                        {
                            c = (char)bytes.Read();
                        }
                    }

                    currentElement.Clear();
                }
            }

        }

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
    }
}
