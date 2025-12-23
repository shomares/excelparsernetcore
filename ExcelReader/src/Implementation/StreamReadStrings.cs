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
        private FileStream? entryStream;
        private List<long>? positions;
        private string? temporaryFile;
        private bool disposedValue;

        /// <summary>
        /// Read string values from sharedStrings.xml file, using streaming in order to reduce memory consumption
        /// </summary>
        /// <param name="fileInfoExcel">File info excel</param>
        /// <param name="zipArchive">Zip file archive</param>
        /// <returns>Collection string, string by string</returns>
        public void LoadInfo(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == fileInfoExcel.PartName?.TrimStart('/'));

            if (entry == null)
            {
                return;
            }

            temporaryFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            entry.ExtractToFile(temporaryFile);

            entryStream = File.OpenRead(temporaryFile);
            ParseContent(entryStream);
        }


        /// <summary>
        /// Parse content of sharedStrings.xml file, reading it as a stream
        /// </summary>
        /// <param name="bytes">Reading it as stream</param>
        /// <returns>A collection string, strings by string</returns>
        private void ParseContent(FileStream fs)
        {
            positions = [];
            var currentElement = new StringBuilder();
            while (fs.Position < fs.Length)
            {
                var c = fs.ReadChar();

                if (c == '<')
                {
                    while (fs.Position < fs.Length && c != ' ' && c != '>')
                    {
                        c = fs.ReadChar();

                        if (c != '>' && c != ' ')
                        {
                            currentElement.Append(c);
                        }
                    }

                    var currentElementStr = currentElement.ToString();

                    if (currentElementStr == "t")
                    {
                        positions.Add(fs.Position);
                    }

                    while (fs.Position < fs.Length && c != '>')
                    {
                        c = fs.ReadChar();
                    }

                    //Read t properties
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
                    positions?.Clear();

                    if (temporaryFile != null && File.Exists(temporaryFile))
                    {
                        File.Delete(temporaryFile);
                    }
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

        /// <summary>
        /// Get value from xml file by index
        /// </summary>
        /// <param name="index">Index of the string</param>
        /// <returns>A string by index</returns>
        public string GetStringByIndex(int index)
        {

            if (entryStream == null || positions == null || positions.Count == 0)
            {
                throw new InvalidOperationException("Stream is not initialized or no positions found.");
            }

            entryStream.Seek(positions[index], SeekOrigin.Begin);
            var value = entryStream.ReadValue();
            return value;

        }
    }
}
