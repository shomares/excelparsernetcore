using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Xml;

namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Summary read string implementation
    /// </summary>
    internal class MemoryReadStrings : IReadStrings
    {
        private Stream? entryStream;
        private bool disposedValue;
        private List<string>? values = [];


        /// <summary>
        /// Load the information from the zip archive into memory
        /// </summary>
        /// <param name="fileInfoExcel">File info to read</param>
        /// <param name="zipArchive">Zip archive to read</param>
        public async Task LoadInfoAsync(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == fileInfoExcel.PartName?.TrimStart('/'));

            if (entry == null)
            {
                return;
            }

            values = [];

            entryStream = entry.Open();

            using var xml = XmlReader.Create(entryStream, new XmlReaderSettings
            {
                IgnoreComments = true,
                Async = true,
                IgnoreWhitespace = true,
            });

            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.Name == "t")
                {
                    string value = await xml.ReadElementContentAsStringAsync();
                    values.Add(value);
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


        /// <summary>
        /// Get the string by index in memory
        /// </summary>
        /// <param name="index">Index to read</param>
        /// <returns>String to read</returns>
        public string GetStringByIndex(int index)
        {
            if (values == null)
            {
                throw new InvalidOperationException("There is not values parsed");

            }
            if (index < 0 || index >= values.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), "Index is bigger than positions");
            }

            return values != null ? values[index] : string.Empty;
        }
    }
}
