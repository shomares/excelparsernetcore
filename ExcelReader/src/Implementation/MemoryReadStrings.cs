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


        public void LoadInfo(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
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
                IgnoreWhitespace = true
            });

            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.Name == "t")
                {
                    string value = xml.ReadElementContentAsString();
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

   

        public string GetStringByIndex(int index)
        {
            return values != null ? values[index] : string.Empty;
        }
    }
}
