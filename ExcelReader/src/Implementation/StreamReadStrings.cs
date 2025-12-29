using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Text;
using System.Xml;

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

        public async Task LoadInfoAsync(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
        {
            if (fileInfoExcel.PartName == null)
            {
                return;
            }

            positions = [];
            temporaryFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".bin");
            entryStream = new FileStream(temporaryFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None);

            var entry = zipArchive.GetEntry(fileInfoExcel.PartName.TrimStart('/'));
            if (entry == null)
            {
                return;
            }

            using var xml = XmlReader.Create(entry.Open(), new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                Async = true,
            });

            using var writer = new BinaryWriter(entryStream, Encoding.UTF8, leaveOpen: true);

            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.Name == "t")
                {
                    positions.Add(entryStream.Position);
                    string value = await xml.ReadContentAsStringAsync();
                    writer.Write(value);
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
                    entryStream?.Dispose();

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
            using var reader = new BinaryReader(entryStream, Encoding.UTF8, leaveOpen: true);
            return reader.ReadString();
        }
    }
}
