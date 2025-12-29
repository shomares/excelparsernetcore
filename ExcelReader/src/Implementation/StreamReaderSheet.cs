using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Reader Sheet implementation, for big files, this process the file in a streaming way
    /// </summary>
    internal class StreamReaderSheet : IReaderSheet
    {

        private Stream? entryStream;
        private StreamReader? streamReader;
        private bool disposedValue;

        /// <summary>
        /// Read the sheet from the zip archive, using streaming in order to reduce memory consumption
        ///</summary>
        public IEnumerable<FileRowInfoExcel> ReadSheet(string sheetPartName, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == sheetPartName.TrimStart('/'));
            if (entry == null)
            {
                return [];
            }

            entryStream = entry.Open();
            streamReader = new StreamReader(entryStream);
            return ParseContent(streamReader);

        }

        /// <summary>
        /// Parse the content of the sheet xml file, reading it as a stream
        /// </summary>
        /// <param name="bytes">Stream to read</param>
        /// <returns>A collection of rows, row by row</returns>
        private static IEnumerable<FileRowInfoExcel> ParseContent(StreamReader bytes)
        {

            using var xml = XmlReader.Create(bytes, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true
            });



            var currentElement = new StringBuilder();
            FileRowInfoExcel? rowInfo = null;

            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.Name == "row")
                {

                    rowInfo = new FileRowInfoExcel
                    {
                        Parameters = ReaderValues.ReadParameters(xml),
                        Columns = ReadColumns(xml)
                    };
                    yield return rowInfo;
                }
            }

        }

        private static List<FileColumnInfoExcel> ReadColumns(XmlReader xml)
        {
            var columns = new List<FileColumnInfoExcel>();
            while (xml.Read() && xml.NodeType == XmlNodeType.Element && xml.Name == "c")
            {
                FileColumnInfoExcel columnInfo = new()
                {
                    Parameters = ReaderValues.ReadParameters(xml)
                };


                xml.Read();
                if (!xml.IsEmptyElement
                    && xml.NodeType == XmlNodeType.Element && xml.Name == "v")
                {
                    columnInfo.V = xml.ReadElementContentAsString();
                }

                columns.Add(columnInfo);

            }

            return columns;
        }




        /// <summary>
        /// Dispose value of streams
        /// </summary>
        /// <param name="disposing">Indicates if the resouece is disposing</param>
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
        /// Implementation of Dispose pattern
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
