using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Reflection.PortableExecutable;
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

            using var reader = XmlReader.Create(bytes, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true
            });



            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "row")
                {
                    var row = new FileRowInfoExcel();

                    while (reader.Read())
                    {
                        if (reader.Name == "c" && reader.NodeType == XmlNodeType.Element)
                        {
                            string cellRef = reader.GetAttribute("r"); // A1, B1, D1
                            string cellType = reader.GetAttribute("t"); // s, n, b, etc.

                            row.Columns.Add(new FileColumnInfoExcel
                            {
                                Parameters = new Dictionary<string, string>()
                                {
                                    {
                                        "r", cellRef
                                    },
                                    {   "t", cellType }
                                }
                            });
                        }

                        if (reader.Name == "v" && reader.NodeType == XmlNodeType.Element)
                        {
                            row.Columns[^1].V = reader.ReadElementContentAsString();
                        }

                        if (reader.Name == "row" && reader.NodeType == XmlNodeType.EndElement)
                        {
                            break;
                        }

                    }

                    yield return row;

                }


            }



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
