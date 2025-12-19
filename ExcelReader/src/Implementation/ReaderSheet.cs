using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Text;

namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Reader Sheet implementation, for big files, this process the file in a streaming way
    /// </summary>
    internal class ReaderSheet : IReaderSheet
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
            var currentElement = new StringBuilder();
            FileRowInfoExcel? rowInfo = null;
            FileColumnInfoExcel? columnInfo = null;

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
                    //Read row properties
                    if (currentElementStr == "row")
                    {
                        rowInfo = new FileRowInfoExcel
                        {
                            Parameters = ReadParameters(bytes)
                        };
                    }
                    else if (currentElementStr == "c")
                    {
                        columnInfo = new FileColumnInfoExcel
                        {
                            Parameters = ReadParameters(bytes)
                        };

                    }

                    else if (currentElementStr == "v")
                    {
                        if (rowInfo != null && columnInfo != null)
                        {
                            columnInfo.V = ReadValue(bytes);
                        }
                    }
                    else if (currentElementStr == "/c")
                    {
                        if (rowInfo != null && columnInfo != null)
                        {
                            rowInfo.Columns.Add(columnInfo);
                            columnInfo = null;
                        }
                    }
                    else if (currentElementStr == "/row")
                    {
                        if (rowInfo != null)
                        {
                            yield return rowInfo;
                        }

                        rowInfo = new FileRowInfoExcel();
                    }
                    else if (currentElementStr == "/sheetData")
                    {
                        yield break;
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
        /// Read only the value inside a tag like <v>value</v>
        /// </summary>
        /// <param name="bytes">A stream to read</param>
        /// <returns>The value parsed</returns>
        private static string ReadValue(StreamReader bytes)
        {
            char? c = null;
            var result = new StringBuilder();
            while (!bytes.EndOfStream && c != '<')
            {
                c = (char)bytes.Read();

                if (c == '<')
                {
                    break;
                }

                result.Append(c);
            }

            return result.ToString();
        }


        /// <summary>
        /// Read parameters inside a tag like <c r="A1" t="s">
        /// </summary>
        /// <param name="bytes">A stream to read</param>
        /// <returns>Dictionary of read parameters parsed</returns>
        private static Dictionary<string, string> ReadParameters(StreamReader bytes)
        {
            var parameters = new Dictionary<string, string>(5);
            char? c = null;
            while (!bytes.EndOfStream && c != '>')
            {
                var parameterName = new StringBuilder();
                var parameterValue = new StringBuilder();
                c = (char)bytes.Read();

                if (c == '>')
                {
                    break;
                }

                while (!bytes.EndOfStream && c != '=')
                {
                    if (c != ' ')
                    {
                        parameterName.Append(c);
                    }

                    c = (char)bytes.Read();
                }

                c = (char)bytes.Read();
                c = (char)bytes.Read();

                while (!bytes.EndOfStream && c != '"')
                {
                    parameterValue.Append(c);
                    c = (char)bytes.Read();
                }

                parameters[parameterName.ToString()] = parameterValue.ToString();
            }

            return parameters;
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
