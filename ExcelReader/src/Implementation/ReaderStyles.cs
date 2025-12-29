using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Xml;

namespace ExcelReader.src.Implementation
{
    internal class ReaderStyles : IReaderStyles
    {
        private List<CellStyle>? styles;

        private Stream? entryStream = null;
        private bool disposedValue;

        public void LoadInfo(ZipArchive zipArchive)
        {
            var formatValue = new Dictionary<int, string>();
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == "xl/styles.xml");

            if (entry == null)
            {
                return;
            }

            styles = [];

            entryStream = entry.Open();
            using var reader = XmlReader.Create(entryStream, new XmlReaderSettings
            {
                IgnoreComments = true,
                Async = true,
                IgnoreWhitespace = true,
            });

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "numFmt")
                {
                    int numFmtId = int.Parse(reader.GetAttribute("numFmtId")!);
                    string? formatCode = reader.GetAttribute("formatCode");
                    formatValue.Add(numFmtId, formatCode ?? string.Empty);
                }


                if (reader.NodeType == XmlNodeType.Element && reader.Name == "cellXfs")
                {
                    var count = int.Parse(reader.GetAttribute("count")!);

                    for (var index = 0; index < count; index++)
                    {
                        reader.Read();

                        if (reader.NodeType != XmlNodeType.Element || reader.Name != "xf")
                        {
                            break;
                        }

                        var numFmtId = int.Parse(reader.GetAttribute("numFmtId")!);
                        styles.Add(new CellStyle
                        {
                            NumFmtId = numFmtId,
                            IsDate = IsDate(numFmtId) || (formatValue.TryGetValue(numFmtId, out var code) &&
                                IsDateFormatCode(code))
                        });

                    }

                    break;
                }
            }
        }

        private static bool IsDate(int numFmtId)
        {
            return numFmtId switch
            {
                14 or 15 or 16 or 17 or
                18 or 19 or 20 or 21 or
                22 or
                45 or 46 or 47 => true,
                _ => false
            };
        }

        private static bool IsDateFormatCode(string format)
        {
            format = format.ToLowerInvariant();

            return format.Contains("yy") ||
                   format.Contains("mm") ||
                   format.Contains("dd") ||
                   format.Contains("hh") ||
                   format.Contains("ss");
        }

        public CellStyle? GetCellStyle(int index)
        {
            if (styles == null)
            {
                return null;
            }

            return styles[index];
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    styles = null;
                    entryStream?.Dispose();
                }

                // TODO: liberar los recursos no administrados (objetos no administrados) y reemplazar el finalizador
                // TODO: establecer los campos grandes como NULL
                disposedValue = true;
            }
        }

        // // TODO: reemplazar el finalizador solo si "Dispose(bool disposing)" tiene código para liberar los recursos no administrados
        // ~ReaderStyles()
        // {
        //     // No cambie este código. Coloque el código de limpieza en el método "Dispose(bool disposing)".
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // No cambie este código. Coloque el código de limpieza en el método "Dispose(bool disposing)".
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
