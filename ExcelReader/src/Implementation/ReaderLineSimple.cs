using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.Dynamic;
using System.Text;

namespace ExcelReader.src.Implementation
{
    internal class ReaderLineSimple(IReadStrings readStrings) : IReaderLine
    {

        /// <summary>
        /// Read columns of the row
        /// </summary>
        /// <param name="fileRowInfoExcel">Columns to read no parse to a r</param>
        /// <returns>A list of string</returns>
        public IDictionary<string, string> ReadColumns(FileRowInfoExcel fileRowInfoExcel)
        {
            var columns = new Dictionary<string, string>();

            foreach (var column in fileRowInfoExcel.Columns)
            {
                var r = column.Parameters != null && column.Parameters.TryGetValue("r", out var referenceValue)
                       ? referenceValue
                       : string.Empty;

                var value = GetValue(column);

                columns.Add(r, value ?? r);

            }

            return columns;
        }

        private string? GetValue(FileColumnInfoExcel column)
        {
            if (column.V == null)
            {
                return null;
            }


            if (column.Parameters != null && column.Parameters.TryGetValue("t", out var typeValue) && typeValue == "s")
            {
                if (int.TryParse(column.V, out var stringIndex))
                {
                    var stringValue = readStrings.GetStringByIndex(stringIndex);
                    return stringValue;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return column.V;
            }

        }


        private static string GetColumnName(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var index = 0;
            var sb = new StringBuilder();

            while (index < text.Length)
            {
                var c = text[index];

                if (c >= '0' && c <= '9')
                {
                    break;
                }

                sb.Append(c);
                index++;
            }

            sb.Append('1');
            return sb.ToString();
        }

        public dynamic ReadRow(FileRowInfoExcel fileRowInfoExcel, IDictionary<string, string> columns)
        {

            var result = new ExpandoObject();
            foreach (var column in fileRowInfoExcel.Columns)
            {
                var r = column.Parameters != null && column.Parameters.TryGetValue("r", out var referenceValue)
                       ? referenceValue
                       : string.Empty;

                r = GetColumnName(r);

                if (columns.TryGetValue(r, out var columnsPropertyName))
                {
#pragma warning disable CS8619 // La nulabilidad de los tipos de referencia del valor no coincide con el tipo de destino
                    var dict = (IDictionary<string, object>)result;
#pragma warning restore CS8619 // La nulabilidad de los tipos de referencia del valor no coincide con el tipo de destino
                    var value = GetValue(column);
                    dict.Add(columnsPropertyName, value ?? string.Empty);
                }
            }

            return result;
        }
    }
}
