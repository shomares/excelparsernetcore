using System.Text;
using System.Xml;


namespace ExcelReader.src.Implementation
{
    internal static class ReaderValues
    {

        /// <summary>
        /// Read only the value inside a tag like <v>value</v>
        /// </summary>
        /// <param name="bytes">A stream to read</param>
        /// <returns>The value parsed</returns>
        public static string ReadValue(this StreamReader bytes)
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
        public static Dictionary<string, string> ReadParameters(this XmlReader bytes)
        {
            var parameters = new Dictionary<string, string>(5);

            for (var index = 0; index < bytes.AttributeCount; index++)
            {
                bytes.MoveToAttribute(index);
                parameters.Add(bytes.Name, bytes.Value);

            }

            return parameters;
        }
    }
}
