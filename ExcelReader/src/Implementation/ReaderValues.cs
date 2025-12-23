using System.Text;


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


        public static string ReadValue(this FileStream bytes)
        {
            char? c = null;
            var result = new StringBuilder();
            do
            {
                c = bytes.ReadChar();

                if (c == '<')
                {
                    break;
                }

                result.Append(c);
            } while (c != '\0' && c != '<');

            return result.ToString();
        }

        /// <summary>
        /// Read parameters inside a tag like <c r="A1" t="s">
        /// </summary>
        /// <param name="bytes">A stream to read</param>
        /// <returns>Dictionary of read parameters parsed</returns>
        public static Dictionary<string, string> ReadParameters(this StreamReader bytes)
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

        public static char ReadChar(this FileStream fs)
        {
            var decoder = Encoding.UTF8.GetDecoder();
            byte[] bytes = new byte[4];   // UTF-8 max bytes per char
            char[] chars = new char[1];


            int read = fs.Read(bytes, 0, 1);
            if (read == 0)
            {
                return '\0';
            }


            decoder.Convert(
                bytes, 0, read,
                chars, 0, 1,
                false,
                out _, out int charsUsed, out _
            );

            return charsUsed > 0 ? chars[0] : '\0';

        }
    }
}
