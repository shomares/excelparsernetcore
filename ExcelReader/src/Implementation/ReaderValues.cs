using System.Xml;


namespace ExcelReader.src.Implementation
{
    internal static class ReaderValues
    {


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
