using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;


namespace ExcelReader.src.Implementation
{
    internal class ReaderSheet : IReaderSheet
    {

        public async Task<IEnumerable<FileRowInfoExcel>> ReadSheetAsync(string sheetPartName, string[] strings, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == sheetPartName.TrimStart('/'));
            if (entry == null)
            {
                return [];
            }
            using var entryStream = entry.Open();
            using var stream = new StreamReader(entryStream);
            return ParseContent(stream, strings);
        }

        private static IEnumerable<FileRowInfoExcel> ParseContent(StreamReader bytes, string[] strings)
        {
            var currentElement = string.Empty;
            var elements = new Stack<string>();

            var rowName = "row";

            while (!bytes.EndOfStream)
            {
                var c = (char)bytes.Read();

                if (c == '<')
                {
                    while (!bytes.EndOfStream && c != ' ' && c != '>')
                    {
                        c = (char)bytes.Read();

                        if (c != '>')
                        {
                            currentElement += c;
                        }
                    }


                    //read the parameters
                    if (currentElement == rowName)
                    {
                        while (!bytes.EndOfStream && c == "\"")
                        {

                        }

                    }


                    while (!bytes.EndOfStream && c != '>')
                    {
                        c = (char)bytes.Read();
                    }




                    elements.Push(currentElement);
                    currentElement = string.Empty;
                }
            }

            return [];
        }


    }
}
