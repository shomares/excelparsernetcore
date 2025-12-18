using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;

namespace ExcelReader.src.Implementation
{
    internal class ReadStrings : IReadStrings
    {

        public async Task<string[]> ReadStringsAsync(FileInfoExcel fileInfoExcel, ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.FullName == fileInfoExcel.PartName?.TrimStart('/'));

            if (entry == null)
            {
                return [];
            }

            var response = new List<string>();
            using var entryStream = entry.Open();
            using var memoryStream = new MemoryStream();
            await entryStream.CopyToAsync(memoryStream);
            var bytes = memoryStream.ToArray();
            return ParseContent(bytes);


        }

        private static string[] ParseContent(byte[] bytes)
        {
            var content = System.Text.Encoding.UTF8.GetString(bytes);
            var result = new List<string>();

            var index = 0;
            var foundedChar = 0;

            var field = "<si>";
            var partName = "<t>";

            while (index < content.Length)
            {
                var c = content[index];

                while (c == field[foundedChar])
                {
                    foundedChar++;

                    if (foundedChar == field.Length)
                    {
                        break;
                    }

                    index++;

                    if (index >= content.Length)
                    {
                        break;
                    }

                    c = content[index];
                }

                if (foundedChar < field.Length)
                {
                    foundedChar = 0;
                    index++;
                    continue;
                }

                if (foundedChar == field.Length)
                {
                    // found
                    var sIndex = content.IndexOf(partName, index);
                    string s = string.Empty;
                    if (sIndex > 0)
                    {
                        sIndex += partName.Length;
                        var endIndex = content.IndexOf('<', sIndex);
                        s = content[sIndex..endIndex];
                    }

                    result.Add(s);
                    foundedChar = 0;

                }

                index++;
            }

            return [.. result];
        }
    }
}
