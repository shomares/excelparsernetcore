using ExcelReader.src.Entity;
using ExcelReader.src.Interfaces;
using System.IO.Compression;
using System.Text;


namespace ExcelReader.src.Implementation
{
    internal class ReaderFileSimple : IReaderFileZip
    {
        public async Task<FileInfoExcel[]> GetFileInfosAsync(ZipArchive zipArchive)
        {
            var entry = zipArchive.Entries.FirstOrDefault(it => it.Name == "[Content_Types].xml");

            if (entry == null)
            {
                return [];
            }



            using var entryStream = entry.Open();
            using var memoryStream = new MemoryStream();
            await entryStream.CopyToAsync(memoryStream);
            var bytes = memoryStream.ToArray();

            var items = ParseContent(bytes);
            return items;

        }


        private static FileInfoExcel[] ParseContent(byte[] bytes)
        {
            var content = Encoding.UTF8.GetString(bytes);

            var index = 0;
            var foundedChar = 0;
            var result = new List<FileInfoExcel>();


            var field = "<Override";
            var partName = "PartName=";
            var contentType = "ContentType=";

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

                // found
                var partNameIndex = content.IndexOf(partName, index);
                string partNameValue = string.Empty;
                string contentTypeValue = string.Empty;
                if (partNameIndex > 0)
                {
                    partNameIndex += partName.Length + 1;
                    var endIndex = content.IndexOf('"', partNameIndex);
                    partNameValue = content[partNameIndex..endIndex];
                }

                var contentTypeIndex = content.IndexOf(contentType, index);
                if (contentTypeIndex > 0)
                {
                    contentTypeIndex += contentType.Length + 1;
                    var endIndex = content.IndexOf('"', contentTypeIndex);
                    contentTypeValue = content[contentTypeIndex..endIndex];
                }

                result.Add(new FileInfoExcel
                {
                    PartName = partNameValue,
                    ContentType = contentTypeValue
                });

                foundedChar = 0;

                index++;
            }

            return [.. result];
        }
    }
}
