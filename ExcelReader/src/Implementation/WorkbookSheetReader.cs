using System.IO.Compression;
using System.Xml;
using ExcelReader.src.Interfaces;

namespace ExcelReader.src.Implementation
{
    internal class WorkbookSheetReader : IWorkbookSheetReader

    {

        private readonly Dictionary<string, string> sheetNames = new();

         /// <summary>
        /// Dispose pattern implementation
        /// </summary>
        /// <param name="disposing">Indicate if resources is disposings</param>
        protected virtual void Dispose(bool disposing)
        {
                if (disposing)
                {
                    sheetNames.Clear();
                }
        }

        /// <summary>
        /// Release pattern implementation
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }


        public string GetSheetNameById(string id) => sheetNames.TryGetValue(id, out var name) ? name : string.Empty;
      
        public void LoadInfo(ZipArchive zipArchive)
        {
             var entry = zipArchive.GetEntry("xl/workbook.xml");
            // Load workbook info from the entry
            if(entry == null)
            {
                return;
            }

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
            });


            while(reader.Read())
            {
             
                // Process workbook XML content
                if(reader.IsStartElement() && reader.Name == "sheet")
                {
                    var name = reader.GetAttribute("name");
                    var sheetId = reader.GetAttribute("sheetId");
                    // Store or process sheet information as needed

                    sheetNames.TryAdd(name!.ToLower(), sheetId!);
                    sheetNames.TryAdd(name, sheetId!);
                }
        }
    }
    }

}