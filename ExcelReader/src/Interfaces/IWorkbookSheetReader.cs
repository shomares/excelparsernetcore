using System.IO.Compression;

namespace ExcelReader.src.Interfaces
{
    internal interface IWorkbookSheetReader: IDisposable
    {
        string GetSheetNameById(string id);

        void LoadInfo(ZipArchive zipArchive); 
    }
}
