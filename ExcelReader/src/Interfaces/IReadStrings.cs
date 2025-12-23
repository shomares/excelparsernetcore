using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReadStrings: IDisposable
    {

        public void LoadInfo(FileInfoExcel fileInfoExcel, ZipArchive zipArchive);

        public string GetStringByIndex(int index);
    }
}
