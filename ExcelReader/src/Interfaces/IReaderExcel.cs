namespace ExcelReader.src.Interfaces
{
    public interface IReaderExcel
    {
        Task<IEnumerable<T>> ReadExcelAsync<T>(string filePath, int sheetIndex = 0) where T : new();
    }
}
