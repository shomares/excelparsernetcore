using ExcelReader.src.Interfaces;


namespace ExcelReader.src.Implementation
{
    internal class ReadStringsFactory : IReadStringsFactory
    {
        public IReadStrings CreateReadStrings(string fileName)
        {
            var fileInfo = new FileInfo(fileName);
            long fileSizeInBytes = fileInfo.Length;

            double fileSizeInMB = (double)fileSizeInBytes / (1024 * 1024);


            return fileSizeInMB > 10000
                ? new StreamReadStrings()
                : new MemoryReadStrings();
        }
    }
}
