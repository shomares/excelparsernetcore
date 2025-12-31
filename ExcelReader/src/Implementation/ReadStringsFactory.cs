using ExcelReader.src.Config;
using ExcelReader.src.Interfaces;


namespace ExcelReader.src.Implementation
{
    /// <summary>
    /// Implementation of factory to create read strings
    /// </summary>
    internal class ReadStringsFactory : IReadStringsFactory
    {
        /// <summary>
        /// Create read strings implementation
        /// </summary>
        /// <param name="fileName">Name of file</param>
        /// <param name="configuration">Configuration reader, if it is null then MemoryReadStrings is used when it's size is minor than 2mb otherwise StreamReadStrings is used</param>
        /// <returns>A reader string implementation</returns>
        public IReadStrings CreateReadStrings(string fileName, ConfigurationReader? configuration = null)
        {
            if (configuration == null)
            {
                var fileInfo = new FileInfo(fileName);
                long fileSizeInBytes = fileInfo.Length;

                double fileSizeInMB = (double)fileSizeInBytes / (1024 * 1024);


                return fileSizeInMB > 2
                    ? new StreamReadStrings()
                    : new MemoryReadStrings();
            }


            return configuration.UseMemoryForStrings
                ? new MemoryReadStrings()
                : new StreamReadStrings();
        }
    }
}
