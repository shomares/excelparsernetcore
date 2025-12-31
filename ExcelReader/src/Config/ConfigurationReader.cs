
namespace ExcelReader.src.Config
{
    /// <summary>
    /// Configuration for reader implementation
    /// </summary>
    public class ConfigurationReader
    {
        /// <summary>
        /// Specify if use memory implementation for strings reading,
        /// true to use memory implementation for files small, false to use stream implementation
        /// </summary>
        public bool UseMemoryForStrings { get; set; } = true;
    }
}
