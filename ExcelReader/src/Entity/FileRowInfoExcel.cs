

namespace ExcelReader.src.Entity
{
    internal record FileRowInfoExcel
    {

        public IDictionary<string, string>? Parameters { get; set; } = null;
        public List<FileColumnInfoExcel> Columns { get; set; } = [];
    }

    internal record FileColumnInfoExcel
    {

        public IDictionary<string, string>? Parameters { get; set; } = null;

        public string? V { get; set; } = null;
 
    }
}
