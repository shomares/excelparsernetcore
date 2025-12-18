

namespace ExcelReader.src.Entity
{
    internal record FileRowInfoExcel
    {

        public string R { get; init; } = string.Empty;
        public List<FileColumnInfoExcel> Columns { get; set; } = [];
    }

    internal record FileColumnInfoExcel
    {
        public string T { get; init; } = string.Empty;

        public string V { get; init; } = string.Empty;

    }
}
