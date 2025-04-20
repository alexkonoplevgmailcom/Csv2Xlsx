namespace Csv2Xlsx
{
    partial class Program
    {
        class Mapping
        {
            public string SheetName { get; set; } // Name of the sheet to write data to
            public Dictionary<string, string> ColumnMappings { get; set; }
            public List<CellMapping> CellMappings { get; set; }
            public int StartRow { get; set; }
        }
    }
}
