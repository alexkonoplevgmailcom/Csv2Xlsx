namespace Csv2Xlsx
{
    partial class Program
    {
        class CellMapping
        {
            public string CsvColumn { get; set; }
            public int CsvRow { get; set; }
            public string ExcelColumn { get; set; }
            public int? ExcelRow { get; set; } // Optional hardcoded row number
            public int? OffsetFromEnd { get; set; } // Optional offset from the end of the table
        }
    }
}
