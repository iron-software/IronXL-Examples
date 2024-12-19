using IronXL.Formatting;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.SetCellDataFormat
{
    public static class Section3
    {
        public static void Run()
        {
            // Create a new workbook
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Use builtin formats
            workSheet["A1"].Value = 123;
            workSheet["A1"].FormatString = BuiltinFormats.Accounting0;
            
            workBook.SaveAs("builtinDataFormats.xlsx");
        }
    }
}