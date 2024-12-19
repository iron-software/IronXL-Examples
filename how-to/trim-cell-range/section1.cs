using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.TrimCellRange
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["A2"].Value = "A2";
            workSheet["A3"].Value = "A3";
            
            workSheet["B1"].Value = "B1";
            workSheet["B2"].Value = "B2";
            workSheet["B3"].Value = "B3";
            workSheet["B4"].Value = "B4";
            
            // Retrieve column
            RangeColumn column = workSheet.GetColumn(0);
            
            // Apply trimming
            Range trimmedColumn = workSheet.GetColumn(0).Trim();
        }
    }
}