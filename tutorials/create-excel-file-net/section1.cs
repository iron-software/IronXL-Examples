using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section1
    {
        public static void Run()
        {
            // Default file format is XLSX, we can override it using CreatingOptions
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            var workSheet = workBook.CreateWorkSheet("example_sheet");
            workSheet["A1"].Value = "Example";
            
            // Set value to multiple cells
            workSheet["A2:A4"].Value = 5;
            workSheet["A5"].Style.SetBackgroundColor("#f0f0f0");
            
            // Set style to multiple cells
            workSheet["A5:A6"].Style.Font.Bold = true;
            
            // Set formula
            workSheet["A6"].Value = "=SUM(A2:A4)";
            if (workSheet["A6"].IntValue == workSheet["A2:A4"].IntValue)
            {
                Console.WriteLine("Basic test passed");
            }
            workBook.SaveAs("example_workbook.xlsx");
        }
    }
}