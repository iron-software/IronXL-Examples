using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section1
    {
        public static void Run()
        {
            /**
            Create & Save Excel File 
            anchor-create-and-save-an-excel-file
            **/
            using IronXL;
            
            
            //default file format is XLSX, we can override it using CreatingOptions
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var sheet = workbook.CreateWorkSheet("example_sheet");
            
            sheet ["A1"].Value = "Example";
            //set value to multiple cells
            sheet ["A2:A4"].Value = 5;
            sheet ["A5"].Style.SetBackgroundColor("#f0f0f0");
            //set style to multiple cells
            sheet ["A5:A6"].Style.Font.Bold = true;
            //set formula
            sheet ["A6"].Value = "=SUM(A2:A4)";
            
            if (sheet ["A6"].IntValue == sheet ["A2:A4"].IntValue)
            {
                Console.WriteLine("Basic test passed");
            }
            
            workbook.SaveAs("example_workbook.xlsx");
        }
    }
}