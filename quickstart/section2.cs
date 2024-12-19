using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section2
    {
        public static void Run()
        {
            // Create new Excel WorkBook document.
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.Metadata.Author = "IronXL";
            
            // Add a blank WorkSheet
            WorkSheet workSheet = workBook.CreateWorkSheet("main_sheet");
            
            // Add data and styles to the new worksheet
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
            
            // Save the excel file
            workBook.SaveAs("NewExcelFile.xlsx");
        }
    }
}