using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ManageWorksheet
{
    public static class Section5
    {
        public static void Run()
        {
            WorkBook firstBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkBook secondBook = WorkBook.Create();
            
            // Select first worksheet in the workbook
            WorkSheet workSheet = firstBook.DefaultWorkSheet;
            
            // Duplicate the worksheet to the same workbook
            workSheet.CopySheet("Copied Sheet");
            
            // Duplicate the worksheet to another workbook with the specified name
            workSheet.CopyTo(secondBook, "Copied Sheet");
            
            firstBook.SaveAs("firstWorksheet.xlsx");
            secondBook.SaveAs("secondWorksheet.xlsx");
        }
    }
}