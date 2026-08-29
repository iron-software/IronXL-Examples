using IronXL;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Export to many formats with fluent saving
            workSheet.SaveAs("NewExcelFile.xls");
            workSheet.SaveAs("NewExcelFile.xlsx");
            workSheet.SaveAsCsv("NewExcelFile.csv");
            workSheet.SaveAsJson("NewExcelFile.json");
            workSheet.SaveAsXml("NewExcelFile.xml");
        }
    }
}