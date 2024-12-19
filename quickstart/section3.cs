using IronXL.Excel;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section3
    {
        public static void Run()
        {
            // Export to many formats with fluent saving
            workSheet.SaveAs("NewExcelFile.xls");
            workSheet.SaveAs("NewExcelFile.xlsx");
            workSheet.SaveAsCsv("NewExcelFile.csv");
            workSheet.SaveAsJson("NewExcelFile.json");
            workSheet.SaveAsXml("NewExcelFile.xml");
        }
    }
}