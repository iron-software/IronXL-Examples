using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section3
    {
        public void Run()
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