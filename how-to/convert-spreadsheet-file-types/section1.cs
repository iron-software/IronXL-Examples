using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ConvertSpreadsheetFileTypes
{
    public static class Section1
    {
        public static void Run()
        {
            // Import any XLSX, XLS, XLSM, XLTX, CSV and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Export the excel file as XLS, XLSX, XLSM, CSV, TSV, JSON, XML
            workBook.SaveAs("sample.xls");
            workBook.SaveAs("sample.tsv");
            workBook.SaveAsCsv("sample.csv");
            workBook.SaveAsJson("sample.json");
            workBook.SaveAsXml("sample.xml");
            
            // Export the excel file as Html
            workBook.ExportToHtml("sample.html");
        }
    }
}