using System.IO;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ConvertSpreadsheetFileTypes
{
    public static class Section2
    {
        public static void Run()
        {
            // Import any XLSX, XLS, XLSM, XLTX, CSV and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Export the excel file as XLS, XLSX, XLSM, CSV, TSV, JSON, XML
            workBook.SaveAs("sample.xls");
            workBook.SaveAs("sample.xlsx");
            workBook.SaveAs("sample.tsv");
            workBook.SaveAsCsv("sample.csv");
            workBook.SaveAsJson("sample.json");
            workBook.SaveAsXml("sample.xml");
            
            // Export the excel file as Html, Html string
            workBook.ExportToHtml("sample.html");
            string htmlString = workBook.ExportToHtmlString();
            
            // Export the excel file as Binary, Byte array, Data set, Stream
            byte[] binary = workBook.ToBinary();
            byte[] byteArray = workBook.ToByteArray();
            System.Data.DataSet dataSet = workBook.ToDataSet(); // Allow easy integration with DataGrids, SQL and EF
            Stream stream = workBook.ToStream();
        }
    }
}