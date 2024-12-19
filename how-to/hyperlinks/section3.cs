using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.Hyperlinks
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set hyperlink to open file sample.xlsx
            workSheet["A1"].Value = "Open sample.xslx";
            workSheet["A1"].First().Hyperlink = "ftp://C:/Users/sample.xlsx";
            
            // Set hyperlink to open file sample.xlsx
            workSheet["A2"].Value = "Open sample.xslx";
            workSheet["A2"].First().Hyperlink = "file:///C:/Users/sample.xlsx";
            
            // Set hyperlink to email example@gmail.com
            workSheet["A3"].Value = "example@gmail.com";
            workSheet["A3"].First().Hyperlink = "mailto:example@gmail.com";
            
            workBook.SaveAs("setOtherHyperlink.xlsx");
        }
    }
}