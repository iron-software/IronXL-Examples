using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.Hyperlinks
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Modify the cell's property
            workSheet["A1"].Value = "Link to ironpdf.com";
            
            // Set hyperlink at A1 to https://ironpdf.com/
            workSheet.GetCellAt(0, 0).Hyperlink = "https://ironpdf.com/";
            
            workBook.SaveAs("setLinkHyperlink.xlsx");
        }
    }
}