using IronXL.Styles;
using IronXL.Excel;
namespace ironxl.CellFontSize
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].StringValue = "Font and Size";
            
            // Set font family
            workSheet["B2"].Style.Font.Name = "Times New Roman";
            
            // Set font size
            workSheet["B2"].Style.Font.Height = 15;
            
            // Set font to bold
            workSheet["B2"].Style.Font.Bold = true;
            
            // Set underline
            workSheet["B2"].Style.Font.Underline = FontUnderlineType.Single;
            
            workBook.SaveAs("fontAndSize.xlsx");
        }
    }
}