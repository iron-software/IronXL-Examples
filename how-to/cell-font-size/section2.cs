using IronXL.Styles;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CellFontSize
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].StringValue = "Advanced";
            
            // Set font family
            workSheet["B2"].Style.Font.Name = "Lucida Handwriting";
            
            // Set font script
            workSheet["B2"].Style.Font.FontScript = FontScript.None;
            
            // Set underline
            workSheet["B2"].Style.Font.Underline = FontUnderlineType.Double;
            
            // Set bold property
            workSheet["B2"].Style.Font.Bold = true;
            
            // Set italic property
            workSheet["B2"].Style.Font.Italic = false;
            
            // Set strikeout property
            workSheet["B2"].Style.Font.Strikeout = false;
            
            // Set font color
            workSheet["B2"].Style.Font.Color = "#00FFFF";
            
            workBook.SaveAs("fontAndSizeAdvanced.xlsx");
        }
    }
}