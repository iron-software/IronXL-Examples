using IronXL.Styles;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.BorderAlignment
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].StringValue = "Top";
            workSheet["B4"].StringValue = "Forward";
            
            // Set top border line
            workSheet["B2"].Style.TopBorder.Type = BorderType.Thick;
            
            // Set diagonal border line
            workSheet["B4"].Style.DiagonalBorder.Type = BorderType.Thick;
            // Set diagonal border direction
            workSheet["B4"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;
            
            workBook.SaveAs("borderLines.xlsx");
        }
    }
}