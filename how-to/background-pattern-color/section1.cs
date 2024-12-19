using IronSoftware.Drawing;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.BackgroundPatternColor
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set background pattern
            workSheet["A1"].Style.FillPattern = FillPattern.AltBars;
            workSheet["A2"].Style.FillPattern = FillPattern.ThickVerticalBands;
            
            // Set background color
            workSheet["A1"].Style.SetBackgroundColor(Color.Aquamarine);
            workSheet["A2"].Style.BackgroundColor = "#ADFF2F";
            
            workBook.SaveAs("setBackgroundPattern.xlsx");
        }
    }
}