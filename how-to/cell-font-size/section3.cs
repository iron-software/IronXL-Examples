using IronSoftware.Drawing;
using IronXL.Excel;
namespace ironxl.CellFontSize
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set Color property
            workSheet["B2"].Style.Font.Color = "#00FFFF";
            
            // Use Hex color code
            workSheet["B2"].Style.Font.SetColor("#00FFFF");
            
            // Use IronSoftware.Drawing
            workSheet["B2"].Style.Font.SetColor(Color.Red);
        }
    }
}