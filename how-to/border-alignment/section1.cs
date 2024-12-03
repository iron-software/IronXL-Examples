using IronXL.Styles;
using IronXL.Excel;
namespace ironxl.BorderAlignment
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].Value = "B2";
            
            // Set cell border
            workSheet["B2"].Style.LeftBorder.Type = BorderType.MediumDashed;
            workSheet["B2"].Style.RightBorder.Type = BorderType.MediumDashed;
            
            // Set text alignment
            workSheet["B2"].Style.HorizontalAlignment = HorizontalAlignment.Center;
            
            workBook.SaveAs("setBorderAndAlignment.xlsx");
        }
    }
}