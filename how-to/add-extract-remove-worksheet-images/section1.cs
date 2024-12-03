using IronXL;
using IronXL.Excel;
namespace ironxl.AddExtractRemoveWorksheetImages
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Insert images
            workSheet.InsertImage("ironpdf.jpg", 2, 2, 4, 4);
            workSheet.InsertImage("ironpdfIcon.png", 2, 6, 4, 8);
            
            workBook.SaveAs("insertImages.xlsx");
        }
    }
}