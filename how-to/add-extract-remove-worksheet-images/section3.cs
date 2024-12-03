using IronXL;
using IronXL.Excel;
namespace ironxl.AddExtractRemoveWorksheetImages
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("insertImages.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Remove image
            workSheet.RemoveImage(3);
            
            workBook.SaveAs("removeImage.xlsx");
        }
    }
}