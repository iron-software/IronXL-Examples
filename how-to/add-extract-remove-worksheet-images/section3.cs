using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AddExtractRemoveWorksheetImages
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("insertImages.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Remove image
            workSheet.RemoveImage(3);
            
            workBook.SaveAs("removeImage.xlsx");
        }
    }
}