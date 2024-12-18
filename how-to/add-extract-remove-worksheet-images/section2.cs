using System.Collections.Generic;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AddExtractRemoveWorksheetImages
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("insertImages.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Retreive images
            List<IronXL.Drawing.Images.IImage> images = workSheet.Images;
            
            // Select each image
            foreach (IronXL.Drawing.Images.IImage image in images)
            {
                // Save the image
                AnyBitmap anyBitmap = image.ToAnyBitmap();
                anyBitmap.SaveAs($"{image.Id}.png");
            
                // Resize the image
                image.Resize(1,3);
            
                // Retrieve image position
                Position position = image.Position;
                Console.WriteLine("top row index: " + position.TopRowIndex);
                Console.WriteLine("bottom row index: " + position.BottomRowIndex);
            
                // Retrieve byte data
                byte[] imageByte = image.Data;
            }
            
            workBook.SaveAs("resizeImage.xlsx");
        }
    }
}