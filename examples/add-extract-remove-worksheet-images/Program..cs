using IronXL;
using IronSoftware.Drawing;
using System.Collections.Generic;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Insert images
workSheet.InsertImage("ironpdf.jpg", 2, 2, 4, 4);
workSheet.InsertImage("ironpdfIcon.png", 2, 6, 4, 8);

// Retreive images
List<IronXL.Drawing.Images.IImage> images = workSheet.Images;
// Select each image
foreach (IronXL.Drawing.Images.IImage image in images)
{
    // Save the image
    AnyBitmap anyBitmap = image.ToAnyBitmap();
    anyBitmap.SaveAs($"{image.Id}.png");
}

// Remove image
workSheet.RemoveImage(3);

workBook.SaveAs("images.xlsx");
