# How to Add, Extract, and Remove Images from Worksheets

***Based on <https://ironsoftware.com/how-to/add-extract-remove-worksheet-images/>***


## Introduction

Incorporating images into spreadsheets can greatly enhance the visual appeal and relevance of the data presented. Conversely, removing images helps streamline content management and editing. The capability to extract images is particularly useful for reusing them in different contexts or for updating them in ongoing projects. Together, these functionalities empower users with robust image management tools, improving both the user experience and the efficiency of working with images in Excel workbooks.

## Add Images Example

To embed an image in a spreadsheet, you can use the `InsertImage` method. This method works with various image formats including JPG/JPEG, BMP, PNG, GIF, and TIFF. To define the image's size, you need to provide the coordinates for both the top-left and bottom-right corners, calculated by the respective column and row indices. Here are a couple of examples to demonstrate:
- For an image size of 1x1:
  - `worksheet.InsertImage("image.gif", 5, 1, 6, 2);`
- For an image size of 2x2:
  - `worksheet.InsertImage("image.gif", 5, 1, 7, 3);`

Image IDs are assigned in an incremental pattern of 1, 3, 5, 7, and so on.

```cs
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
```

### Output Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-extract-remove-worksheet-images/insert-image.png" alt="Insert Image" class="img-responsive add-shadow">
    </div>
</div>

## Extract Images Example

To retrieve images from a worksheet, access the `Images` property. This property provides a list of all images in the worksheet. You can then export, resize, get the location, and acquire the byte array of each image. Image IDs continue their sequence with odd numbers such as 1, 3, 5, and 7.

```cs
using System.Collections.Generic;
using IronXL.Excel;
namespace ironxl.AddExtractRemoveWorksheetImages
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("insertImages.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Retrieve images
            List<IronXL.Drawing.Images.IImage> images = workSheet.Images;
            
            // Process each image
            foreach (IronXL.Drawing.Images.IImage image in images)
            {
                // Save the image
                AnyBitmap anyBitmap = image.ToAnyBitmap();
                anyBitmap.SaveAs($"{image.Id}.png");
            
                // Resize the image
                image.Resize(1,3);
            
                // Get image position
                Position position = image.Position;
                Console.WriteLine("top row index: " + position.TopRowIndex);
                Console.WriteLine("bottom row index: " + position.BottomRowIndex);
            
                // Access byte data
                byte[] imageByte = image.Data;
            }
            
            workBook.SaveAs("resizeImage.xlsx");
        }
    }
}
```

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/add-extract-remove-worksheet-images/extract-image.png" alt="Extracted Images" class="img-responsive add-shadow" >
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Extracted Images</p>
    </div>
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/add-extract-remove-worksheet-images/image-size.png" alt="Image Size" class="img-responsive add-shadow">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Image Size</p>
    </div>
</div>

## Remove Image Example

Following the extract images example, if you wish to remove an image, simply provide its ID to the `RemoveImage` method. This operation will delete the image from its worksheet.

```cs
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
```