# How to Add, Extract, and Remove Images from Worksheets

***Based on <https://ironsoftware.com/how-to/add-extract-remove-worksheet-images/>***


## Introduction

Incorporating images into spreadsheets can significantly enhance the presentation of data by integrating pertinent graphics or illustrations. Conversely, the ability to delete or remove images from a workbook simplifies the content management process. Furthermore, extracting images from spreadsheets is crucial for their reutilization in different contexts or for updating them within the same document. These capabilities collectively improve the user experience by allowing for efficient image handling within Excel workbooks.

## Begin Using IronXL

---

## Example of Adding Images

When embedding an image in a spreadsheet, you can use the `InsertImage` method. This method is compatible with various image formats including JPG/JPEG, BMP, PNG, GIF, and TIFF. It requires the specification of the coordinates of the top-left and bottom-right corners to set its dimensions. For instance:
- For sizing the image to 1x1:
  - `worksheet.InsertImage("image.gif", 5, 1, 6, 2);`
- For a 2x2 image size:
  - `worksheet.InsertImage("image.gif", 5, 1, 7, 3);`

Generated image IDs are numbered in an odd sequence, such as 1, 3, 5, 7, etc.

```cs
using IronXL;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Adding images to the worksheet
workSheet.InsertImage("ironpdf.jpg", 2, 2, 4, 4);
workSheet.InsertImage("ironpdfIcon.png", 2, 6, 4, 8);

workBook.SaveAs("insertImages.xlsx");
```

### Output Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-extract-remove-worksheet-images/insert-image.png" alt="Insert Image" class="img-responsive add-shadow">
    </div>
</div>

## Extracting Images Example

Accessing the `Images` property on a worksheet allows you to retrieve a list containing all the embedded images. This list can be used to perform several operations such as exporting the images, resizing them, finding their positions, or accessing their binary data. Note that the IDs of the images have an odd-numbered sequence.

```cs
using IronSoftware.Drawing;
using IronXL;
using IronXL.Drawing;
using System;
using System.Collections.Generic;

WorkBook workBook = WorkBook.Load("insertImages.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Extracting images
List<IronXL.Drawing.Images.IImage> images = workSheet.Images;

// Processing each image
foreach (IronXL.Drawing.Images.IImage image in images)
{
    // Save the image
    AnyBitmap anyBitmap = image.ToAnyBitmap();
    anyBitmap.SaveAs($"{image.Id}.png");

    // Resize the image
    image.Resize(1,3);

    // Obtain image position
    Position position = image.Position;
    Console.WriteLine("top row index: " + position.TopRowIndex);
    Console.WriteLine("bottom row index: " + position.BottomRowIndex);

    // Access binary data of the image
    byte[] imageByte = image.Data;
}

workBook.SaveAs("resizeImage.xlsx");
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

## Removing Images Example

Referencing the previous example of extracting images, one can delete a specific image by providing its ID to the `RemoveImage` method. This effectively removes the designated image from the workbook.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("insertImages.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Removing an image by ID
workSheet.RemoveImage(3);

workBook.SaveAs("removeImage.xlsx");
```