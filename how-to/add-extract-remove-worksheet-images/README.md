# How to Add, Extract, and Remove Images from Worksheets

## Introduction

Integrating images into worksheets can significantly enhance the data presentation by including relevant graphics or illustrations. Conversely, extracting or deleting images can streamline content management and organization. Moreover, extracting images is particularly useful for reusing them in different documents or updating their contents. These functionalities collectively improve user interaction with images in Excel workbooks, making image handling more intuitive and flexible.

## Add Images Example

To embed an image into a spreadsheet, use the `InsertImage` method. This method is compatible with various image formats including JPG/JPEG, BMP, PNG, GIF, and TIFF. Placement of the image is determined by specifying the top-left and bottom-right corners, which define the image's dimensions through the column and row indices. Below are a couple of methods to insert different sized images:
- For a 1x1 image size:
  - `worksheet.InsertImage("image.gif", 5, 1, 6, 2);`
- For a 2x2 image size:
  - `worksheet.InsertImage("image.gif", 5, 1, 7, 3);`

Image IDs are generated in an odd-numbered series, like 1, 3, 5, 7, etc.

```cs
using IronXL;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Insert images with explanatory comments
workSheet.InsertImage("ironpdf.jpg", 2, 2, 4, 4);  // Insert an image named 'ironpdf.jpg'
workSheet.InsertImage("ironpdfIcon.png", 2, 6, 4, 8);  // Insert another image named 'ironpdfIcon.png'

workBook.SaveAs("insertImages.xlsx");
```

### Output Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-extract-remove-worksheet-images/insert-image.png" alt="Insert Image" class="img-responsive add-shadow">
    </div>
</div>

## Extract Images Example

To extract images from a worksheet, access the `Images` property, which lists all images within the sheet. This allows you to export, resize, or retrieve the byte data for each image. Just like in the adding example, image IDs increase in odd numbers.

```cs
using IronSoftware.Drawing;
using IronXL;
using IronXL.Drawing;
using System;
using System.Collections.Generic;

WorkBook workBook = WorkBook.Load("insertImages.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve images
List<IronXL.Drawing.Images.IImage> images = workSheet.Images;

// Loop through each image
foreach (IronXL.Drawing.Images.IImage image in images)
{
    // Save the image
    AnyBitmap anyBitmap = image.ToAnyBitmap();
    anyBitmap.SaveAs($"{image.Id}.png");

    // Adjust the image size
    image.Resize(1,3);

    // Determine image location
    Position position = image.Position;
    Console.WriteLine("top row index: " + position.TopRowIndex);
    Console.WriteLine("bottom row index: " + position.BottomRowIndex);

    // Get byte array data from image
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

## Remove Image Example

Building on the extract images example above, removing an image is as straightforward as using the image's ID with the `RemoveImage` method.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("insertImages.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Remove an image by its ID
workSheet.RemoveImage(3);

workBook.SaveAs("removeImage.xlsx");
```