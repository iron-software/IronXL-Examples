***Based on <https://ironsoftware.com/examples/add-extract-remove-worksheet-images/>***

## Inserting Images into Spreadsheets

To add an image to your spreadsheet, utilize the `InsertImage` function. This method is compatible with multiple image formats including JPG/JPEG, BMP, PNG, GIF, and TIFF. To define the placement and size of the image, specify the coordinates for the top-left and bottom-right corners, which are determined by the differences in their column and row indices.

- For placing an image that spans 1x1 cells: `worksheet.InsertImage("image.gif", 5, 1, 6, 2);`
- For placing an image that spans 2x2 cells: `worksheet.InsertImage("image.gif", 5, 1, 7, 3);`

## Retrieving Images from a Worksheet

To retrieve images from a worksheet, use the `Images` attribute which lists all the images embedded in the sheet. This feature allows for several functionalities including exporting images, adjusting their sizes, determining their positions, and accessing the images' byte data. It is important to note that the image IDs increase in an odd sequence such as 1, 3, 5, 7, etc.

## Deleting Images from Worksheets

Building off the previous example of image extraction, deleting an image is straightforward. Simply provide the unique ID of the image to the `RemoveImage` method for deletion from your worksheet.