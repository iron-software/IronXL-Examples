***Based on <https://ironsoftware.com/examples/add-extract-remove-worksheet-images/>***

## Example of Adding Images

To incorporate an image within a spreadsheet, utilize the `InsertImage` method, which is compatible with a range of image formats including JPG/JPEG, BMP, PNG, GIF, and TIFF. Define the placement of the image by specifying the coordinates for the top-left and bottom-right corners, which helps determine the size based on the difference in column and row indices.

- For an image occupying a single cell: `worksheet.InsertImage("image.gif", 5, 1, 6, 2);`
- For an image spanning over four cells: `worksheet.InsertImage("image.gif", 5, 1, 7, 3);`

## Example of Extracting Images

When you need to retrieve images from a specific worksheet, use the Images property. This property provides access to a comprehensive list of images in the worksheet. With this list, you can execute various tasks such as exporting images, adjusting their size, and acquiring both the position and the binary data of each image. It’s important to note that the IDs of these images increase by odd numbers: 1, 3, 5, 7, etc.

## Example of Removing Images

Building on the process of extracting images, an image can be readily deleted from the worksheet by referring to its unique index number. Employ the `RemoveImage` method and specify the image's ID number to eliminate it from the worksheet.