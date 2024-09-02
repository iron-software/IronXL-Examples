## Example of Adding Images

To place an image inside a spreadsheet, utilize the method `InsertImage`. This function is compatible with several image formats including JPG/JPEG, BMP, PNG, GIF, and TIFF. You must define the coordinates for the top-left and bottom-right corners to set the size of the image, which is determined by the difference in column and row numbers.

- For an image with dimensions of 1x1: `worksheet.InsertImage("image.gif", 5, 1, 6, 2);`
- For an image with dimensions of 2x2: `worksheet.InsertImage("image.gif", 5, 1, 7, 3);`

## Example of Extracting Images

To retrieve images from a particular worksheet, use the `Images` property. This attribute provides access to a collection of all images in that worksheet. You can then manage these images by exporting them, changing their size, learning their coordinates, and accessing their byte data. It is important to note that the images are identified by odd numbers, following a sequence like 1, 3, 5, 7, etc.

## Example of Removing Images

Continuing from the example of extracting images, you can delete any specific image with ease. Just provide the unique ID of the image to the `RemoveImage` method, and it will be deleted from the worksheet. This process simplifies the management of images within your spreadsheets by allowing their removal through their respective IDs.