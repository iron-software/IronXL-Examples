# How to Copy Cells


The "Copy cell" function enables the replication of a cell’s contents—including data, formulas, and formatting—into another cell or a range of cells. This is especially useful for spreading consistent data, formulas, or styles across a worksheet.

Moreover, the `Copy` method preserves any existing styling, thereby ensuring that data replication across single or multiple worksheets remains precise and seamless when using IronXL.

## Example: Copying a Single Cell

To clone the contents of a specific cell, employ the `Copy` method. This requires specifying the source worksheet as the first argument and the destination position as the second. Notably, the `Copy` method conserves the original cell's styling.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Duplicate a cell's contents
workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet1"), "B3");

workBook.SaveAs("copiedSingleCell.xlsx");
```

### Resulting Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper" width='70%'>
         <img src="https://ironsoftware.com/static-assets/excel/how-to/copy-cells/copy-cells-copy-single-cell.png" alt="Copied Single Cell" class="img-responsive add-shadow">
    </div>
</div>

## Example: Copying a Cell Range

As with the <a href="https://ironsoftware.com/csharp/excel/how-to/clear-cells/">Clear</a> method, you can leverage the `Range` class to execute copy operations over any size and shape of data range:

- Copy a single cell `C10`: **workSheet ["C10"].Copy(workBook.GetWorkSheet("Sheet1"), "B13")**
- Copy an entire column `A`: **workSheet.GetColumn(0).Copy(workBook.GetWorkSheet("Sheet1"), "H1")**
- Copy a specific row, `Row 4`: **workSheet.GetRow(3).Copy(workBook.GetWorkSheet("Sheet1"), "A15")**
- Copy a two-dimensional range `D6:F8`: **workSheet ["D6:F8"].Copy(workBook.GetWorkSheet("Sheet1"), "H17")**

The target starting location for the copied data is specified by the second parameter, where the data expands rightward and downward from.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Execute various copy commands
workSheet["C10"].Copy(workBook.GetWorkSheet("Sheet1"), "B13");
workSheet.GetColumn(0).Copy(workBook.GetWorkSheet("Sheet1"), "H1");
workSheet.GetRow(3).Copy(workBook.GetWorkSheet("Sheet1"), "A15");
workSheet["D6:F8"].Copy(workBook.GetWorkSheet("Sheet1"), "H17");

workBook.SaveAs("copiedCellRange.xlsx");
```

### Resulting Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/copy-cells/copy-cells-copy-cell-range.png" alt="Copied Cell Range" class="img-responsive add-shadow">
    </div>
</div>

## Example: Copying Cells Across Different Worksheets

Since the `Copy` method also accepts worksheet objects, you can easily transfer cell contents across various worksheets. Here, we designate a different worksheet as the initial argument for the method.

In this instance, using `Sheet2`:

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Copy content to another worksheet
workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet2"), "B3");

workBook.SaveAs("copyBetweenWorksheets.xlsx");
```