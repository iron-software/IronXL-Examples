# How to Copy Cells

***Based on <https://ironsoftware.com/how-to/copy-cells/>***


The "Copy cell" functionality enables you to clone the content of a cell, allowing you to transfer it to another cell or multiple cells. This feature is very useful for duplicating data, formulas, formatting, or other elements across your spreadsheet.

Furthermore, the `Copy` method not only copies the data but also preserves the formatting, ensuring data consistency across single or multiple worksheets through IronXL.

<h3>Getting Started with IronXL</h3>

---

## Example of Copying a Single Cell

To duplicate the contents of a specific cell, the `Copy` method should be utilized. This method requires the worksheet object as the initial parameter and the target cell's starting position as the second parameter. This method efficiently maintains all existing cell styling.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Duplicate the content of cell A1 to B3
workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet1"), "B3");

workBook.SaveAs("copySingleCell.xlsx");
```

### Output of the Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper" width='70%'>
         <img src="https://ironsoftware.com/static-assets/excel/how-to/copy-cells/copy-cells-copy-single-cell.png" alt="Copy Single Cell" class="img-responsive add-shadow">
    </div>
</div>

## Example of Copying a Cell Range

Just like the <a href="https://ironsoftware.com/csharp/excel/how-to/clear-cells/">Clear</a> method, the **Range** class offers a similar copying function which can be applied to a single cell, a column, a row, or a specific cell block, regardless of the range's size. Here’s how it operates:

Duplicate a single cell (C10):
- **workSheet["C10"].Copy(workBook.GetWorkSheet("Sheet1"), "B13")**

Duplicate an entire column (A):
- **workSheet.GetColumn(0).Copy(workBook.GetWorkSheet("Sheet1"), "H1")**

Duplicate a row (4):
- **workSheet.GetRow(3).Copy(workBook.GetWorkSheet("Sheet1"), "A15")**

Duplicate a block (D6:F8):
- **workSheet["D6:F8"].Copy(workBook.GetWorkSheet("Sheet1"), "H17")**

The target location is specified as the second parameter, dictating the starting point for the pasted data.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Copy single cell C10 to B13
workSheet["C10"].Copy(workBook.GetWorkSheet("Sheet1"), "B13");

// Copy column A to column H
workSheet.GetColumn(0).Copy(workBook.GetWorkSheet("Sheet1"), "H1");

// Copy row 4 to row 15
workSheet.GetRow(3).Copy(workBook.GetWorkSheet("Sheet1"), "A15");

// Copy block D6:F8 to H17
workSheet["D6:F8"].Copy(workBook.GetWorkSheet("Sheet1"), "H17");

workBook.SaveAs("copyCellRange.xlsx");
```

### Output of the Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/copy-cells/copy-cells-copy-cell-range.png" alt="Copy Cell Range" class="img-responsive add-shadow">
    </div>
</div>

## Example of Copying Cells Across Different Worksheets

You can also copy and paste content across different worksheets by passing different worksheet objects as the first parameter. Below is an example where content from one worksheet ("Sheet1") is copied to another ("Sheet2").

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Transfer cell content from Sheet1 to Sheet2
workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet2"), "B3");

workBook.SaveAs("copyAcrossWorksheet.xlsx");
```