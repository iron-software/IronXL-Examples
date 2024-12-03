# How to Duplicate Cell Contents

***Based on <https://ironsoftware.com/how-to/copy-cells/>***


The feature for duplicating cell contents enables you to copy the data from one cell and paste it into another cell or multiple cells. This tool is especially useful for copying data, formulas, formatting, and other cell attributes across the spreadsheet.

Moreover, the `Copy` method preserves the original cell's formatting, making it a robust tool for duplicating data with precision across single or multiple sheets using IronXL.

## Example of Copying a Single Cell

To duplicate the content of a single cell, utilize the `Copy` method. Provide the worksheet instance as the initial argument and the target cell as the second argument. This method also ensures that all original formatting is carried over.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CopyCells
{
    public class SingleCellCopy
    {
        public void Execute()
        {
            WorkBook workbook = WorkBook.Load("sample.xlsx");
            WorkSheet worksheet = workbook.GetWorkSheet("Sheet1");
            
            // Duplicate content of a cell
            worksheet["A1"].Copy(workbook.GetWorkSheet("Sheet1"), "B3");
            
            workbook.SaveAs("duplicatedSingleCell.xlsx");
        }
    }
}
```

### Output Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper" width='70%'>
         <img src="https://ironsoftware.com/static-assets/excel/how-to/copy-cells/copy-cells-copy-single-cell.png" alt="Duplicated Single Cell" class="img-responsive add-shadow">
    </div>
</div>

## Example of Copying a Cell Range

Copying a cell range operates similarly to the <a href="https://ironsoftware.com/csharp/excel/how-to/clear-cells/">Clear</a> method in the **Range** class, permitting its use across any size of cell range. Below are some examples:

Copying operations include:
- Single cell copying (C10): **worksheet["C10"].Copy(workbook.GetWorkSheet("Sheet1"), "B13")**
- Entire column copying (A): **worksheet.GetColumn(0).Copy(workbook.GetWorkSheet("Sheet1"), "H1")**
- Row copying (4): **worksheet.GetRow(3).Copy(workbook.GetWorkSheet("Sheet1"), "A15")**
- Two-dimensional range copying (D6:F8): **worksheet["D6:F8"].Copy(workbook.GetWorkSheet("Sheet1"), "H17")**

Each destination is indicated with an address that serves as the starting point for the copied data.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CopyCells
{
    public class RangeCopy
    {
        public void Execute()
        {
            WorkBook workbook = WorkBook.Load("sample.xlsx");
            WorkSheet worksheet = workbook.GetWorkSheet("Sheet1");
            
            // Execute several types of copy operations
            worksheet["C10"].Copy(workbook.GetWorkSheet("Sheet1"), "B13");
            worksheet.GetColumn(0).Copy(workbook.GetWorkSheet("Sheet1"), "H1");
            worksheet.GetRow(3).Copy(workbook.GetWorkSheet("Sheet1"), "A15");
            worksheet["D6:F8"].Copy(workbook.GetWorkSheet("Sheet1"), "H17");
            
            workbook.SaveAs("duplicatedCellRange.xlsx");
        }
    }
}
```

### Output Spreadsheet

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/copy-cells/copy-cells-copy-cell-range.png" alt="Copied Cell Range" class="img-responsive add-shadow">
    </div>
</div>

## Example of Copying Cells Across Different Worksheets

By providing different worksheet objects as the first parameter to the `Copy` method, it is possible to copy and paste cell contents across various worksheets. Below is an example:

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CopyCells
{
    public class AcrossSheetsCopy
    {
        public void Execute()
        {
            WorkBook workbook = WorkBook.Load("sample.xlsx");
            WorkSheet worksheet = workbook.GetWorkSheet("Sheet1");
            
            // Copy content to a different worksheet
            worksheet["A1"].Copy(workbook.GetWorkSheet("Sheet2"), "B3");
            
            workbook.SaveAs("copiedAcrossWorksheets.xlsx");
        }
    }
}
```