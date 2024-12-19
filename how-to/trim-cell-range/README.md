# How to Trim Cell Range

***Based on <https://ironsoftware.com/how-to/trim-cell-range/>***


The IronXL library simplifies removal of all empty rows and columns at the borders of a range in C# without the need for Office Interop. This capability significantly enhances the efficiency of data handling and manipulation, sidestepping the need to interact with the Office suite directly.

### Getting Started with IronXL

---

## Example: Trimming a Cell Range

Identify the specific <a href="https://ironsoftware.com/csharp/excel/how-to/select-range/">Range</a> of cells you want to modify and use the `Trim` method on it. This function cuts away the empty cells at the beginning and end of the selected range.

Note that the `Trim` method does not eliminate empty cells that are situated in the middle of the range across rows and columns. To organize these, you might consider <a href="https://ironsoftware.com/csharp/excel/how-to/sort-cells/">sorting</a> the cells, thereby moving the empty ones to either the top or bottom of the range.

```cs
using IronXL;

// Initializing workbook and worksheet
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Setting values in cells
worksheet["A2"].Value = "A2";
worksheet["A3"].Value = "A3";

worksheet["B1"].Value = "B1";
worksheet["B2"].Value = "B2";
worksheet["B3"].Value = "B3";
worksheet["B4"].Value = "B4";

// Fetch the first column
RangeColumn column = worksheet.GetColumn(0);

// Execute trimming on the column
Range trimmedRange = column.Trim();
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/trim-cell-range/trim-cell-range-column.png" alt="Trimmed Column" class="img-responsive add-shadow">
    </div>
</div>
<hr>