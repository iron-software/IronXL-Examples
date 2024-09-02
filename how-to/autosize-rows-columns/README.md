# Managing Auto-Resize for Spreadsheet Rows and Columns

Automatically resizing rows and columns in a spreadsheet not only optimizes space utilization but also enhances readability. The **IronXL** C# library empowers developers with methods to automate resizing tasks for rows and columns directly within C#, greatly simplifying what would otherwise be a manual update process.

## Example: Automatically Resizing Rows

The `AutoSizeRow` method dynamically adjusts the height of a specified row based on its content:

```cs
using IronXL;

// Load an existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Automatically adjust the height of the second row
workSheet.AutoSizeRow(1);  // Note: Rows are zero-indexed

workBook.SaveAs("autoResize.xlsx");
```

### Demonstrating Auto-Resized Rows
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-rows.png" alt="Auto Resize Row" class="img-responsive add-shadow">
    </div>
</div>

## Example: Automatically Resizing Columns

Leverage the `AutoSizeColumn` method to adjust the width of a column based on the length of its content:

```cs
using IronXL;

// Load an existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Automatically adjust the width of column A
workSheet.AutoSizeColumn(0);  // Note: Columns are zero-indexed

workBook.SaveAs("autoResizeColumn.xlsx");
```

### Demonstrating Auto-Resized Columns
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-columns.png" alt="Auto Resize Column" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Auto Resizing for Rows

An advanced use of `AutoSizeRow` includes an additional Boolean parameter that, when set to `true`, considers the height of merged cells by dividing the height of the content in the top-left cell by the number of rows in the merge:

```cs
using IronXL;

// Load an existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply advanced auto resizing to individual rows
workSheet.AutoSizeRow(0, true);
workSheet.AutoSizeRow(1, true);
workSheet.AutoSizeRow(2, true);

workBook.SaveAs("advanceAutoResizeRow.xlsx");
```

### Advanced Row Resizing Example

For example, if a merged cell region spanning three rows contains content with a height of 192 pixels, adjusting any of these rows will result in a uniform height of 64 pixels per row:

### Additional Info on Non-Merged Resizing

When the boolean is set to `false`, the height is adjusted solely based on the tallest cell content in each row, which avoids extra computations involved with merged cells. Here's a visual demonstration:
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-advance-rows-false.png" alt="Advance Auto Resize Row" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Auto Resizing for Columns

Similarly, `AutoSizeColumn` can also be parameterized to account for the width of merged cells if required:

```cs
using IronXL;

// Load the spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply advanced auto resizing to columns
workSheet.AutoSizeColumn(0, true);
workSheet.AutoSizeColumn(1, true);
workSheet.AutoSizeColumn(2, true);

workBook.SaveAs("advanceAutoResizeColumn.xlsx");
```

### Visual Comparison for Auto Resizing Differences

Comparing IronXL with Excel shows how padding is applied in Excel's autofit features, which is absent in IronXL’s crisp alignment:

### Manual Adjustments for Height and Width

In scenarios where automatic adjustments do not meet specific requirements, IronXL allows for manual settings of row heights and column widths:

```cs
using IronXL;

// Open an existing file
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

RangeRow row = workSheet.GetRow(0);
row.Height = 10;  // Manually set height

RangeColumn col = workSheet.GetColumn(0);
col.Width = 10;  // Manually set width

workBook.SaveAs("manualHeightAndWidth.xlsx");
```

This flexibility ensures that IronXL can adapt to various content and styling requirements, making it an indispensable tool for developers managing Excel data in .NET environments.