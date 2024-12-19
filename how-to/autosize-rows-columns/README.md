# How to Automatically Adjust Row and Column Sizes

***Based on <https://ironsoftware.com/how-to/autosize-rows-columns/>***


Adjusting row and column sizes efficiently in a spreadsheet contributes to improved readability and space optimization. **IronXL**, a robust C# library, simplifies the process by enabling automatic resizing of rows and columns in .NET environments. This automation replaces the tedious manual adjustments typically required in spreadsheets.

## Initialize IronXL for Use

### Automatically Adjusting Row Heights

The `AutoSizeRow` method dynamically adjusts the height of the designated row based on its content, enhancing readability.

```cs
using IronXL;

// Opening an existing spreadsheet
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Automatically resizing the second row
worksheet.AutoSizeRow(1);

workbook.SaveAs("autoResize.xlsx");
```

#### Visualization

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-rows.png" alt="Auto Resize Row Example" class="img-responsive add-shadow">
    </div>
</div>

### Automatically Adjusting Column Widths

Implement the `AutoSizeColumn` method to dynamically modify the width of a column based on the length of its content.

```cs
using IronXL;

// Opening an existing spreadsheet
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Automatically resizing column A
worksheet.AutoSizeColumn(0);

workbook.SaveAs("autoResizeColumn.xlsx");
```

#### Visualization

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-columns.png" alt="Auto Resize Column Example" class="img-responsive add-shadow">
    </div>
</div>

All rows and columns are zero-indexed.

<hr>

## Enhanced Automatic Row Resizing

The enhanced `AutoSizeRow` method also accepts a **Boolean** parameter enabling the adjustment of row heights in merged cell scenarios based on the content's dominant height.

```cs
using IronXL;

// Opening an existing spreadsheet
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Applying enhanced auto resizing to specific rows
worksheet.AutoSizeRow(0, true);
worksheet.AutoSizeRow(1, true);
worksheet.AutoSizeRow(2, true);

workbook.SaveAs("advanceAutoResizeRow.xlsx");
```

### Demonstrative Example

Consider a case where the content's height is **192 pixels** spread across **3 rows** in a merged region. Using the enhanced autosize feature, each row adjusts to **64 pixels** in height.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-advance-rows-true.png" alt="Enhanced Auto Resize Row" class="img-responsive add-shadow">
    </div>
</div>

### Without Merged Cell Consideration

If `false` is specified, the `AutoSizeRow` method adjusts heights based solely on the tallest cell's content within unmerged cells.

```cs
using IronXL;

worksheet.Merge("A1:A3");

worksheet.AutoSizeRow(0, false);
worksheet.AutoSizeRow(1, false);
worksheet.AutoSizeRow(2, false);
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-advance-rows-false.png" alt="Auto Resize Row without Merged Cell Consideration" class="img-responsive add-shadow">
    </div>
</div>

## Enhanced Column Width Adjustment

Similarly, the `AutoSizeColumn` method can be tailored to account for multiple merged columns' content, distributing the width adjustment accordingly.

```cs
using IronXL;

// Loading an existing file
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Adjusting width per column in merged areas
worksheet.AutoSizeColumn(0, true);
worksheet.AutoSizeColumn(1, true);
worksheet.AutoSizeColumn(2, true);

workbook.SaveAs("advanceAutoResizeColumn.xlsx");
```

### Demonstrative Example

For a scenario where the content width is **117 pixels** across **2 columns**, each column is resized to **59 pixels**.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-advance-columns-true.png" alt="Enhanced Auto Resize Column" class="img-responsive add-shadow">
    </div>
</div>

### Ignoring Merged Cell Width

Setting `false` for the merged cell parameter results in column widths set solely based on the content width of unmerged cells.

```cs
worksheet.Merge("A1:B1");

worksheet.AutoSizeColumn(0, false);
worksheet.AutoSizeColumn(1, false);
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-advance-columns-false.png" alt="Auto Resize Column without Merged Cell Width Adjustment" class="img-responsive add-shadow">
    </div>
</div>

Comparisons between **IronXL** and Excel show distinct padding differences, with Excel applying more noticeable padding around cell contents.

Furthermore, manual adjustments of row and column sizes cater to specialized requirements: simply set the `Height` and `Width` properties on `RangeRow` and `RangeColumn` respectively.

The unit of measurement for row heights is 1/20 of a point and for column widths, it is determined by the number of zeros that can fit horizontally in a cell using the normal font. IronXL aligns well with Excel's measurement standards, though with slight differences due to its unique calculation method, ensuring that users can flexibly adjust dimensions in spreadsheet documents as needed.