# How to Automatically Adjust Row and Column Sizes

***Based on <https://ironsoftware.com/how-to/autosize-rows-columns/>***


Adjusting the sizes of rows and columns in a spreadsheet can enhance its readability and overall compactness. The **IronXL** C# library offers a straightforward functionality for automatically resizing rows and columns within your spreadsheets. Implemented in C#, these methods facilitate the automation of what is typically a manual adjustment process in spreadsheets.

## Example: Automatically Resizing Rows

The `AutoSizeRow` method dynamically adjusts the height of specified row(s) to fit their contents.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section1
    {
        public void Run()
        {
            // Load the spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Auto resize the second row
            workSheet.AutoSizeRow(1);
            
            // Save the modified spreadsheet
            workBook.SaveAs("autoResize.xlsx");
        }
    }
}
```

### Visual demonstration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-rows.png" alt="Auto Resize Row" class="img-responsive add-shadow">
    </div>
</div>

## Example: Automatically Resizing Columns

Leverage the `AutoSizeColumn` method to automatically adjust the width of column(s) according to the length of their content.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section2
    {
        public void Run()
        {
            // Load the spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Auto resize the first column
            workSheet.AutoSizeColumn(0);
            
            // Save the updated spreadsheet
            workBook.SaveAs("autoResizeColumn.xlsx");
        }
    }
}
```

### Visual demonstration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-columns.png" alt="Auto Resize Column" class="img-responsive add-shadow">
    }
</div>

## Advanced Auto-Resizing of Rows with Merged Cells

The `AutoSizeRow` method also supports an overload that includes a boolean parameter. This parameter, when set to `true`, includes merged cells by dividing the total height of the merged cell by the number of rows in the merge region.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section3
    {
        public void Run()
        {
            // Load the spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Apply auto resize individually to rows considering merged cells
            workSheet.AutoSizeRow(0, true);
            workSheet.AutoSizeRow(1, true);
            workSheet.AutoSizeRow(2, true);
            
            // Save the advanced resized spreadsheet
            workBook.SaveAs("advanceAutoResizeRow.xlsx");
        }
    }
}
```

### Example Explanation

Consider a merged region with a content height of **192 pixels** spanning **3 rows**. When resizing using the autosize feature, each row will have an adjusted height of **64 pixels**.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/autosize-rows-columns/autosize-rows-columns-advance-rows-true.png" alt="Advance Auto Resize Row" class="img-responsive add-shadow">
    </div>
</div>

### What happens when the `useMergedCells` parameter is false?

With `useMergedCells` set as `false`, `AutoSizeRow` will adjust the row height based pure