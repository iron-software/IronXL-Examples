# How to Sort Cell Ranges

***Based on <https://ironsoftware.com/how-to/sort-cells/>***


Sorting data by either alphabetical or numerical order is crucial for proper data analysis in Microsoft Excel. Utilizing IronXL, developers can efficiently sort Excel columns, rows, or specified cell ranges using C# and VB.NET.

## Example of Independently Sorting Columns

You can implement sorting by using the `SortAscending` or `SortDescending` methods on the desired range or column. When sorting a range containing multiple columns, these methods will sort each column independently.

To deal with any resultant empty cells at the top or bottom of your sorted range, utilize the [Trim](https://ironsoftware.com/csharp/excel/how-to/trim-cell-range/) method post-sort. This cleans up your data set by removing these empties.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.SortCells
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Select and sort a column (A) ascendingly
            var column = workSheet.GetColumn(0);
            column.SortAscending();  // From A to Z
            
            // Now sort the same column descendingly
            column.SortDescending(); // From Z to A
            
            workBook.SaveAs("sortedExcelRange.xlsx");
        }
    }
}
```
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/sort-cells/sort-cells-range.png" alt="Sort Ascending and Descending" class="img-responsive add-shadow">
    </div>
</div>

---

## Example of Sorting by a Specific Column

Utilizing the `SortByColumn` method, you can sort a cell range based on a specific column’s values. This method requires the specification of the column for sorting and the range that the sort will be applied to.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.SortCells
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Define a range and sort it according to column (B)
            var range = workSheet["A1:D10"];
            range.SortByColumn("B", SortOrder.Ascending);  // Sort range using Column B
            
            workBook.SaveAs("sortedByColumnRange.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/sort-cells/sort-cells-sort-by-column.png" alt="Sort by Specific Column" class="img-responsive add-shadow">
    </div>
</div>

Please note that current functionality does not support multi-column sequential sorting, such as sorting first by column A, then by column B.