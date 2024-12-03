# How to Trim Cell Range

***Based on <https://ironsoftware.com/how-to/trim-cell-range/>***


The IronXL library facilitates the removal of all blank rows and columns surrounding the designated area in C# without requiring Office Interop. This capability greatly enhances data handling and manipulation by eliminating the need to interface with the Office suite directly.

## Trim Cell Range Example

To trim a cell range, first select the <a href="https://ironsoftware.com/csharp/excel/how-to/select-range/">desired range</a> and use the `Trim` method on it. This function clears away the leading and trailing empty cells from your chosen range.

Keep in mind, the `Trim` method will not eliminate any empty cells that might exist between rows or columns inside the range. To tackle this, consider <a href="https://ironsoftware.com/csharp/excel/how-to/sort-cells/">sorting</a> the range which will shift the empty cells to the top or bottom of the range.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.TrimCellRange
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Setting values
            workSheet["A2"].Value = "A2";
            workSheet["A3"].Value = "A3";
            
            workSheet["B1"].Value = "B1";
            workSheet["B2"].Value = "B2";
            workSheet["B3"].Value = "B3";
            workSheet["B4"].Value = "B4";
            
            // Retrieve the first column
            RangeColumn column = workSheet.GetColumn(0);
            
            // Trim the column and store in a new range variable
            Range trimmedColumn = column.Trim();
        }
    }
}
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/trim-cell-range/trim-cell-range-column.png" alt="Trim Column" class="img-responsive add-shadow">
    </div>
</div>
<hr>