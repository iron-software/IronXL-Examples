# How to Select Range

***Based on <https://ironsoftware.com/how-to/select-range/>***


IronXL provides a straightforward approach for selecting and manipulating ranges within an Excel Worksheet, without the need for Office Interop.

## Select Range Example

Utilizing IronXL, you're able to carry out various operations on selected ranges including, but not limited to, [sorting](https://ironsoftware.com/csharp/excel/how-to/sort-cells/), calculations, and aggregation tasks.

When you execute methods that alter or reposition cell values, the targeted range, row, or column will refresh to show the updated values. IronXL enables the combination of multiple `IronXL.Ranges.Range` objects using the '+' operator.

### Select a Range

To target a range from cell **A2** to **B8**, utilize the following code:

```cs
using IronXL.Excel;  
using System.Linq;

namespace ironxl.SelectRange
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Accessing range from worksheet
            var range = workSheet["A2:B8"];
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-range.png" alt="Select Range" class="img-responsive add-shadow">
    </div>
</div>

### Select a Row

To select the 4th row, utilize the `GetRow(3)` method, which employs zero-based indexing. The range will cover any empty cells that are filled in other rows within the same column, ensuring completeness of the selected row across the board.

```cs
using IronXL.Excel;
using System.Linq;

namespace ironxl.SelectRange
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Retrieve a specific row from the worksheet
            var row = workSheet.GetRow(3);
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-row.png" alt="Select Row" class="img-responsive add-shadow">
    </div>
</div>

### Select a Column

To access column C, either utilize the `GetColumn(2)` method or directly refer to the column with `workSheet ["C:C"]`. Similar to the `GetRow` method, the selected column will include all related cells across rows.

```cs
using IronXL.Excel;
using System.Linq;

namespace ironxl.SelectRange
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Access a specific column from the worksheet
            var column = workSheet.GetColumn(2);
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-column.png" alt="Select Column" class="img-responsive add-shadow">
    </div>
</div>

### Combine Ranges

IronXL also allows for the merging of multiple `IronXL.Ranges.Range` objects. Using the '+' operator, you can seamlessly join ranges forming a new, consolidated range. However, direct combination of rows and columns with the '+' operator is unsupported.

Successfully combining ranges will update the original range object. For instance, in the following snippet, the `range` variable will reflect the newly combined ranges.

```cs
using IronXL.Excel;
using System.Linq;

namespace ironxl.SelectRange
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Initially accessing a specific range
            var range = workSheet["A2:B2"];
            
            // Merging ranges
            var combinedRange = range + workSheet["A5:B5"];
        }
    }
}
```

All row- and column-related index positions adhere to zero-based indexing.