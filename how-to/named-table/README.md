# Adding a Named Table in Excel

***Based on <https://ironsoftware.com/how-to/named-table/>***


A named table, often referred to as an Excel Table, is a specific range in Excel that is designated with a name and possesses enhanced properties and functionalities.

## Example: Creating a Named Table

To create a named table, utilize the `AddNamedTable` method. This method necessitates specifying the table's name as a string, its range, and optionally, its style and whether to include a filter.

```cs
using IronXL.Styles;
using IronXL.Excel;

namespace ironxl.NamedTable
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Populate data
            workSheet["A2:C5"].StringValue = "Text";

            // Define the range for the table
            var selectedRange = workSheet["A1:C5"];
            bool showFilter = false;
            var tableStyle = TableStyle.TableStyleDark1;

            // Create the named table
            workSheet.AddNamedTable("table1", selectedRange, showFilter, tableStyle);

            // Save the workbook
            workBook.SaveAs("addNamedTable.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/named-table/named-table.webp" alt="Named Table" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Example: Retrieving Named Tables

### Retrieve All Named Tables

The `GetNamedTableNames` method retrieves all named tables in a worksheet, returning them as a list of strings.

```cs
using IronXL;
using IronXL.Excel;

namespace ironxl.NamedTable
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Fetch all named tables
            var namedTableList = workSheet.GetNamedTableNames();
        }
    }
}
```

### Retrieve a Specific Named Table

Retrieve an individual named table using the `GetNamedTable` method.

```cs
using IronXL;
using IronXL.Excel;

namespace ironxl.NamedTable
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Access a specific named table
            var namedRangeAddress = workSheet.GetNamedTable("table1");
        }
    }
}
```

IronXL also supports adding named ranges. Learn more at [How to Add Named Range](https://ironsoftware.com/csharp/excel/how-to/named-range/).