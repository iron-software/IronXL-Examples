# How to Create and Utilize Named Ranges in Excel

***Based on <https://ironsoftware.com/how-to/named-range/>***


Named ranges provide a convenient way to refer to groups of cells within an Excel spreadsheet by using a distinct identifier, rather than standard cell coordinates like A1:B10. For instance, a named range could be called "SalesData" and used in functions such as `SUM(SalesData)`, enhancing clarity and simplifying formula applications.

## Example of Adding a Named Range

To create a named range, utilize the `AddNamedRange` method within IronXL by providing a name for the range and the range itself.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Define a cell range
            var selectedRange = workSheet["A1:A5"];
            
            // Assign a name to the range
            workSheet.AddNamedRange("range1", selectedRange);
            
            workBook.SaveAs("addNamedRange.xlsx");
        }
    }
}
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/named-range/named-range.webp" alt="Named Range" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Fetching Named Ranges

### Retrieve All Named Ranges

You can retrieve all named ranges present in a worksheet by using the `GetNamedRanges` method, which delivers a list of range names.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Access all named ranges
            var namedRangeList = workSheet.GetNamedRanges();
        }
    }
}
```

### Access a Specific Named Range

To fetch a specific named range, employ the `FindNamedRange` method for acquiring its absolute reference. This reference can be used to locate or refer to the named range directly.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Obtain the address of the named range
            string namedRangeAddress = workSheet.FindNamedRange("range1");
            
            // Utilize the named range
            var range = workSheet[$"{namedRangeAddress}"];
        }
    }
}
```

<hr>

## Removing a Named Range

To eliminate a named range from a worksheet, use the `RemoveNamedRange` method by specifying the name of the range to be removed.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Removal of the named range
            workSheet.RemoveNamedRange("range1");
        }
    }
}
```