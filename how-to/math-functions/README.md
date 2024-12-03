# Utilizing Mathematical Functions in IronXL

***Based on <https://ironsoftware.com/how-to/math-functions/>***


IronXL offers robust mathematical aggregation functions like Average, Sum, Min, and Max, vital for data calculations and analysis within Excel. These functions allow you to pull useful insights, support decision-making processes, and perform intricate numerical analyses directly in Excel spreadsheets, all without depending on Interop.

## Example of Using Aggregate Functions

While handling cell ranges in Excel files, IronXL empowers you to implement various aggregate functions to execute vital calculations. Below is an outline of important methods provided by IronXL:

- The `Sum()` method computes the total sum of the selected cell range.
- The `Avg()` method calculates the average value across the selected cell range.
- The `Min()` method finds the smallest number within the selected cell range.
- The `Max()` method determines the highest number within the selected cell range.

These methods are crucial for analyzing information and extracting significant insights from Excel data. Calculations exclude any non-numeric values.

```cs
using System.Linq;
using IronXL.Excel;

namespace ironxl.MathFunctions
{
    public class Section1
    {
        public void Run()
        {
            // Load an Excel workbook
            WorkBook workBook = WorkBook.Load("sample.xls");
            // Access the first worksheet
            WorkSheet workSheet = workBook.WorkSheets.First();

            // Define a range within the worksheet
            var range = workSheet["A1:A8"];
            
            // Compute the sum of numbers within the defined range
            decimal sum = range.Sum();
            
            // Compute the average of numbers in the range
            decimal avg = range.Avg();
            
            // Find the maximum number in the range
            decimal max = range.Max();
            
            // Determine the minimum number in the range
            decimal min = range.Min();
        }
    }
}
```

For enhanced adaptability, these functions are not limited to ranges; they can flexibly be applied across single or multiple rows and columns. Brush up on how to effectively [select rows and columns](https://ironsoftware.com/csharp/excel/how-to/select-range/).