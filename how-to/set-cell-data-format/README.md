# How to Apply Data Formats to Cells in Excel

***Based on <https://ironsoftware.com/how-to/set-cell-data-format/>***


Data formatting in Excel provides the means to control the appearance of numbers, dates, times, and other contents within cells. This enhances both the clarity and the accuracy of the data presented. Specific formatting options such as currency or percentage formats help tailor the display to match data interpretation needs, while number formats give fine control over decimal and digit groupings.

IronXL, a library for .NET, facilitates the setting of both data and number formats directly in C#. It significantly eases the way developers can create, format, and manipulate Excel documents programmatically, making it an indispensable tool for any C# driven data management and presentation application.

## Example: Setting Cell Data Formats

You can utilize the `FormatString` property in IronXL to apply formats over cells, columns, rows, or ranges in Excel. Below is an example demonstrating this:

```cs
using IronXL.Excel;
using System.Linq;

namespace ironxl.SetCellDataFormat
{
    public class Example
    {
        public void Execute()
        {
            // Initialize a new workbook
            WorkBook workbook = WorkBook.Create();
            WorkSheet sheet = workbook.DefaultWorkSheet;
            
            // Apply percentage format
            sheet["A1"].Value = 123;
            sheet["A1"].FormatString = BuiltinFormats.Percent2;
            
            // Apply custom number format
            sheet["A2"].Value = 123;
            sheet["A2"].FormatString = "0.0000";
            
            // Apply date and time format across a range
            var startDateTime = new DateTime(2020, 1, 1, 12, 12, 12);
            sheet["A3"].Value = startDateTime;
            sheet["A4"].Value = new DateTime(2022, 3, 3, 10, 10, 10);
            sheet["A5"].Value = new DateTime(2021, 2, 2, 11, 11, 11);

            Range dateRange = sheet["A3:A5"];
            dateRange.FormatString = "MM/dd/yy h:mm:ss";
            
            workbook.SaveAs("FormattedData.xlsx");
        }
    }
}
```

### Set Cell Value as String

In scenarios where you need to input text exactly as it appears, use `StringValue` to bypass automatic data type conversion in Excel.

```cs
using IronXL.Excel;

namespace ironxl.SetCellDataFormat
{
    public class ExampleString
    {
        public void Execute()
        {
            // Directly assign a string
            workSheet["A1"].StringValue = "4402-12";
        }
    }
}
```

## Usage of Built-in Formats

IronXL offers several built-in format strings, which are readily available through `IronXL.Formatting.BuiltinFormats`, to standardize cell data display:

```cs
using IronXL.Excel;
using IronXL.Formatting;

namespace ironxl.SetCellDataFormat
{
    public class BuiltInFormatsExample
    {
        public void Execute()
        {
            // Initialize and format workbook
            WorkBook workbook = WorkBook.Create();
            WorkSheet sheet = workbook.DefaultWorkSheet;

            // Assigning a preset format
            sheet["A1"].Value = 123;
            sheet["A1"].FormatString = BuiltinFormats.Accounting0;
            
            workbook.SaveAs("PrebuiltFormats.xlsx");
        }
    }
}
```

### Comprehensive List of Data Formats

IronXL supports various data types, including durations, dates, accounting figures, time, and scientific notations. Here are some of the formats available:

- **General**: Displays numbers as entered.
- **Duration**: Shows time lengths with formats for minutes, seconds, and hours.
- **Accounting**: Number formats intended for financial records.
- **Time**: Multiple formats including both 12-hour and 24-hour clocks.
- **Date**: Short and long date formats.
- **Fraction**: Fractional representations.
- **Scientific**: Scientific notations.
- **Percent and Currency**: Formats for representing financial percentages and currency values.
- **Number**: Decimal and plain numerical formats.
- **Text**: Text format.

These formatting options can be utilized to ensure the data adheres to expected standards for display and interpretation.

![Data Formats Example](https://ironsoftware.com/static-assets/excel/how-to/set-cell-data-format/data-format.png)

![All Available Data Formats](https://ironsoftware.com/static-assets/excel/how-to/set-cell-data-format/all-available-data-formats.png)