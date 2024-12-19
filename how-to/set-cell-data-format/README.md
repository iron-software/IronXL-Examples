# How to Apply Cell Data Formats

***Based on <https://ironsoftware.com/how-to/set-cell-data-format/>***


Formulating cell data and number formats in Excel empowers users to influence the visual representation of numbers, dates, times, and other data types. By employing specific formats like currency or percentage, you can not only enhance visual comprehension but also uphold data precision. Data formats ensure that information is represented in your preferred style, while number formats allow familiarity in expressing numerical data with varied decimal and display preferences.

Utilizing the IronXL library, you can implement these data or number formats conveniently in C#. This library streamlines the tasks of generating, formatting, and manipulating Excel documents through code, proving essential for efficient data management and presentation in C# applications.

### Begin with IronXL

-------------------------------------

## Example: Setting Cell Data Formats

Access the `FormatString` property via cells or ranges, permitting the configuration of data formats for individual cells or groupings like columns, rows, or any specified range.

```cs
using IronXL;
using IronXL.Formatting;
using System;
using System.Linq;

// Initialize a new workbook
WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Apply percent format to the cell
worksheet["A1"].Value = 123;
worksheet["A1"].FormatString = BuiltinFormats.Percent2;

// Set a custom numeric format
worksheet["A2"].Value = 123;
worksheet["A2"].FormatString = "0.0000";

// Format a range to show date and time
DateTime customDate = new DateTime(2020, 1, 1, 12, 12, 12);
worksheet["A3"].Value = customDate;
worksheet["A4"].Value = new DateTime(2022, 3, 3, 10, 10, 10);
worksheet["A5"].Value = new DateTime(2021, 2, 2, 11, 11, 11);

IronXL.Range dateRange = worksheet["A3:A5"];
dateRange.FormatString = "MM/dd/yy h:mm:ss";

workbook.SaveAs("formattedData.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/set-cell-data-format/data-format.png" alt="Data Format" class="img-responsive add-shadow">
    </div>
</div>

### Setting Cell Values as Strings

To place a text value into a cell without converting its type, use the `StringValue` property. This is akin to entering text in Excel prefixed by an apostrophe.

```cs
// Directly assign string value
worksheet["A1"].StringValue = "4402-12";
```

## Example of Using Built-In Formats

IronXL provides a range of built-in format strings accessible through the  `IronXL.Formatting.BuiltinFormats` class, facilitating customized data displays in your Excel sheets.

```cs
using IronXL;
using IronXL.Formatting;

// Start a new workbook
WorkBook newWorkbook = WorkBook.Create();
WorkSheet newWorksheet = newWorkbook.DefaultWorkSheet;

// Demonstrate the use of built-in format
newWorksheet["A1"].Value = 123;
newWorksheet["A1"].FormatString = BuiltinFormats.Accounting0;

newWorkbook.SaveAs("builtinFormatsDemo.xlsx");
```

### Overview of Predefined Data Formats

Here's a summary of format types available for time durations, accounting, time display, date styles, scientific notation, percentages, currency, number formatting, and text:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/set-cell-data-format/all-available-data-formats.png" alt="All Available Data Formats" class="img-responsive add-shadow">
    </div>
</div>

#### Descriptions

- **General**: Displays values without any formatting applied, just as they are.
- **Duration Formats**: Range from presenting minutes and seconds to displaying hours, minutes, and milliseconds.
- **Accounting Formats**: Show financial figures in various styles, either with or without decimal places, and optionally highlighting negative values in red.
- **Time Formats**: Available in both 12-hour and 24-hour settings, with or without seconds.
- **Date Formats**: Include short and long styles, with options to display the year, month, and day in various combinations.
- **Fractional Formats**: Allow representation of numbers as one or two-digit fractions.
- **Scientific Formats**: Present numbers in scientific notation, differing by the number of decimals.
- **Percentage**: Showcase values as percentages, with the choice to include decimals.
- **Currency**: Specify monetary values in several styles, highlighting negative amounts and choosing decimal precision.
- **Number Formats**: General number displays that can be adjusted for thousands separation and decimal places.
- **Text**: Treats cell values as plain text, applying no specific formatting.

By mastering these formatting options, your Excel files become more functional and tailored to your specific data presentation requirements.