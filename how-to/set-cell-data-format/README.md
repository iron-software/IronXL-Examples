# How to Set Cell Data Formats

Manipulating Excel files to display cell data in specific formats can greatly enhance clarity and promote the accuracy of the content displayed. By applying data and number formats, you can control the appearance of numbers, dates, times, and other forms of data within a spreadsheet. These adjustments help in presenting data effectively, catering to specifics such as currencies, percentages, and customized numeric precision.

The IronXL library provides a seamless way to apply these formatting techniques in C#. This toolkit assists developers in generating, customizing, and managing Excel files programmatically—a monumental asset for software applications that manage considerable data operations.

## Example: Configuring Cell Data Formats

IronXL empowers developers to format cells and ranges efficiently using the **FormatString** attribute. This feature is versatile and can be used to format individual cells, entire columns or rows, or larger ranges.

```cs
using IronXL;
using System;

// Instantiate a new Excel workbook
WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply percentage format to cell A1
workSheet["A1"].Value = 123;
workSheet["A1"].FormatString = BuiltinFormats.Percent2;

// Apply custom numeric format to cell A2
workSheet["A2"].Value = 123;
workSheet["A2"].FormatString = "0.0000";

// Set and format a range of cells with date time values
DateTime dateValue = new DateTime(2020, 1, 1, 12, 12, 12);
workSheet["A3"].Value = dateValue;
workSheet["A4"].Value = new DateTime(2022, 3, 3, 10, 10, 10);
workSheet["A5"].Value = new DateTime(2021, 2, 2, 11, 11, 11);

IronXL.Range range = workSheet["A3:A5"];
range.FormatString = "MM/dd/yy h:mm:ss";

workBook.SaveAs("dataFormats.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/set-cell-data-format/data-format.png" alt="Data Format" class="img-responsive add-shadow">
    </div>
</div>

### Assigning Cell Values as String Directly

When setting cell values in IronXL, you can use **StringValue** to assign a text string directly to a cell. This approach is akin to prefacing the cell value with an apostrophe in Excel to avoid automatic conversion.

```cs
// Directly assign string value to cell
workSheet["A1"].StringValue = "4402-12";
```

## Utilizing Built-In Formats

IronXL features a collection of predefined format strings, housed within the **IronXL.Formatting.BuiltinFormats** class. These allow for detailed customization of how data is rendered in Excel documents.

```cs
using IronXL;
using IronXL.Formatting;

// Create and format a new workbook
WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Format cell A1 as accounting format
workSheet["A1"].Value = 123;
workSheet["A1"].FormatString = BuiltinFormats.Accounting0;

workBook.SaveAs("builtinDataFormats.xlsx");
```

### Overview of Built-In Data Formats

Below is a detailed summary of the built-in formats available for various data types, enabling developers to choose the optimal presentation style for data ranging from durations and times to numeric and text formats for specific contexts:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/set-cell-data-format/all-available-data-formats.png" alt="All Available Data Formats" class="img-responsive add-shadow">
    </div>
</div>

#### Descriptions of Available Formats

- **General**: Shows numbers as entered. No specific format applied.
- **Duration1 through Duration3**: Formats time as minutes and seconds, or hours, minutes, seconds, and even milliseconds.
- **Accounting0 through Accounting2Red**: Presents financial data with varying precision and color highlighting for negative numbers.
- **Time1 through Time4**: Varies in showing hours and minutes with or without the inclusion of seconds, in either 12-hour or 24-hour format.
- **ShortDate through LongDate3**: Ranges from short date representations to more extensive formats includ...
- **Fraction1 and Fraction2**: Depicts numerical values as fractions with one or two-digit denominators.
- **Scientific1 and Scientific2**: Uses scientific notation for displaying numbers.
- **Percent and Percent2**: Count values as percentages with or without decimal places.
- **Currency0 through Currency2Red**: Presents currency formats adjusting digits and negative representation in default or red color.
- **Thousands0 and Thousands2**: Numbers are formatted with thousand separators.
- **Number0 and Number2**: Standard number formats with the inclusion of zero or two decimal places.
- **Text**: Plain text format without any applied numeric or date conversions.