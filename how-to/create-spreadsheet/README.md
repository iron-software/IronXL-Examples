# Generating New Excel Spreadsheets

XLSX is a contemporary file format used for storing Microsoft Excel spreadsheets, which adopts the Open XML standard that was introduced with Office 2007. This format is capable of supporting sophisticated features such as charts and conditional formatting, making it highly suitable for data analysis and various business functionalities.

On the other hand, XLS is the older version of Excel's file format, which operates on a binary system. Older versions of Excel utilized this format, but it doesn't include the upgraded capabilities found in XLSX and has been largely replaced in modern applications.

IronXL enables developers to create both XLSX and XLS files effortlessly, often requiring only a single line of code.

## Example of Spreadsheet Creation

To initiate a new Excel workbook, which can contain multiple sheets or worksheets, utilize the static `Create` method. This function, by default, constructs a workbook in the XLSX format.

```cs
using IronXL;

// Initialize a new spreadsheet
WorkBook workBook = WorkBook.Create();
```

<hr>

## Selecting the Spreadsheet Format

The `Create` method also supports the `ExcelFileFormat` enumeration, allowing developers to specify whether the spreadsheet should be in the XLSX or XLS format. XLSX represents the modernized, XML-based version introduced with Office 2007, while XLS is the more antiquated, binary format from previous versions. Generally, XLSX is preferred because of its enhanced features and greater efficiency.

```cs
using IronXL;

// Generating an XLSX formatted spreadsheet
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
```

There's an additional version of the `Create` method that accepts `CreatingOptions` as a parameter. The `CreatingOptions` class, however, primarily contains a single property, DefaultFileFormat, which determines whether to produce an XLSX or XLS file. Here's how you might use it:

```cs
using IronXL;

// Configure and create an XLSX formatted spreadsheet
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
```