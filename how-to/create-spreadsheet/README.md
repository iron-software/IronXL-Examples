# How to Generate New Excel Files

***Based on <https://ironsoftware.com/how-to/create-spreadsheet/>***


XLSX is a contemporary file format utilized for storing Microsoft Excel spreadsheets. It adheres to the Open XML standard, which was introduced in Office 2007. The XLSX format is equipped to handle sophisticated functions such as charts and conditional formatting, making it highly suitable for data analysis and various business applications.

On the other hand, XLS is the older binary format for Excel files, predominant in earlier software versions. It does not support the enhanced functionalities of XLSX and has become increasingly rare.

IronXL empowers users to generate both XLSX and XLS files effortlessly using a single line of code.

### Begin with IronXL

----------------------------------

## Example of Creating a Spreadsheet

To create an Excel workbook, which serves as a container for one or more sheets, utilize the static `Create` method. This method by default constructs a workbook in the XLSX format.

```cs
using IronXL;

// Initialize a new spreadsheet
WorkBook workBook = WorkBook.Create();
```

<hr>

## Selecting the Spreadsheet Format

The `Create` method can also employ an **ExcelFileFormat** enumeration to determine the type of file to generate, whether XLSX or XLS. XLSX is the modern, XML-based format launched with Office 2007, offering enhanced functionality and efficiency. In contrast, XLS is the antiquated binary format from earlier Excel versions, now less frequently used due to its limitations.

```cs
using IronXL;

// Generate an XLSX file
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
```

There's an additional variant of the `Create` method, which accepts **CreatingOptions** as its parameter. Presently, the **CreatingOptions** class includes a single property, `DefaultFileFormat`, used to decide whether to produce an XLSX or XLS file. Refer to the following example for details:

```cs
using IronXL;

// Create an XLSX file with specific options
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
```