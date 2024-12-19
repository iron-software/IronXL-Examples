# How to Create Hyperlinks in Excel

***Based on <https://ironsoftware.com/how-to/hyperlinks/>***


Excel hyperlinks are interactive elements that enable users to jump to different locations in the workbook, access various files, navigate to web pages, or compose emails. These features enhance the user experience by simplifying access to related data and external references. The addition of hyperlinks makes spreadsheets more dynamic and user-friendly, streamlining the interaction with additional information or external resources.

IronXL supports the addition of hyperlinks to URLs, the opening of external files from local and FTP file systems, email addresses, specific cell addresses, and named cells all without requiring Interop within .NET C#.

### Getting Started with IronXL

----------------------------------

## Example of Creating a Link Hyperlink

In the **Cell** class of IronXL, there is a **Hyperlink** property. Accessing the cell with `workSheet["A1"]` returns a **Range** object, which allows the use of the `First` method to fetch the initial cell in the range.

Alternatively, the `GetCellAt` method provides direct access to a cell, making it straightforward to manipulate its **Hyperlink** property.

Let us delve into an example of creating link hyperlinks, supporting both HTTP and HTTPS protocols.

Note: Utilizing the `GetCellAt` method on an uninitialized cell will lead to a *System.NullReferenceException: 'Object reference not set to an instance of an object.'*

```cs
using IronXL;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Assigning content to the cell
workSheet["A1"].Value = "Link to ironpdf.com";

// Establishing a hyperlink at cell A1 pointing to https://ironpdf.com/
workSheet.GetCellAt(0, 0).Hyperlink = "https://ironpdf.com/";

workBook.SaveAs("setLinkHyperlink.xlsx");
```

### Demonstration

![Link Hyperlink](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-link-hyperlink.gif)

## Example of Creating Hyperlinks Across Worksheets

For hyperlinks within the same worksheet, simply use the cell's address like Z20. For linking across different worksheets, utilize the convention "worksheetName!address". For instance, "Sheet2!A1".

Named cells can be scoped globally (workbook) or locally (worksheet). Hyperlinks to named cells within the same sheet or those with global scope can directly use the name. If the named cell has a local scope on a different sheet, specify the worksheet name as illustrated previously.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet1 = workBook.CreateWorkSheet("Sheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("Sheet2");

// Global named range creation
workSheet1["D5"].SaveAsNamedRange("Iron", true);

// Local named range creation
workSheet2["D10"].SaveAsNamedRange("Hello", false);

// Hyperlink within the same worksheet to cell Z20
workSheet1["A1"].Value = "Z20";
workSheet1["A1"].First().Hyperlink = "Z20";

// Hyperlink to the named range "Iron"
workSheet1["A2"].Value = "Iron";
workSheet1["A2"].First().Hyperlink = "Iron";

// Hyperlink to cell A1 of Sheet2
workSheet1["A3"].Value = "A1 of Sheet2";
workSheet1["A3"].First().Hyperlink = "Sheet2!A1";

// Hyperlink to the named range "Hello" on Sheet2
workSheet1["A4"].Value = "Define name Hello of Sheet2";
workSheet1["A4"].First().Hyperlink = "Sheet2!Hello";

workBook.SaveAs("setHyperlinkAcrossWorksheet.xlsx");
```

### Demonstration

![Hyperlink Across Worksheet](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-hyperlink-across-worksheet.gif)

## Example of Creating Various Hyperlink Types

IronXL facilitates the creation of other types of hyperlinks such as FTP, file, and email links.

- **FTP**: Begin with **ftp://**
- **File**: Use an absolute path that begins with **file:///**
- **Email**: Start with **mailto:**

Both FTP and File link types require full paths.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet = workBook.DefaultWorkSheet;

// FTP hyperlink to open a file
workSheet["A1"].Value = "Open sample.xlsx";
workSheet["A1"].First().Hyperlink = "ftp://C:/Users/sample.xlsx";

// File hyperlink to open a file
workSheet["A2"].Value = "Open sample.xlsx";
workSheet["A2"].First().Hyperlink = "file:///C:/Users/sample.xlsx";

// Email hyperlink creation
workSheet["A3"].Value = "example@gmail.com";
workSheet["A3"].First().Hyperlink = "mailto:example@gmail.com";

workBook.SaveAs("setOtherHyperlink.xlsx");
```

### Demonstration

![Other Types of Hyperlinks](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-other-hyperlink.gif)