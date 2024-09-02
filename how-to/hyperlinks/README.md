# How to Implement Hyperlinks in Excel

Excel hyperlinks enable users to create accessible links to specific locations within the workbook, different documents, websites, or email addresses. They improve the navigational experience by offering shortcuts to relevant information and external sites, enhancing the interactivity and efficiency of spreadsheets by providing quick links to additional references or external websites.

IronXL facilitates the addition of hyperlinks in various forms, including URLs, the access to external files via local and FTP (File Transfer Protocol), email addresses, and addresses within a spreadsheet cell or named ranges, all without needing Interop in .NET C# applications.

## Example: Creating a URL Hyperlink

In IronXL, the **Hyperlink** attribute is accessible within the **Cell** class. Using the code `workSheet["A1"]`, you obtain a **Range** object. The `First` method allows you to interact with the first cell in that range.

Alternatively, the `GetCellAt` method lets you directly address the cell, which facilitates a straightforward access to the **Hyperlink** property.

In the following example, we demonstrate how to implement a hyperlink using both HTTP and HTTPS protocols. Note that trying to access a cell that hasn't been initialized will result in a *System.NullReferenceException* indicating no object reference set.

```cs
using IronXL;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Assign the cell a value
workSheet["A1"].Value = "Visit IronPDF";

// Set the A1 hyperlink to point to IronPDF's website
workSheet.GetCellAt(0, 0).Hyperlink = "https://ironpdf.com/";

workBook.SaveAs("setURLHyperlink.xlsx");
```

### Demonstration

![Link Hyperlink](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-link-hyperlink.gif)

## Example: Creating a Cross-Worksheet Hyperlink

Creating a hyperlink to navigate within the same worksheet involves simply mentioning the cell reference. For different worksheets, use the format "worksheetName!cellAddress". For instance, linking to cell A1 in "Sheet2" would be formatted as "Sheet2!A1".

Defined name scopes can be limited to either a specific worksheet or the entire workbook. To link to a named range within the same worksheet or the entire workbook, directly use the defined name. To link a named range in a different worksheet, specify it as seen in the example below.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet1 = workBook.CreateWorkSheet("Sheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("Sheet2");

// Define a name with workbook scope
workSheet1["D5"].SaveAsNamedRange("GlobalName", true);

// Define a name with worksheet scope
workSheet2["D10"].SaveAsNamedRange("LocalName", false);

// Creating hyperlinks within the same worksheet
workSheet1["A1"].Value = "Z20";
workSheet1["A1"].First().Hyperlink = "Z20";
workSheet1["A2"].Value = "GlobalName";
workSheet1["A2"].First().Hyperlink = "GlobalName";

// Creating hyperlinks across worksheets
workSheet1["A3"].Value = "A1 on Sheet2";
workSheet1["A3"].First().Hyperlink = "Sheet2!A1";
workSheet1["A4"].Value = "LocalName on Sheet2";
workSheet1["A4"].First().Hyperlink = "Sheet2!LocalName";

workBook.SaveAs("setCrossWorksheetHyperlink.xlsx");
```

### Demonstration

![Hyperlink Across Worksheet](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-hyperlink-across-worksheet.gif)

## Example: Creating Various Other Types of Hyperlinks

IronXL supports hyperlinks for different protocols including FTP, file paths, and email addresses:

- **FTP**: Use the prefix `ftp://`
- **File**: Begin with the protocol `file:///`
- **Email**: Start with `mailto:`

Here’s how you can implement these various hyperlink types:

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Link to open a file via FTP
workSheet["A1"].Value = "Link to a file via FTP";
workSheet["A1"].First().Hyperlink = "ftp://C:/Users/sample.xlsx";

// Link to open a local file
workSheet["A2"].Value = "Open local file";
workSheet["A2"].First().Hyperlink = "file:///C:/Users/sample.xlsx";

// Send an email
workSheet["A3"].Value = "Send email to example@gmail.com";
workSheet["A3"].First().Hyperlink = "mailto:example@gmail.com";

workBook.SaveAs("setMiscHyperlinks.xlsx");
```

### Demonstration

![Other Types of Hyperlinks](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-other-hyperlink.gif)