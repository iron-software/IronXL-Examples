# How to Create Hyperlinks in Excel

***Based on <https://ironsoftware.com/how-to/hyperlinks/>***


Excel hyperlinks allow for clickable links to various destinations like other workbook locations, different files, web pages, or email addresses. This functionality improves usability by offering instant access to relevant data and external sites, making spreadsheets more interactive and friendly for users.

With IronXL, you can easily embed hyperlinks for URLs, external files from local or FTP systems, email addresses, cell references, and named cells directly in your .NET C# projects without requiring Interop.

## Example: Creating a URL Hyperlink

The `Hyperlink` attribute can be found within the `Cell` class. To obtain a `Range` object, you could use `workSheet["A1"]` and access the first cell using the `First` method.

Alternatively, the `GetCellAt` method lets you target a cell directly if you wish to manage its `Hyperlink` attribute.

Here, we delve into an example that illustrates how to set up hyperlinks using both the HTTP and HTTPS protocols.

It's worth noting that utilizing the `GetCellAt` method on a cell that has not been initialized could potentially result in a *System.NullReferenceException: 'Object reference not set to an instance of an object.'*

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.Hyperlinks
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.New(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.FirstSheet;
            
            // Assign a value to the cell
            workSheet["A1"].Value = "Visit IronPDF";
            
            // Apply a hyperlink to https://ironpdf.com/ at cell A1
            workSheet.GetCellAt(0, 0).Hyperlink = "https://ironpdf.com/";
            
            workBook.ExportAs("createURLHyperlink.xlsx");
        }
    }
}
```

### Visualization

![Link Hyperlink](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-link-hyperlink.gif)

## Example: Creating Hyperlinks Across Different Worksheets

To link to a cell within the same worksheet, you simply use the cell's address like Z20. For inter-worksheet links, employ the format "worksheetName!address". For instance, to link to cell A1 on "Sheet2", you'd write "Sheet2!A1".

Named cells can be scoped either globally (workbook-wide) or locally (worksheet-only). To link a globally scoped name within the same worksheet or a locally scoped name on a different sheet, you'd specify the worksheet name first followed by the name.

```cs
using System.Linq;
using IronXL.Excel;
namespace ironxl.Hyperlinks
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.New(ExcelFileFormat.XLSX);
            WorkSheet sheetOne = workBook.AddNewSheet("Sheet1");
            WorkSheet sheetTwo = workBook.AddNewSheet("Sheet2");
            
            // Define a global named cell
            sheetOne["D5"].SaveAsNamedRange("Iron", workbookLevel: true);
            
            // Define a local named cell
            sheetTwo["D10"].SaveAsNamedRange("Hello", workbookLevel: false);
            
            // Same worksheet hyperlink to cell Z20
            sheetOne["A1"].Value = "Go to Z20";
            sheetOne["A1"].First().Hyperlink = "Z20";
            
            // Same worksheet hyperlink to named range "Iron"
            sheetOne["A2"].Value = "Go to Iron";
            sheetOne["A2"].First().Hyperlink = "Iron";
            
            // Different worksheet hyperlinks
            sheetOne["A3"].Value = "Go to A1 on Sheet2";
            sheetOne["A3"].First().Hyperlink = "Sheet2!A1";
            
            sheetOne["A4"].Value = "Go to Hello on Sheet2";
            sheetOne["A4"].First().Hyperlink = "Sheet2!Hello";
            
            workBook.ExportAs("createWorksheetHyperlink.xlsx");
        }
    }
}
```

### Visualization

![Hyperlink Across Worksheet](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-hyperlink-across-worksheet.gif)

## Example: Creating FTP, File, and Email Hyperlinks

In addition to standard URL hyperlinks, IronXL supports creating hyperlinks for FTP paths, file locations, and email addresses.

Note that FTP and File hyperlinks should use absolute paths.

```cs
using System.Linq;
using IronXL.Excel;
namespace ironxl.Hyperlinks
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.New(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.FirstSheet;
            
            // Creating an FTP hyperlink
            workSheet["A1"].Value = "Access FTP sample.xlsx";
            workSheet["A1"].First().Hyperlink = "ftp://C:/Users/sample.xlsx";
            
            // Creating a file hyperlink
            workSheet["A2"].Value = "Open file sample.xlsx";
            workSheet["A2"].First().Hyperlink = "file:///C:/Users/sample.xlsx";
            
            // Creating an email hyperlink
            workSheet["A3"].Value = "Send email to example@gmail.com";
            workSheet["A3"].First().Hyperlink = "mailto:example@gmail.com";
            
            workBook.ExportAs("createMiscHyperlinks.xlsx");
        }
    }
}
```

### Visualization

![Other Types of Hyperlinks](https://ironsoftware.com/static-assets/excel/how-to/hyperlinks/hyperlinks-set-other-hyperlink.gif)