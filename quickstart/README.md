# Excel Spreadsheet Handling in C# and VB.NET Applications

***Based on <https://ironsoftware.com/docs/docs/>***


Effortlessly manipulate Excel (XLS, XLSX, and CSV) files in C# and various other .NET languages using the robust IronXL software library from Iron Software.

IronXL is designed to function seamlessly without the necessity of having Excel installed on your machine, eliminating the need for Microsoft Office Interop. Its API is more agile and user-friendly compared to **Microsoft.Office.Interop.Excel**.

IronXL is compatible with the following environments:

* .NET Framework 4.6.2 or newer, supporting Windows and Azure deployments.
* .NET Core 2 or higher, accommodating Windows, Linux, MacOS, and Azure.
* .NET 5, .NET 6, .NET 7, and .NET 8 as well as Mono, Mobile, and Xamarin platforms.

## Setting up IronXL

To begin with IronXL, you can either utilize our NuGet package or by [downloading the DLL directly from our site](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

The easiest method to install IronXL is by using the NuGet Package Manager integrated in Visual Studio:
The required package to search for is **IronXL.Excel**.

```shell
Install-Package IronXL.Excel
```

[Visit our NuGet page for more details](https://www.nuget.org/packages/ironxl.excel/)

## How to Read an Excel File

Utilizing IronXL to extract data from an Excel document is straightforward.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section1
    {
        public void Run()
        {
            // Reading spreadsheet formats such as XLSX, XLS, CSV, and TSV
            WorkBook workBook = WorkBook.Load("data.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();

            // Access cells using Excel notation and retrieve values, dates, texts, or formulas.
            int cellValue = workSheet["A2"].IntValue;

            // Elegant handling of cell ranges.
            foreach (var cell in workSheet["A2:B10"])
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
            }
        }
    }
}
```

## Creating New Excel Documents

IronXL offers a straightforward approach to create new Excel files in C# or VB.NET.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section2
    {
        public void Run()
        {
            // Initialize a new Excel WorkBook.
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.Metadata.Author = "IronXL";

            // Create a new sheet.
            WorkSheet workSheet = workBook.CreateWorkSheet("main_sheet");

            // Insert data and style elements into cells
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;

            // Save the document to a file
            workBook.SaveAs("NewExcelFile.xlsx");
        }
    }
}
```

## File Export Options

IronXL supports exporting to numerous popular file formats.

```cs
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section3
    {
        public void Run()
        {
            // Fluent file saving across multiple formats
            workSheet.SaveAs("NewExcelFile.xls");
            workSheet.SaveAs("NewExcelFile.xlsx");
            workSheet.SaveAsCsv("NewExcelFile.csv");
            workSheet.SaveAsJson("NewExcelFile.json");
            workSheet.SaveAsXml("NewExcelFile.xml");
        }
    }
}
```

## Styling Cells and Ranges

Customize the style of Excel cells and ranges using IronXL.

```cs
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section4
    {
        public void Run()
        {
            // Customize cell values and styles
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
        }
    }
}
```

## Sorting Cell Ranges

IronXL allows for the easy sorting of Excel cell ranges.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section5
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();

            // Retrieve and sort a range in the worksheet
            Range range = workSheet["A2:A8"];

            range.SortAscending();
            workBook.Save();
        }
    }
}
```

## Editing Cell Formulas

Adjusting Excel formulas is straightforward with IronXL, with real-time calculation.

```cs
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section6
    {
        public void Run()
        {
            // Assign and calculate a formula
            workSheet["A1"].Value = "=SUM(A2:A10)";
            
            decimal sum = workSheet["A1"].DecimalValue;
        }
    }
}
```

## Why Favor IronXL?

IronXL provides an accessible and efficient API for .NET developers to handle Excel documents, bypassing the need for direct Excel software installations or Microsoft Excel Interop.

## Advancing with IronXL

To discover more about IronXL, delve into the [.NET API Reference](https://ironsoftware.com/csharp/excel/object-reference/) styled like MSDN for comprehensive insights and deeper understanding.