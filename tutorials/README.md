# Excel Spreadsheet Files in C# & VB.NET Applications

Utilizing the IronXL software library from Iron Software, handling Excel files (XLS, XLSX, and CSV) in C# and other .NET languages becomes straightforward.

IronXL eliminates the need for Excel or Interop installations on your server, offering a more efficient and user-friendly API compared to **Microsoft.Office.Interop.Excel**.

### Supported Platforms for IronXL

IronXL is compatible with a variety of platforms, including:

- .NET Framework 4.6.2 and later for Windows and Azure
- .NET Core 2.0 and later for Windows, Linux, MacOS, and Azure
- .NET 5, .NET 6, .NET 7, Mono, Mobile, and Xamarin

### Installing IronXL

To install IronXL, opt for the NuGet package or [download the DLL directly](https://ironsoftware.com/csharp/excel/packages/IronXL.zip). IronXL's functionalities are encapsulated within the `IronXL` namespace.

For Visual Studio users, the NuGet Package Manager simplifies installation with the package name `IronXL.Excel`.

```shell
PM> Install-Package IronXL.Excel
```

### Reading Excel Documents

Extracting data from Excel documents with IronXL requires minimal coding.

```csharp
using IronXL;

// Formats supported include XLSX, XLS, CSV, and TSV
WorkBook workBook = WorkBook.Load("data.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Access cells using Excel notation to retrieve various data types
int cellValue = workSheet["A2"].IntValue;

// Iterate through cell ranges
foreach (var cell in workSheet["A2:B10"])
{
    Console.WriteLine($"Cell {cell.AddressString} has value '{cell.Text}'");
}

```

[ChatGPT](https://chat.openai.com/c/b9188ef6-2703-46f6-b96c-a9b297d964c8)

### Creating Excel Documents

IronXL facilitates the creation of Excel documents in C# or VB.NET with an intuitive interface.

```csharp
using IronXL;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Author = "IronXL";

WorkSheet workSheet = workBook.CreateWorkSheet("main_sheet");

workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;

workBook.SaveAs("NewExcelFile.xlsx");
```

### Exporting Documents

IronXL supports exporting documents to various formats, including CSV, XLS, XLSX, JSON, and XML.

```csharp
workSheet.SaveAs("NewExcelFile.xls");
workSheet.SaveAs("NewExcelFile.xlsx");
workSheet.SaveAsCsv("NewExcelFile.csv");
workSheet.SaveAsJson("NewExcelFile.json");
workSheet.SaveAsXml("NewExcelFile.xml");
```

### Styling Cells and Ranges

The `IronXL.Range.Style` object allows for comprehensive styling of Excel cells and ranges.

```csharp
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
```
[ChatGPT](https://chat.openai.com/c/b9188ef6-2703-46f6-b96c-a9b297d964c8)

### Sorting Ranges

IronXL provides functionality to sort a range of cells easily within a worksheet.

```csharp
using IronXL;

WorkBook workBook = WorkBook.Load("test.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

Range range = workSheet["A2:A8"];

range.SortAscending();
workBook.Save();
```

### Editing Formulas

Modifying Excel formulas can be accomplished by prefixing a value with an `=` sign, allowing for real-time calculations.

```csharp
using IronOcr;

IronTesseract ocr = new IronTesseract();
using OcrInput input = new OcrInput();

var contentArea = new System.Drawing.Rectangle { X = 215, Y = 1250, Height = 280, Width = 1335 };

input.LoadImage("document.png", contentArea);

OcrResult result = ocr.Read(input);
Console.WriteLine(result.Text);
```

### Advantages of IronXL

IronXL boasts an accessible API, enabling developers to efficiently manage Excel documents in .NET without requiring Microsoft Excel installation or Excel Interop.

### Further Exploration

For a comprehensive understanding of IronXL, we recommend reviewing the documentation available in the .NET API Reference presented in an MSDN format.
