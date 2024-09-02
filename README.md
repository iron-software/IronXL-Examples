![Nuget Version](https://img.shields.io/nuget/v/IronXL.Excel?color=informational&label=latest) ![Installs Count](https://img.shields.io/nuget/dt/IronXL.Excel?color=informational&label=installs&logo=nuget) ![Build Status](https://img.shields.io/badge/build-%20%E2%9C%93%202425%20tests%20passed%20(0%20failed)%20-107C10?logo=visualstudio) [![Windows Support](https://img.shields.io/badge/%E2%80%8E%20-%20%E2%9C%93-107C10?logo=windows)](https://ironsoftware.com/csharp/excel/docs/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield) [![macOS Support](https://img.shields.io/badge/%E2%80%8E%20-%20%E2%9C%93-107C10?logo=apple)](https://ironsoftware.com/csharp/excel/docs/questions/macos?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield) [![Linux Support](https://img.shields.io/badge/%E2%80%8E%20-%20%E2%9C%93-107C10?logo=linux&logoColor=white)](https://ironsoftware.com/csharp/excel/docs/questions/linux?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield) [![Docker Support](https://img.shields.io/badge/%E2%80%8E%20-%20%E2%9C%93-107C10?logo=docker&logoColor=white)](https://ironsoftware.com/csharp/excel/docs/questions/docker-support?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield) [![AWS Support](https://img.shields.io/badge/%E2%80%8E%20-%20%E2%9C%93-107C10?logo=amazonaws)](https://ironsoftware.com/csharp/excel/docs/questions/aws-lambada-support?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield) [![Microsoft Azure Support](https://img.shields.io/badge/%E2%80%8E%20-%20%E2%9C%93-107C10?logo=microsoftazure)](https://ironsoftware.com/csharp/excel/docs/questions/azure-support?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield) [![Live Chat](https://img.shields.io/badge/Live%20Chat-Active-purple?logo=googlechat&logoColor=white)](https://ironsoftware.com/csharp/excel/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topshield#helpscout-support)

## IronXL - The C# Excel Library

[![IronXL NuGet Trial Banner Image](https://raw.githubusercontent.com/iron-software/iron-nuget-assets/main/IronXL-readme/nuget-trial-banner-large.png)](https://ironsoftware.com/csharp/excel/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=topbanner#trial-license)

[Get Started](https://ironsoftware.com/csharp/excel/docs/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=navigation) | [Features](https://ironsoftware.com/csharp/excel/features/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=navigation) | [Code Examples](https://ironsoftware.com/csharp/excel/examples/read-excel/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=navigation) | [Licensing](https://ironsoftware.com/csharp/excel/licensing/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=navigation) | [Free Trial](https://ironsoftware.com/csharp/excel/docs/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=navigation#trial-license)

IronXL is a robust C# library developed by Iron Software for managing and interacting with Excel and other spreadsheet files within .NET projects. It enables developers to read, create, and manipulate spreadsheet documents directly from their applications without needing to install Microsoft Office, and it's fully compatible with .NET, .NET Core, and Azure platforms.

#### Core Capabilities of IronXL:

  * Data import from formats including XLS, XLSX, CSV, and TSV.
  * Export data to formats like XLS, XLSX, CSV, TSV, and JSON.
  * Excel file encryption and decryption.
  * Operate on Excel sheets as `System.Data.DataSet` and `System.Data.DataTable`.
  * Automatic recalculation of formulas upon modifications.
  * Simple cell reference syntax like `WorkSheet["A1:B10"]`.
  * Sorting capabilities for ranges, columns, and rows.
  * Extensive cell formatting options such as font, background, border styles, and alignment.

##### Document Management

  * Supported Formats for Loading and Editing: XLS, XLSX, XLSM, XLST, CSV, TSV
  * Formats for Saving and Exporting: XLS, XLSX, XLSM, XLST, CSV, TSV, JSON
  * System.Data Integration: Manage spreadsheets as `System.Data.DataSet` and `System.Data.DataTable`

##### Spreadsheet Operations:

  * Formula Management: Seamless integration and automatic updates of Excel formulas.
  * Data Formatting: Extensive options including text, number, date, currency, and custom formats.
  * Advanced Sorting and Cell Styling: Customize cell appearance and spreadsheet structure.

#### Cross-Platform Compatibility:

  * Supports **.NET 8** and earlier versions, .NET Core, Standard, Framework
  * Compatible with Windows, macOS, Linux, Docker, Azure, and AWS environments

[![Cross-Platform Compatibility](https://raw.githubusercontent.com/iron-software/iron-nuget-assets/main/IronXL-readme/cross-platform-compatibility.png)](https://ironsoftware.com/csharp/excel/docs/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=crossplatformbanner)

Discover more about our API and complete licensing details on our website.

#### Getting Started with IronXL:

To integrate IronXL into your project, simply install the NuGet package:

```plaintext
PM> Install-Package IronXL.Excel
```

To jumpstart your project, add `using IronXL` to your C# files. Here’s an initiation example:

```csharp
using IronXL;
using System.Linq;

// Read various spreadsheet formats:
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet sheet = workbook.WorkSheets.First();

// Fetch a specific cell using Excel-like syntax:
int cellValue = sheet["A2"].IntValue;

// Elegantly access and read range of cells:
foreach (var cell in sheet["A2:A10"]) {
    Console.WriteLine($"Cell {cell.AddressString} has value '{cell.Text}'");
}

// Use LINQ to compute aggregates:
decimal sum = sheet["A2:A10"].Sum();
decimal max = sheet["A2:A10"].Max(c => c.DecimalValue);
```

### Features and Licensing

[![IronXL Features](https://raw.githubusercontent.com/iron-software/iron-nuget-assets/main/IronXL-readme/features-table.png)](https://ironsoftware.com/csharp/excel/features/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=featuresbanner)

For more code samples, tutorials and detailed documentation, visit [https://ironsoftware.com/csharp/excel/](https://ironsoftware.com/csharp/excel/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=supportanddocs)

For detailed support, contact us via support@ironsoftware.com.

### Helpful Resources

  * Code Samples: [https://ironsoftware.com/csharp/excel/examples/](https://ironsoftware.com/csharp/excel/examples/read-excel/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=supportanddocs)
  * API Documentation: [https://ironsoftware.com/csharp/excel/object-reference/api/](https://ironsoftware.com/csharp/excel/object-reference/api/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=supportanddocs)
  * Tutorials: [https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=supportanddocs)
  * Licensing Information: [https://ironsoftware.com/csharp/excel/licensing/](https://ironsoftware.com/csharp/excel/licensing/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=supportanddocs)
  * Live Chat Assistance: [https://ironsoftware.com/csharp/excel/#helpscout-support](https://ironsoftware.com/csharp/excel/?utm_source=nuget&utm_medium=organic&utm_campaign=readme&utm_content=supportanddocs#helpscout-support)

For questions or direct assistance, please email support@ironsoftware.com. We provide comprehensive support and licensing options for your commercial projects.