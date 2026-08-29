# IronXL.Examples

Runnable C# examples for [IronXL](https://ironsoftware.com/csharp/excel/?utm_source=github), a .NET Excel library that reads, writes, and edits XLSX, XLS, CSV, and TSV files without Microsoft Office or COM Interop.

## Install

```bash
dotnet add package IronXL.Excel
```

## Quickstart

```csharp
using IronXL;

WorkBook workbook = WorkBook.Load("data.xlsx");
WorkSheet sheet = workbook.WorkSheets.First();

// Read cells using Excel-like addressing
string customerName = sheet["A2"].StringValue;
decimal total = sheet["B2:B100"].Sum();

// Write a value and save
sheet["C1"].Value = "Updated";
workbook.SaveAs("updated.xlsx");
```

Ranges support the LINQ-style aggregations you'd expect (`Sum`, `Max`, `Min`, `Average`) and iterate cell-by-cell. Worksheets can also be projected into `System.Data.DataTable` and `System.Data.DataSet` for integration with existing .NET data pipelines.

For production use, set a license key via `License.LicenseKey = "YOUR-KEY"`. Without one, saved files include a watermark sheet.

## What's in this repo

Each folder contains a self-contained .NET project you can open and run:

- `examples/` — focused snippets demonstrating individual features
- `how-to/` — task-oriented guides for specific spreadsheet operations
- `quickstart/` — end-to-end project scaffolds
- `tutorials/` — longer walkthroughs combining multiple features

## Common tasks covered

- Reading and writing XLSX, XLS, XLSM, CSV, and TSV files
- Cell formulas with automatic recalculation on edit
- Range arithmetic: Sum, Max, Min, Average, and LINQ-style aggregations
- Cell formatting: fonts, colors, borders, alignment, number formats
- Sorting ranges, columns, and rows
- DataSet and DataTable conversion for ADO.NET integration
- Excel encryption, decryption, and password protection
- Export to JSON
- Conditional formatting and named ranges

## Platform support

.NET 8, 7, 6, 5, .NET Core, .NET Standard, and .NET Framework. Windows, macOS, Linux, Docker, Azure, and AWS. See the [installation docs](https://ironsoftware.com/csharp/excel/docs/?utm_source=github) for environment-specific notes.

## Documentation and support

- Full documentation: [ironsoftware.com/csharp/excel/docs](https://ironsoftware.com/csharp/excel/docs/?utm_source=github)
- API reference: [ironsoftware.com/csharp/excel/object-reference/api](https://ironsoftware.com/csharp/excel/object-reference/api/?utm_source=github)
- Issues with these examples: file directly on this repository
- Product support: [support@ironsoftware.com](mailto:support@ironsoftware.com)

## About

This repository is maintained by [Iron Software](https://ironsoftware.com/?utm_source=github). IronXL is a commercial library — see [licensing](https://ironsoftware.com/csharp/excel/licensing/?utm_source=github) for terms and trial details.
