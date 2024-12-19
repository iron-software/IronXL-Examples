***Based on <https://ironsoftware.com/examples/excel-print-setup/>***

IronXL provides detailed control over the **Print Setup** for any Excel file, allowing precise management of various print settings whether the output is to a physical or PDF printer.

In addition, document headers and footers can be customized. This includes support for "mail merge" style variables that dynamically insert data:

- `&P`: page numbers
- `&N`: total page count
- `&D`: the current date
- `&T`: the current time
- `&Z&F`: full file path
- `&F`: file name only
- `&A`: active sheet name

These variables can be straightforwardly inserted into the `Footer` property. For example, to have the footer display the phrase "Page &P of &N" on each printed page, you would use:
```csharp
workSheet.Footer.Center = "Page &P of &N";
```

This functionality allows IronXL to offer comprehensive control over how Excel spreadsheets are handled and printed.