IronXL is a library tailored for C# and .NET environments, designed to facilitate the loading and reading of Excel data from various file formats such as **XLSX, XLS, XLSM, XLTX, CSV, and TSV** without the need for _Microsoft.Office.Interop.Excel_. While the `Load` method can read files in any of the supported formats, using the `LoadCSV` method is specifically advisable for handling CSV files.

## Worksheet Selection

In the context of IronXL, a **WorkSheet** is akin to a sheet within an **Excel Book**, and it can also be accessed or manipulated in various ways:

- Access a worksheet by its index within the collection: `workBook.WorkSheets[0]`
- Retrieve a worksheet by name using the `GetWorkSheet` method: `workBook.GetWorkSheet("workSheet")`
- Utilize the **DefaultWorkSheet** property of the workbook: `workBook.DefaultWorkSheet`

Note that using the **DefaultWorkSheet** property will automatically provide the first worksheet in the workbook or create a new one named "Sheet1" if none exist.

Additionally, you can target specific **Ranges**, **Rows**, and **Columns** within a **WorkSheet** for data access, modification, or formula application.

Explore further details on targeting specifics within a worksheet by visiting [Select Excel Range](https://ironsoftware.com/csharp/excel/examples/select-excel-range/) for more on selecting **Ranges**, **Rows**, and **Columns**.