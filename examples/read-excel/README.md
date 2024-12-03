***Based on <https://ironsoftware.com/examples/read-excel/>***

IronXL is a robust Excel Library tailored for C# and .NET, enabling developers to read Excel content from formats such as **XLSX, XLS, XLSM, XLTX, CSV, and TSV** _without relying on Microsoft.Office.Interop.Excel_. While the `Load` method supports all these formats, it is advised to utilize the `LoadCSV` method specifically for CSV files.

## Worksheet Selection

A **WorkSheet** denotes a single sheet or tab within a **WorkBook**. There are several methods to select a WorkSheet for reading and editing purposes:

- Select by the worksheet's index in the collection: `workBook.WorkSheets[0]`
- Select by the worksheet's name using `GetWorkSheet`: `workBook.GetWorkSheet("workSheet")`
- Use the **DefaultWorkSheet** property to access the first worksheet: `workBook.DefaultWorkSheet`
- If no worksheets exist, it will generate and return a new one named "Sheet1."

In addition, you can access and modify individual **Ranges**, **Rows**, and **Columns** within a **WorkSheet** to manage cell data or implement formulas.

Discover more about how to select **Range**, **Row**, and **Column** by visiting [Select Excel Range](https://ironsoftware.com/csharp/excel/examples/select-excel-range/).