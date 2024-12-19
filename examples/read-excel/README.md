***Based on <https://ironsoftware.com/examples/read-excel/>***

IronXL is a C# and .NET Excel library that enables developers to read Excel data from formats such as **XLSX, XLS, XLSM, XLTX, CSV, and TSV** _without relying on Microsoft.Office.Interop.Excel_. Although all formats can be accessed using the `Load` method, it is advisable to utilize the `LoadCSV` method specifically for CSV files.

## Selecting a Worksheet

A **WorkSheet** is essentially a single page or tab within a **WorkBook**. There are various ways to select a specific WorkSheet for reading and editing:
 
- By the worksheet's index within the workbook's collection: `workBook.WorkSheets[0]`
- By specifying the worksheet's name in the `GetWorkSheet` method: `workBook.GetWorkSheet("workSheet")`
- Through the **DefaultWorkSheet** property of the workbook: `workBook.DefaultWorkSheet`
  
  Note: The default property selects the first worksheet, and if none exist, it creates a new worksheet named "Sheet1".
 
Additionally, you can interact with specific **Ranges**, **Rows**, and **Columns** within a **WorkSheet** to modify cell data or apply formulas.

Visit [Select Excel Range](https://ironsoftware.com/csharp/excel/examples/select-excel-range/) to learn more about access techniques for **Ranges**, **Rows**, and **Columns**.