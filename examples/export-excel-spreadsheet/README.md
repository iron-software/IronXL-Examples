***Based on <https://ironsoftware.com/examples/export-excel-spreadsheet/>***

The **IronXL** library enables the creation of Excel documents in both XLS and XLSX formats. You can easily populate your workbook using IronXL's user-friendly APIs and then utilize the `SaveAs` method to save your workbook in various formats including **XLS, XLSX, XLSM, CSV, TSV, JSON, XML, or HTML**. Furthermore, IronXL supports exporting to numerous in-code data types such as **HTML string, Binary, Byte array, Data set, and Memory stream**.

Although the `SaveAs` method can export to CSV, JSON, XML, and HTML, it is advisable to use specific methods designed for each of these formats for optimal results:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

It's important to recognize that for CSV, TSV, JSON, and XML formats, a separate file is generated for each sheet within the workbook, adhering to the naming pattern `fileName.sheetName.format`. For instance, in the scenario described, a CSV export would produce a file named `sample.new_sheet.csv`.