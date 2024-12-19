***Based on <https://ironsoftware.com/examples/export-excel-spreadsheet/>***

The **IronXL** library facilitates the creation of Excel documents from both XLS and XLSX formats. With IronXL's user-friendly APIs, you can easily populate your workbook and utilize the `SaveAs` method to save it in formats like **XLS, XLSX, XLSM, CSV, TSV, JSON, XML, or HTML**. Additionally, IronXL supports exporting data in various code formats such as **HTML string, Binary, Byte array, Data set, and Memory stream**.

While the `SaveAs` method allows exporting to CSV, JSON, XML, and HTML, it is advisable to use format-specific methods for these tasks:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

It's important to note that when dealing with CSV, TSV, JSON, and XML formats, each sheet will generate a separate file. The files are named using the pattern `fileName.sheetName.format`. For instance, if using the CSV format, the filename would be `sample.new_sheet.csv`.