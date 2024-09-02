The **IronXL** library enables the creation of Excel documents using XLS and XLSX formats. Leverage IronXL’s straightforward APIs to fill your workbook and utilize the `SaveAs` method to store it in various formats such as **XLS, XLSX, XLSM, CSV, TSV, JSON, XML, or HTML**. Additionally, IronXL supports exporting data to formats like **HTML string, Binary, Byte array, Data set, and Memory stream**.

While the `SaveAs` method allows for exporting to CSV, JSON, XML, and HTML, it is advisable to use dedicated methods for each file type for enhanced efficiency:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

It's important to note that when exporting to CSV, TSV, JSON, and XML, a separate file will be produced for each worksheet. The files are named according to the pattern `fileName.sheetName.format`. For example, the output file name for the CSV format would be `sample.new_sheet.csv`.

Explore how to switch between different spreadsheet formats using this [Code Example](https://ironsoftware.com/csharp/excel/examples/convert-excel-spreadsheet/).