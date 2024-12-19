***Based on <https://ironsoftware.com/examples/export-excel-to-csv-xml-html-xlsx/>***

The **IronXL** library facilitates the creation of Excel documents in both XLS and XLSX formats. With IronXL's straightforward APIs, you can easily populate your workbook and utilize the `SaveAs` method to save documents in various formats like **XLS, XLSX, XLSM, CSV, TSV, JSON, XML, or HTML**. Additionally, IronXL supports exporting data into formats such as **HTML string, Binary, Byte array, Dataset, and Memory stream**.

While the `SaveAs` method can export to CSV, JSON, XML, and HTML formats, it is advisable to use specialized methods for each specific file type for optimal results:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

It's important to note that when dealing with CSV, TSV, JSON, and XML file types, each sheet will result in a separate file. The files are named following the pattern `fileName.sheetName.format`. For instance, a CSV file generated from this process would be named `sample.new_sheet.csv`.

Enhance your skills in converting between different spreadsheet formats by exploring this [Code Example](https://ironsoftware.com/csharp/excel/examples/convert-excel-spreadsheet/).