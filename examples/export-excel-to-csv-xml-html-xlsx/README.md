***Based on <https://ironsoftware.com/examples/export-excel-to-csv-xml-html-xlsx/>***

The **IronXL** library provides a robust solution for creating Excel documents, supporting both XLS and XLSX formats. Utilizing IronXL's straightforward APIs, you can conveniently populate your workbook and then save it using the `SaveAs` method in various formats including **XLS, XLSX, XLSM, CSV, TSV, JSON, XML, or HTML**. Additionally, IronXL enables exporting data in formats such as **HTML string, Binary, Byte array, DataSet, and MemoryStream**.

While the `SaveAs` method can handle exports in CSV, JSON, XML, and HTML, it is advisable to use specific methods designed for each file type for better efficiency:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

It is important to note that when dealing with CSV, TSV, JSON, and XML formats, each sheet will be exported as a separate file named following the pattern `fileName.sheetName.format`. For example, the CSV export would be named `sample.new_sheet.csv`.

Enhance your understanding of converting between different spreadsheet formats by checking out this [Code Example](https://ironsoftware.com/csharp/excel/examples/convert-excel-spreadsheet/).