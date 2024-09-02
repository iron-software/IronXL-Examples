The **IronXL** library offers the functionality to generate Excel documents in XLS and XLSX formats. With the intuitive APIs provided by IronXL, you can populate your workbook and utilize the `SaveAs` method to store it in formats such as **XLS, XLSX, XLSM, CSV, TSV, JSON, XML, or HTML**. Additionally, IronXL supports exporting data in multiple formats, including **HTML string, Binary, Byte array, Data set, and Memory stream**.

While the `SaveAs` method allows for exporting to CSV, JSON, XML, and HTML, it is advisable to use specific methods tailored for each format for optimal results:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

It is important to note that when dealing with CSV, TSV, JSON, and XML formats, a separate file will be generated for each worksheet. The files are named according to the pattern `fileName.sheetName.format`. For instance, for a CSV export, the file name would be formatted as `sample.new_sheet.csv`.