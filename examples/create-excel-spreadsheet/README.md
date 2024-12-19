***Based on <https://ironsoftware.com/examples/create-excel-spreadsheet/>***

The IronXL library is adept at generating Excel documents from both XLS and XLSX formats. Employ IronXL's user-friendly APIs to modify and populate your Excel workbook efficiently. You can access and set the value of a cell through the `Value` property, while you can also modify the style of the cells.

The following list enumerates the style properties that you can adjust using IronXL:

- `DiagonalBorder`
- `Indention`
- `Rotation`
- `FillPattern`
- `VerticalAlignment`
- `HorizontalAlignment`
- `DiagonalBorderDirection`
- `WrapText`
- `ShrinkToFit`
- `TopBorder`
- `RightBorder`
- `LeftBorder`
- `BackgroundColorFont`
- `BottomBorder`
- `SetBackgroundColor`

For formats like CSV, TSV, JSON, and XML, IronXL will create a separate file for each sheet in the workbook. The files are named using the format `fileName.sheetName.format`. For instance, in the case of exporting to CSV, the filename would be `sample.new_sheet.csv`.