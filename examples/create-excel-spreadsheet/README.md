***Based on <https://ironsoftware.com/examples/create-excel-spreadsheet/>***

The IronXL library provides the functionality to develop Excel files using both XLS and XLSX formats. This is facilitated through the user-friendly APIs offered by IronXL, which allow you to modify and populate your workbook effectively. You can manage the value of a cell using the `Value` property and modify a cell's style with IronXL's capabilities.

The `Style` properties that can be customized include:

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

For file formats such as CSV, TSV, JSON, and XML, IronXL will generate a separate file for each sheet within the workbook. The format for the file names is `fileName.sheetName.format`. As an example, for a CSV file, the resultant file name would be `sample.new_sheet.csv`.