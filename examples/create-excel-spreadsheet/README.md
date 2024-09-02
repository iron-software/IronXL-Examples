The IronXL library enables the creation of Excel documents in both XLS and XLSX formats through its straightforward API, which allows you to modify and populate your workbook efficiently. You can access the content of a cell using the `Value` property, and you can also alter the cell's style using various settings in IronXL.

The styling options available with IronXL include the following properties:

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

It's important to note that for formats like CSV, TSV, JSON, and XML, each sheet in the workbook will be saved as a separate file. The naming scheme for these files follows the pattern `fileName.sheetName.format`. For example, a CSV file created would be named `sample.new_sheet.csv`.