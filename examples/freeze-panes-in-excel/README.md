The code example above illustrates how to implement freeze panes, which lock specified rows and columns in place to keep them visible as you scroll. This functionality is particularly valuable for maintaining the visibility of headers when you need to rapidly compare data.

## `CreateFreezePane(column, row)`

The initial version of the `CreateFreezePane` function takes the number of columns and rows as parameters to set up the freeze pane. For instance, calling `workSheet.CreateFreezePane(1, 4)` establishes a freeze pane beginning from **column(A)** to **row(1-4)**.

## `CreateFreezePane(column, row, subsequentColumn, subsequentRow)`

This variant of the `CreateFreezePane` method not only sets the freeze panes based on the specified number of columns and rows, but it also enables scrolling within the worksheet. For example, a call to `workSheet.CreateFreezePane(5, 2, 6, 7)` results in a freeze pane from **column(A-E)** and **row(1-2)** with scrolling over **1 column** and **5 rows**. Upon opening, the worksheet initially displays columns A-E, G-... and rows 1-2, 8-...

Freezing rows or columns can be tremendously beneficial for navigating through large datasets in Excel, as it allows static viewing of selected rows or columns while you scroll through other parts of your worksheet.



For further details and examples, please visit the ["Freeze Panes" How-To](https://ironsoftware.com/csharp/excel/how-to/add-freeze-panes/) article.