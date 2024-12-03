***Based on <https://ironsoftware.com/examples/freeze-panes-in-excel/>***

The example provided illustrates the technique of creating a freeze pane in spreadsheets, which secures rows and columns so they remain visible when scrolling through the document. This is an essential feature for maintaining visibility of headers while you navigate and compare large sets of data efficiently.

## `CreateFreezePane(column, row)`

The basic usage of the `CreateFreezePane` method involves specifying the column and row to establish the freeze pane based on these. For instance, invoking `workSheet.CreateFreezePane(1, 4)` effectively creates a freeze pane starting from **column (A)** to **row (1-4)**.

## `CreateFreezePane(column, row, subsequentColumn, subsequentRow)`

A more advanced overload of `CreateFreezePane` allows not only fixing columns and rows but also setting parameters for scrolling in the worksheet. By calling `workSheet.CreateFreezePane(5, 2, 6, 7)`, the freeze pane includes **columns (A-E)** and **rows (1-2)** and allows scrolling from column **6** and row **7**. Initially, when the worksheet is opened, it displays columns A-E, G-... and rows 1-2, 8-...

Freezing panes is particularly beneficial when dealing with extensive data tables in Excel, enabling you to keep critical rows or columns in view as you scroll through the rest of your data.

For further guidance and examples, please visit the ["Freeze Panes" How-To](https://ironsoftware.com/csharp/excel/how-to/add-freeze-panes/) article.