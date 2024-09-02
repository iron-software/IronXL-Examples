IronXL empowers users to efficiently select and manipulate ranges within any Excel `WorkSheet`. The demonstrated code snippets show how to easily access and work with ranges, rows, and columns. With IronXL, you can perform additional operations on these data sets, such as `SortAscending()`, `SortDescending()`, `Sum()`, `Max()`, `Min()`, and `Avg()`. It’s important to note that methods which alter or relocate cell values also automatically adjust the values of the associated range, row, and column.

## Range

To select a range from **A2** to **A8**, use the following code: `var range = sheet["A2:A8"]`.

## Row

For selecting row **1**, apply the `GetRow(0)` method. This utilizes zero-based indexing, and the selected range includes all cells populated in row 1 as determined by their collective union.

## Column

To access column **A**, you can either use `GetColumn(0)` or directly specify the range with `sheet["A:A"]`. Like rows, the range for columns is also defined by a union of all populated cells in column A.