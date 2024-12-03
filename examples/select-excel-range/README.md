***Based on <https://ironsoftware.com/examples/select-excel-range/>***

IronXL facilitates the easy selection and manipulation of ranges within any Excel `WorkSheet`. The provided code illustrates how ranges, rows, and columns are effortlessly selected and manipulated. With IronXL, performing operations like `SortAscending()`, `SortDescending()`, `Sum()`, `Max()`, `Min()`, and `Avg()` on these data collections is straightforward. It's important to be aware that methods which alter or relocate cell values will accordingly update the values in the affected ranges, rows, and columns.

## Range

To select a range from **A2** to **A8**, you can use: `var range = sheet["A2:A8"]`.

## Row

For selecting the first row, employ the method `GetRow(0)`. This method adheres to zero-based indexing, and the covered range consists of cells from all populated cells in the row, including those in row 1 itself.

## Column

To access column **A**, you can either use `GetColumn(0)` or directly set the range with `sheet["A:A"]`. The selected range will include all populated cells across the column, up to and including those in column A.