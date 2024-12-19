***Based on <https://ironsoftware.com/examples/select-excel-range/>***

IronXL empowers users to effortlessly access and manipulate ranges within any Excel `WorkSheet`. The examples provided showcase how to select ranges, rows, and columns seamlessly. With IronXL, you can enhance this data set by implementing methods like `SortAscending()`, `SortDescending()`, `Sum()`, `Max()`, `Min()`, and `Avg()`. It's important to remember that methods modifying or moving cell values will impact the corresponding range, row, and column values as well.

## Range

To select a specific range from **A2 to A8**, you can use the following code snippet: 
```csharp
var range = sheet["A2:A8"];
```

## Row

For selecting row **1**, the method `GetRow(0)` is utilized, adhering to a zero-based indexing system. The cell range for this row is defined by the combined area of all filled cells within row 1.

## Column

To access column **A**, you can either employ `GetColumn(0)` or specify the column directly using:
```csharp
var column = sheet["A:A"];
```
The cell range for the column is defined by the populated area covering all of column A.