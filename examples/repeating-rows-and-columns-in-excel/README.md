***Based on <https://ironsoftware.com/examples/repeating-rows-and-columns-in-excel/>***

When handling multi-page Excel spreadsheets, it becomes significantly easier and faster to comprehend the data when column or row titles are printed on the top or right of each page. These titles are referred to as _Repeating Rows and Columns_ or _Header Rows and Columns_. IronXL simplifies the implementation of this effective feature with a minimalist amount of code.

## `SetRepeatingRows(startRow, endRow)`

Utilize this function to designate rows that should repeat across pages. For instance, `workSheet.SetRepeatingRows(3, 4)` configures repetition for `row(4-5)`.

## `SetRepeatingColumns(startColumn, endColumn)`

Similarly, this function marks columns for repetition. For example, calling `workSheet.SetRepeatingColumns(0, 2)` will repeat the contents of `column(A-C)`.

Both functions employ zero-based indexing for parameter values, where column(0) corresponds to "A" and row(1) to 2. It's important to note the behavior when combining these methods as shown above.

Content spanning multiple pages and aligned along the right of the first page will only adhere to the repeating column rules. Refer to Figure 1 for an illustration:

## Figure 1

![Figure 1](https://ironsoftware.com/static-assets/excel/examples/repeating-rows-and-columns-in-excel/repeating-rows-and-columns-in-excel-2.webp)

For content that spans along the bottom side of the first page, the repeating row settings will be in effect.

Pages on the inner side will see both repeating columns and rows. Figure 2 provides a visual depiction of these scenarios:

## Figure 2

![Figure 2](https://ironsoftware.com/static-assets/excel/examples/repeating-rows-and-columns-in-excel/repeating-rows-and-columns-in-excel-3.webp)