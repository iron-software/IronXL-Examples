***Based on <https://ironsoftware.com/examples/repeating-rows-and-columns-in-excel/>***

When dealing with multi-page Excel spreadsheets, readability improves significantly when the column or row headings are printed on each page. These headings are often referred to as _Repeating Rows and Columns_ or _Header Rows and Columns_ in Excel. Using IronXL, you can easily implement this feature with just a few lines of code.

## `SetRepeatingRows(startRow, endRow)`

This method configures rows to repeat on multiple pages. For instance, `workSheet.SetRepeatingRows(3, 4)` ensures that rows 4 and 5 (under zero-based indexing) are repeated across pages.

## `SetRepeatingColumns(startColumn, endColumn)`

Similarly, this function is designed to define repeating columns. For example, the command `workSheet.SetRepeatingColumns(0, 2)` will cause columns A to C to repeat across pages.

Please remember, these methods accept zero-based indices, meaning column(0) corresponds to "A" and row(1) corresponds to 2\. Also, note that using these methods in tandem, like the examples provided, will have specific interactions.

Content that extends multiple pages horizontally from the right side of the first page will obey the repeating column rules exclusively. The illustration below demonstrates this:

## Figure 1

![Figure 1](https://ironsoftware.com/static-assets/excel/examples/repeating-rows-and-columns-in-excel/repeating-rows-and-columns-in-excel-2.webp)

Content that spans vertically from the bottom side of the first page will follow the repeating row rules.

Finally, pages situated within the inside part will adhere to both repeating column and row rules. Figure 2 captures these last two scenarios:

## Figure 2

![Figure 2](https://ironsoftware.com/static-assets/excel/examples/repeating-rows-and-columns-in-excel/repeating-rows-and-columns-in-excel-3.webp)