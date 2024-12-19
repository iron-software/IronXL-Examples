***Based on <https://ironsoftware.com/examples/auto-resize-rows-and-columns/>***

Adjusting the dimensions of rows and columns in a spreadsheet can enhance its readability and conserve space. The `IronXL` C# library supports the functionality to automatically resize rows and columns. Utilizing C#, these resize operations can be applied to all existing rows and columns, automating what would otherwise be a manual adjustment in the spreadsheet.

## Auto Resize Rows

The `AutoSizeRow` method dynamically alters the height of a specified row to fit the content length. An additional overload of the `AutoSizeRow` method accepts a _Boolean_ parameter. Setting this parameter to `true` will adjust the height considering the dimensions of any merged cells within the row.

## Auto Resize Columns

To modify column width automatically, the `AutoSizeColumn` method adjusts the width of one or more columns based on the content length. Similar to `AutoSizeRow`, this feature can also accommodate the dimensions of merged cells when resizing columns.

Please be aware that the indexes for both rows and columns start from zero.