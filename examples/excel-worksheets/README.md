***Based on <https://ironsoftware.com/examples/excel-worksheets/>***

**IronXL** library simplifies managing worksheets in C# by providing operations to create, delete, reposition, and activate worksheets without needing Office Interop.

## Create Worksheet

Using the `CreateWorkSheet` method, you can easily create new worksheets. This function requires only the name of the worksheet as its parameter.

## Set Worksheet Position

To reorder or move a worksheet within a workbook, use the `SetSheetPosition` method. This method requires two parameters: the name of the worksheet (as a string) and its new position index (as an integer).

## Set Active Worksheet

To specify which worksheet opens by default when the workbook is accessed, employ the `SetActiveTab` method. Pass the index position of the worksheet you want to activate.

## Remove Worksheet

The `RemoveWorkSheet` method in IronXL allows for the removal of worksheets. You can specify the worksheet either by its index position for known locations or by its name if the position is unknown.

It's important to note that all the index positions referred to here are zero-based, meaning they start counting from zero.