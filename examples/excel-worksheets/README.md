**IronXL Library Overview**

IronXL library streamlines the manipulation of worksheets using C# code, rendering tasks like creation and deletion of worksheets, repositioning, and setting an active worksheet in an Excel file straightforward without the need for Office Interop.

## Creating a Worksheet

The `CreateWorkSheet` function facilitates the creation of a new worksheet. This method requires solely the name of the worksheet as a parameter.

## Adjusting Worksheet Position

With the `SetSheetPosition` method, you can rearrange or relocate a worksheet within a workbook. This method necessitates two parameters: the worksheet name as a `string` and its new index position as an `int`.

## Activating a Worksheet

To determine which worksheet appears first upon opening a workbook, use the `SetActiveTab` method with the worksheet’s index position as the parameter.

## Deleting a Worksheet

Worksheets can be efficiently removed in IronXL using the `RemoveWorkSheet` method. This method allows for the deletion of a worksheet by specifying either its index position or its name when the position is uncertain.

Please be aware that the index positions mentioned are zero-based, meaning they start counting from zero.