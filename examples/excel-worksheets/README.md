***Based on <https://ironsoftware.com/examples/excel-worksheets/>***

The **IronXL** library simplifies the management of Excel worksheets using C#. It offers functionalities such as creating and deleting worksheets, rearranging them, and setting a default active worksheet, all without the need for Office Interop.

## Create a Worksheet

Creating a new worksheet is streamlined with the `CreateWorkSheet` method. This method requires just one parameter: the name of the new worksheet.

## Adjust Worksheet Position

To reposition a worksheet within a workbook, use the `SetSheetPosition` method. This method requires two parameters: the worksheet's name (as a String) and its new position index (as an Integer).

## Activate a Worksheet

Activating a worksheet sets it as the first visible tab when the Excel workbook is opened. This is done using the `SetActiveTab` method, which requires the worksheet's index position.

## Delete a Worksheet

The `RemoveWorkSheet` method facilitates the removal of a worksheet. It uses the worksheet's index position for deletion. If the index is unknown, the worksheet's name can alternatively be used to perform the deletion.

Note: All index positions referred to above are zero-based.