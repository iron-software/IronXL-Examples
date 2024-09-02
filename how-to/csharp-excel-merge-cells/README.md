# How to Merge and Unmerge Cells

Merging cells is the process by which multiple adjacent cells are combined into one larger cell. Conversely, unmerging cells involves splitting a previously merged cell back into its original, separate cells. These functionalities enhance flexibility, improve alignment and facilitate better analysis of data.

IronXL supports the programmatic merging and unmerging of cells in a spreadsheet.

***

***

## Example of Merging Cells

To merge cells, the `Merge` method is used. It combines specified cells while retaining all original data, though only the value from the first cell in the merged area is displayed. All values remain accessible via IronXL.

Be cautious when merging cells within a filter range as it may lead to conflicts that require running Excel's repair tool to regain a normal view of the spreadsheet.

Below is an example showing how to merge cells using specific cell addresses.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

var range = workSheet["B2:B5"];

// Merge cells from B7 to E7
workSheet.Merge("B7:E7");

// Merge a specified range
workSheet.Merge(range.RangeAddressAsString);

workBook.SaveAs("mergedCell.xlsx");
```

### Visual Demonstration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/merge-cells/merge-cells-merge.png" alt="Merge Cells Demonstration" class="img-responsive add-shadow">
    </div>
</div>

## Example of Retrieving Merged Regions

Identifying merged regions is vital for understanding the visible data in applications like Microsoft Excel. The `GetMergedRegions` method allows for the listing of these regions.

```cs
using IronXL;
using System.Collections.Generic;
using System;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Implement merging
workSheet.Merge("B4:C4");
workSheet.Merge("A1:A4");
workSheet.Merge("A6:D9");

// Extract merged regions
List<IronXL.Range> mergedRegionsList = workSheet.GetMergedRegions();

foreach (IronXL.Range mergedRegion in mergedRegionsList)
{
    Console.WriteLine(mergedRegion.RangeAddressAsString);
}
```

## Example of Unmerging Cells

To unmerge cells, you can either directly specify the cell range to be unmerged or use an index from the list of merged regions. The regions are indexed in the order they were merged.

It should be noted that only complete merged regions can be unmerged; partial unmerging is not supported.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("mergedCell.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Disband the merged region stretching from B7 to E7
workSheet.Unmerge("B7:E7");

workBook.SaveAs("unmergedCell.xlsx");
```

### Visual Demonstration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/merge-cells/merge-cells-unmerge.png" alt="Unmerge Cells Demonstration" class="img-responsive add-shadow">
    </div>
</div>