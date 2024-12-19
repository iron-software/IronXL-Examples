# How to Combine and Separate Cells in Spreadsheets

***Based on <https://ironsoftware.com/how-to/csharp-excel-merge-cells/>***


Combining cells, more commonly known as cell merging, refers to the technique of fusing two or more neighboring cells into a single larger cell. Conversely, separating cells or unmerging refers to splitting a previously merged cell back into its original, individual cells. These functionalities enhance flexibility, ensure alignment uniformity, and aid in more effective data management.

IronXL supports the programmatic merging and unmerging of cells in spreadsheets.

***

***

<h3>Getting Started with IronXL</h3>

----------------------------------

## Example of Merging Cells

The `Merge` function allows for the combination of a cell range. This action amalgamates the cells while preserving all existing content, although only the contents of the first cell in the merged area will be visible. Nevertheless, the information from all cells in the merged area remains accessible in IronXL.

It's noteworthy that merging cells within certain ranges might introduce conflicts in the Excel document, potentially necessitating the use of an Excel repair tool to access the spreadsheet normally.

The following code snippet demonstrates how to merge a specified range of cells:

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

var range = workSheet["B2:B5"];

// Merging cells from B7 to E7
workSheet.Merge("B7:E7");

// Merging the selected range
workSheet.Merge(range.RangeAddressAsString);

workBook.SaveAs("mergedCell.xlsx");
```

### Visualization Example
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/merge-cells/merge-cells-merge.png" alt="Merge Cells Demonstration" class="img-responsive add-shadow">
    </div>
</div>

## Example of Retrieving Merged Regions

The ability to identify merged regions is especially useful in applications such as Microsoft Excel for interpreting the layout and data flow. The `GetMergedRegions` method facilitates the retrieval of these regions.

```cs
using IronXL;
using System.Collections.Generic;
using System;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Implementing merges
workSheet.Merge("B4:C4");
workSheet.Merge("A1:A4");
workSheet.Merge("A6:D9");

// Gathering merged regions
List<IronXL.Range> retrieveMergedRegions = workSheet.GetMergedRegions();

foreach (IronXL.Range mergedRegion in retrieveMergedRegions)
{
    Console.WriteLine(mergedRegion.RangeAddressAsString);
}
```

## Example of Unmerging Cells

To reverse the merge, you can either specify the precise cell addresses to be separated, such as *"B3:B6"*, or utilize the index of the merged region which can be identified via a list acquired previously.

The cell addresses must match exactly to the original merged region for successful separation.

The separation of partial merged regions is not supported.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("mergedCell.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Separating the merged region from B7 to E7
workSheet.Unmerge("B7:E7");

workBook.SaveAs("unmergedCell.xlsx");
```

### Visualization Example
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/merge-cells/merge-cells-unmerge.png" alt="Unmerge Cells Demonstration" class="img-responsive add-shadow">
    </div>
</div>