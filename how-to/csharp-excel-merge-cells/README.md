# How to Merge and Unmerge Cells

***Based on <https://ironsoftware.com/how-to/csharp-excel-merge-cells/>***


Merging cells involves combining adjacent cells into a single cell that's larger, while unmerging involves splitting a previously merged area into the original separate cells. This functionality enhances alignment, promotes consistency, and facilitates easier data analysis.

IronXL provides the ability to both merge and unmerge cells within a spreadsheet through programming.

## Example of Merging Cells

To merge a cluster of cells, we utilize the `Merge` method. This function consolidates the designated cells, preserving their original content, though only the upper-left cell's content will be visible post-merging. All original cell values are still retained and can be accessed in IronXL.

It's important to note that merging cells within certain ranges might lead to file conflicts that require using Excel repair to open the compromised file.

Below is an example illustrating how to merge cells by specifying their range:

```cs
using IronXL;
using IronXL.Excel;

namespace ironxl.CsharpExcelMergeCells
{
    public class Section1
    {
        public void Run()
        {
            // Load an existing workbook
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            // Obtain the default worksheet
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            var range = workSheet["B2:B5"];
            
            // Merge cells from B7 to E7
            workSheet.Merge("B7:E7");
            
            // Merge another selected range
            workSheet.Merge(range.RangeAddressAsString);
            
            // Save the changes to a new file
            workBook.SaveAs("mergedCell.xlsx");
        }
    }
}
```

### Illustration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/merge-cells/merge-cells-merge.png" alt="Merge Cells Demonstration" class="img-responsive add-shadow">
    </div>
</div>

## How to Retrieve Merged Regions

Obtaining merged regions is essential to uncover which cell values appear in applications like Microsoft Excel. The `GetMergedRegions` method enables you to gather a list of these regions.

```cs
using System;
using IronXL.Excel;

namespace ironxl.CsharpExcelMergeCells
{
    public class Section2
    {
        public void Run()
        {
            // Create a new workbook
            WorkBook workBook = WorkBook.Create();
            // Obtain the default worksheet
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Execute merges on specific ranges
            workSheet.Merge("B4:C4");
            workSheet.Merge("A1:A4");
            workSheet.Merge("A6:D9");
            
            // Fetch all merged regions
            List<IronXL.Range> retrieveMergedRegions = workSheet.GetMergedRegions();
            
            // Output addresses of all merged regions
            foreach (IronXL.Range mergedRegion in retrieveMergedRegions)
            {
                Console.WriteLine(mergedRegion.RangeAddressAsString);
            }
        }
    }
}
```

## How to Unmerge Cells

Unmerging can be accomplished using either directly specified cell ranges or by using an index of previously merged regions.

```cs
using IronXL;
using IronXL.Excel;

namespace ironxl.CsharpExcelMergeCells
{
    public class Section3
    {
        public void Run()
        {
            // Load a workbook containing merged cells
            WorkBook workBook = WorkBook.Load("mergedCell.xlsx");
            // Access the default worksheet
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Unmerge a specific region by specifying its range
            workSheet.Unmerge("B7:E7");
            
            // Save the unmerged workbook
            workBook.SaveAs("unmergedCell.xlsx");
        }
    }
}
```

### Illustration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/merge-cells/merge-cells-unmerge.png" alt="Unmerge Cells Demonstration" class="img-responsive add-shadow">
    </div>
</div>