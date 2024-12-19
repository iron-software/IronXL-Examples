***Based on <https://ironsoftware.com/examples/merge-and-unmerge-cells/>***

IronXL provides functionality for both merging and unmerging cells in spreadsheets programmatically. Below is a detailed explanation of how cells can be merged by specifying their addresses in the spreadsheet.

## Merge

The `Merge` method allows for the combination of a range of cells into a single cell. During this merging process, no data is deleted; however, only the value from the first cell in the specified range will be displayed. The contents of all merged cells remain accessible through IronXL.

## Unmerge

To revert a merged cell back to individual cells, there are a couple of methods available. The simplest method is to directly specify the cell range that was originally merged, like "D1:D3". It is important to note that the specified range must match the originally merged range exactly, and partial unmerging within a region is not supported.

Alternatively, unmerging can be performed based on the index of the merged region. Merged regions are stored in a list that is maintained in the order they were created. Currently, it is not feasible to retrieve or interact with the list of all merged regions.