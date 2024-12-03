***Based on <https://ironsoftware.com/examples/merge-and-unmerge-cells/>***

IronXL supports both merging and separating cells within a spreadsheet programmatically. The provided code snippet illustrates how to seamlessly merge a specific range of cells by defining their addresses.

## Merge

Using the `Merge` method, it is possible to combine multiple cells into a single cell. This action keeps all original cell values intact, but only the value located in the top-left cell of the merged area will be displayed. Nonetheless, all data from the merged cells remain retrievable in IronXL.

## Unmerge

To separate previously merged cells, there are a couple of strategies you can employ. The most direct method is to identify the cell range, such as "D1:D3", ensuring that the range matches exactly with the initially merged area. It is important to recognize that it is not feasible to unmerge just a segment of a merged region.

Alternatively, you can unmerge cells by referring to the index of the merged regions. These indices are arranged in the order they were merged. Currently, there is no capability to query the comprehensive list of merged regions on the platform.