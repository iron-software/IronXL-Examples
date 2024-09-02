IronXL supports the programmable merging and unmerging of cells in a spreadsheet. The code example provided demonstrates how you can effortlessly merge a range of cells by indicating the cell addresses.

## Merge

Utilizing the `Merge` method allows you to combine a specified range of cells. During the merge process, no data from the cells are deleted; however, only the value from the first cell in both the row and column of the merged area will be displayed. The data from all merged cells remains intact and accessible within IronXL.

## Unmerge

To split a previously merged cell region, there are two methods available. The most direct approach is by specifying the exact cell range that was merged, for example, `"D1:D3"`. It is important to ensure that the address corresponds exactly to the merged region as partial unmerging within a merged region is not supported.

Alternatively, unmerging can be done based on the index of the merged region. The list of merged regions is maintained in chronological order. However, it is important to note that retrieving this list is not currently supported.