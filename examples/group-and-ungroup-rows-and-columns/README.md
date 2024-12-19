***Based on <https://ironsoftware.com/examples/group-and-ungroup-rows-and-columns/>***

Enable row and column grouping efficiently without Office Interop using the **IronXL** library.

## Grouping and Ungrouping Rows

Utilize the `GroupRows` method to group rows by specifying their index positions. You can create multiple groupings by invoking this method repeatedly.

For ungrouping, employ the `UngroupRows` method. Think of it as a way to split a group; applying it to a segment within a row group splits the group into smaller sections. For instance, if you use `UngroupRows(2,4)` on a group from rows 0 to 9, you'll end up with groups spanning rows 0-1 and 5-9.

## Grouping and Ungrouping Columns

Similarly, columns can be grouped using the `GroupColumns` method by specifying either the index positions or the names of the columns as `string` values. It's possible to create multiple column groups.

To ungroup columns, the `UngroupColumn` method works just like it does for rows. Applying it to the middle of a column group will divide it into two separate groups. For example, ungrouping columns C to D in a group from A to F will produce groups for columns A-B and E-F.

Please note that all the index positions mentioned are zero-based. Also, note that grouping can only be set up to the last cell that contains a value.