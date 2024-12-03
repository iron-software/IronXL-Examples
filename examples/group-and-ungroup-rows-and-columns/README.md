***Based on <https://ironsoftware.com/examples/group-and-ungroup-rows-and-columns/>***

Streamline the process of grouping and ungrouping rows and columns without the need for Office Interop using the **IronXL** library's intuitive APIs.

## Group and Ungroup Rows

To group rows, use the `GroupRows` method, which requires the index positions of the rows you want to group. You can apply multiple groupings by invoking this method multiple times.

For ungrouping, the `UngroupRows` method removes specific groupings, acting akin to a cutting tool. When applied to the center of a row group, it splits the group into two. However, these new segments do not form new individual groups. For instance, ungrouping rows 2-4 in a grouped set of 0-9 will result in two segments: rows 0-1 and 5-9.

## Group and Ungroup Columns

Grouping columns works similarly to grouping rows. The `GroupColumns` method allows you to define column groups by specifying either the index positions or the names of the columns as strings. It is also possible to create multiple groups of columns.

Ungrouping columns with the `UngroupColumn` method similarly acts as a splitting mechanism. Applying it to the center of a column group divides it into two distinct groups. For example, ungrouping columns C-D from a grouped set of A-F will yield groups A-B and E-F.

It is important to note that all index positions are zero-based. Furthermore, groups can only be formed up to the last cell that contains a value.