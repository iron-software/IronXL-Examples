Leverage the **IronXL** library to efficiently manage groupings and ungroupings of rows and columns without the need for Office Interop integration.

## Grouping and Ungrouping Rows

The `GroupRows` function enables you to cluster rows together based on their index positions. You can create multiple groups by invoking this method multiple times.

On the other hand, the `UngroupRows` function serves to dismantle these groupings. Think of this as slicing through a grouped segment. If you ungroup a subset within a larger grouping, such as ungrouping rows 2 through 4 in a group spanning rows 0 to 9, the resulting groups would be rows 0-1 and 5-9, where the original grouping is modified but remains as a single group overall.

## Grouping and Ungrouping Columns

Grouping columns works similarly to rows. The `GroupColumns` method allows you to define grouped columns either by their index or by specifying the column names directly as strings. Multi-level grouping is feasible by using this method consecutively.

For columns, the `UngroupColumn` function acts to split a group of columns into separate entities, functioning similarly to a precision cut. By ungrouping columns C through D in a grouped set from A to F, for instance, you would end up with two distinct groups: A-B and E-F.

Please be aware that all mentioned index positions are based on zero-based indexing. Furthermore, grouping operations are constrained by the presence of values in the cells, limiting the applicability of groupings.