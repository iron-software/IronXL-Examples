***Based on <https://ironsoftware.com/examples/insert-new-rows-and-columns/>***

**IronXL** library facilitates the insertion of single and multiple rows and columns directly through C# code, circumventing the need for Office Interop.

### Insert Row

The methods `InsertRow` and `InsertRows` make it feasible to insert new rows. These functions add row(s) at the position prior to the given index.

### Insert Column

To introduce new column(s), you can use the `InsertColumn` and `InsertColumns` methods, which place the column(s) before the designated index.

It's important to note that all index positions referred to above are zero-based indexes.