***Based on <https://ironsoftware.com/examples/excel-sql-datatable/>***

Transform various spreadsheet formats such as XLSX, XLS, XLSM, XLTX, CSV, and TSV into a `System.Data.DataTable`. This conversion facilitates seamless integration with `System.Data.SQL` or enhances the ability to fill a `DataGrid`.

Set the first row as the header by passing `true` to the `ToDataTable` method, effectively using the first row's values as column names within the `DataTable`. This `DataTable` can then be used to populate a `DataGrid`.