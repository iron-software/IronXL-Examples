***Based on <https://ironsoftware.com/examples/excel-sql-datatable/>***

Convert Excel and CSV file formats like XLSX, XLS, XLSM, XLTX, CSV, and TSV into a `System.Data.DataTable`. This conversion facilitates seamless interaction with `System.Data.SQL` or allows for easy filling of a `DataGrid`.

When using the `ToDataTable()` method, pass `true` to designate the first row as the header, which sets the column names in the `DataTable`. This structured data can then be used to populate a `DataGrid` effectively.