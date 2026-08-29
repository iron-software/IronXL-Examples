> Full guide: [Excel sql datatable](https://ironsoftware.com/csharp/excel/examples/excel-sql-datatable/?utm_source=github)

Convert Excel and CSV file formats like XLSX, XLS, XLSM, XLTX, CSV, and TSV into a `System.Data.DataTable`. This conversion facilitates interaction with `System.Data.SQL` or allows for easy filling of a `DataGrid`.

When using the `ToDataTable()` method, pass `true` to designate the first row as the header, which sets the column names in the `DataTable`. This structured data can then be used to populate a `DataGrid` effectively.