***Based on <https://ironsoftware.com/how-to/c-sharp-parse-excel-file/>***

class="content-table parse-excel-file">
  <table>
    <tbody>
      <tr class="tr-head">
          <th class="tcol1">DataType</th>
          <th class="tcol2">Method</th>
          <th class="tcol3">Description</th>
      </tr>
      <tr>
          <td>Array</td>
          <td>WorkSheet ["From:To"].ToArray()</td>
          <td>This method enables the conversion of specified cell range data into an array format.</td>
      </tr>
      <tr>
          <td>DataTable</td>
          <td>WorkSheet.ToDataTable()</td>
          <td>Parses the whole worksheet into a DataTable, allowing for structured data utilization in .NET.</td>
      </tr>
      <tr>
          <td>DataSet</td>
          <td>WorkBook.ToDataSet()</td>
          <td>Converts entire Excel Workbook into a DataSet; each worksheet becomes a DataTable within the DataSet.</td>
      </tr>
    </tbody>
  </table>
</div>

Here are specific examples for each data structure conversion:

### 7.1. Convert Excel Data Into an Array

IronXL simplifies the process of converting a designated range of Excel data into an array:

```cs
var array = WorkSheet ["From:To"].ToArray();
```

To access a particular item from the array:

```cs
string item = array[ItemIndex].Value.ToString();
```

Example showcasing the conversion of a range into an array and accessing an element from it:

```cs
/**
Convert Range to Array
anchor-parse-excel-data-into-array
**/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    var array = ws["B6:F6"].ToArray();
    int item = array.Count();
    string total_items = array[0].Value.ToString();
    Console.WriteLine("First item in the array: {0}", item);
    Console.WriteLine("Total items from B6 to F6: {0}", total_items);
    Console.ReadKey();
}
```

Visual representation of the output:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

Confirmation of `sample.xlsx` data range visualized:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

### 7.2. Convert an Excel Worksheet Into a DataTable

With IronXL, converting a specific Excel Sheet into a DataTable is straightforward:

```cs
DataTable dt = WorkSheet.ToDataTable();
```

If the first row of the Excel file should be used as DataTable Column Names:

```cs
DataTable dt = WorkSheet.ToDataTable(True);
```

Explore further on managing [ExcelWorksheet as DataTable in C#](https://ironsoftware.com/csharp/excel/#excel-sql-datatable).

Here's the implementation for parsing a worksheet into a DataTable:

```cs
/**
Convert Worksheet to DataTable
anchor-parse-excel-worksheet-into-datatable
**/
using IronXL;
using System.Data; 
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    DataTable dt = ws.ToDataTable(true); 
}
```

### 7.3. Convert an Excel File into a DataSet

To convert a complete Excel file into a DataSet, which allows for handling multiple worksheets as separate DataTables:

```cs
DataSet ds = WorkBook.ToDataSet();
```

Here's how to access a DataTable from the DataSet, with each worksheet treated as individual DataTables:

```cs
/**
Convert File to DataSet
anchor-parse-excel-file-into-dataset
**/
using IronXL;
using System.Data; 
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    DataSet ds = wb.ToDataSet();
    DataTable dt = ds.Tables[0];
}
```

Further insights can be found at [Excel SQL Dataset](https://ironsoftware.com/csharp/excel/#excel-sql-dataset).

<hr class="separator">
<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>Documentation for Excel in C#</h3>
      <p>Explore detailed IronXL documentation to leverage extensive functionality for managing Excel in C# projects.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Documentation for Excel in C# <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>