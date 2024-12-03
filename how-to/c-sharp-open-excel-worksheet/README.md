# C# Open Excel Worksheets

***Based on <https://ironsoftware.com/how-to/c-sharp-open-excel-worksheet/>***


Delve into the use of C# to open and manipulate Excel Worksheets, including files such as `.xls`, `.csv`, `.tsv`, and `.xlsx`. Opening an Excel Worksheet, retrieving its content, and manipulating it programmatically are crucial tasks for developers working on applications that require data handling. Here, we present an efficient and code-minimal approach for developers targeting swift execution and simplified coding requirements.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Opening an Excel File in C#</h2>
      <ul class="list-unstyled">
        <li>Install a C# library to handle Excel files</li>
        <li><a href="#anchor-2-load-excel-file">Load the Excel file into the <strong>WorkBook</strong> object</a></li>
        <li><a href="#anchor-3-open-excel-worksheet">Discover multiple methods to select a <strong>WorkSheet</strong> in the loaded Excel file</a></li>
        <li><a href="#anchor-4-get-data-from-worksheet">Access data through the selected <strong>WorkSheet</strong> object</a></li>
        <li><a href="#anchor-4-3-get-data-from-row">Retrieve data within a specified range of rows and columns</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h2>Steps to Open an Excel Worksheet in C#</h2>

1. Install an Excel library to enable reading Excel files.
2. Load the Excel file into a `Workbook` object.
3. Activate a chosen Excel Worksheet.
4. Extract data from the Excel Workbook.
5. Handle and display the extracted data.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Incorporating Excel C# Library

Integrate your project with the [Excel C# Library via DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.excel.worksheet.zip) or install it through your favorite [NuGet manager](https://www.nuget.org/packages/IronXL.Excel). Once included, the IronXL library provides a comprehensive set of functions to manage Excel Worksheets effectively.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. File Loading

Utilize `WorkBook.Load()` method from IronXL to import an Excel file into the application. You'll need to provide a string parameter indicating the file's path:

```cs
WorkBook wb = WorkBook.Load("YourExcelFilePath"); // Specify the path to your Excel file
```

By specifying the path, `wb` will now contain the loaded Excel file. Next, define the Worksheet to be used.

<hr class="separator">

## 3. Accessing Excel WorkSheet

Utilize IronXL's `WorkBook.GetWorkSheet()` method to access a particular `WorkSheet` by its name:

```cs
WorkSheet ws = wb.GetWorkSheet("YourSheetName"); // Replace 'YourSheetName' with the actual sheet name
```

This will allow access to `ws` containing all needed data. Multiple methods to access various `WorkSheet`s include:

```cs
// By sheet index
WorkSheet ws = wb.WorkSheets[0];               // Accessing the first sheet
// Default sheet
WorkSheet ws = wb.DefaultWorkSheet;            // The default worksheet
// First sheet explicitly
WorkSheet ws = wb.WorkSheets.First();
// First or default sheet
WorkSheet ws = wb.WorkSheets.FirstOrDefault();
```

Continue by retrieving data from the chosen Excel `WorkSheet`.

<hr class="separator">

## 4. Extracting Data from the Worksheet

Data can be extracted from an active Excel `WorkSheet` in several ways:

1. Retrieve a specific cell's value.
2. Fetch data from a defined range.
3. Extract all data present in the `WorkSheet`.

Explore these methods as follows:

### 4.1. Retrieve Specific Cell Value

To extract specific cell data, you can directly reference the cell address:

```cs
string value = ws["CellAddress"].ToString(); // Replace 'CellAddress' with the actual address like 'A1'
```

Additionally, access data using row and column indices:

```cs
string value = ws.Rows[RowIndex].Columns[ColumnIndex].Value.ToString(); // Specify RowIndex and ColumnIndex
```

The following example demonstrates opening an Excel file in C# and extracting specific cell values:

```cs
using IronXL;
static void Main(string[] args)
{
    // Load the Excel file
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Access a specific WorkSheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Extract value using cell address
    int intValue = ws["C6"].Int32Value;
    // Extract value using row and column indices
    string strValue = ws.Rows[3].Columns[1].Value.ToString();

    Console.WriteLine($"Value obtained by Cell Address: {intValue}");
    Console.WriteLine($"Value obtained by Row and Column Indices: {strValue}");
    Console.ReadKey();
}
```

The code produces the output as illustrated below:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/1output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/1output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

**Value from the Excel file `sample.xlsx` in `Row [3].Column [1]` and `Cell C6`**:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

Rows and columns are indexed starting from `0`.

Additionally, learn more about accessing data directly from Excel WorkSheets using C# by visiting [Read Excel Data in C#](https://ironsoftware.com/csharp/excel/#read-excel).

### 4.2. Extract Data Within a Specific Range

IronXL offers an effective method to specify a range and extract data accordingly. Simply define the 'from' and 'to' cell addresses:

```cs
var rangeData = ws["StartCellAddress:EndCellAddress"]; // Replace 'StartCellAddress' and 'EndCellAddress'
```

Here's how you might apply this method to get data from an open `WorkSheet`:

```cs
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Define and iterate through the specified range
    foreach (var cell in ws["B2:B10"])
    {
        Console.WriteLine($"Value is: {cell.Text}");
    }
    Console.ReadKey();
}
```

This script retrieves data from the cells `B2` to `B10`, and the output appears as follows:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/2output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/2output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

The values from Excel file `sample.xlsx` in the range `B2 to B10` are visible here:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-open-excel-worksheet/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

### 4.3. Specific Row Data

Additionally, you can define a range across an entire row like this:

```cs
var rowRange = ws["A1:E1"]; // Data from 'A1' to 'E1'
```

For further details on handling various row and column formats, visit [C# Excel Ranges](https://ironsoftware.com/csharp/excel/#excel-ranges).

### 4.4. Extract All Data from a WorkSheet

Extracting all cell data from an open Excel `WorkSheet` using IronXL is straightforward. Here's how you can navigate through each cell in a row-by-row and column-by-column manner:

```cs
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample2.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Traverse through each row
    for (int i = 0; i < ws.Rows.Count(); i++)
    {    
        // Traverse through each column in the current row
        for (int j = 0; j < ws.Columns.Count(); j++)
    {
            // Access the value of each cell
            Console.WriteLine(ws.Rows[i].Columns[j].Value.ToString());
        }
    }
    Console.ReadKey();
}
```

The above code fetches and displays each cell value within the fully loaded Excel `WorkSheet`.

<hr class="separator">
<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100%, height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>API Reference Resource</h3>
      <p>Explore the IronXL API Reference for comprehensive details on available functions, classes, namespaces, method fields, enums, and features for your projects.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Explore API Reference <i class="fa fa-chevron-right"></i></a>
    </div>
  </div>
</div>