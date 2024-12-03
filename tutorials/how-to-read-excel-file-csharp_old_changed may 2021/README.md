# C# Excel File Reading Examples

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp_old_changed may 2021/>***


This guide demonstrates how to extract data from Excel files using C#, and how to deploy the IronXL library for common operations such as data validation, transferring data to databases, storing results from Web APIs, and altering formulas inside Excel files. The examples illustrated in this document are implemented using IronXL within a .NET Core Console application.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>How To Read Excel Files in C# .NET:</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-the-ironxl-c-library-for-free">Download the IronXL Library for C#</a></li>
        <li><a href="#anchor-3-create-a-workbook">Create a Workbook</a></li>
        <li><a href="#anchor-4-edit-cell-values-within-a-range">Edit Cell Values Within a Range</a></li>
        <li><a href="#anchor-8-export-data-using-entity-framework">Export Data using Entity Framework</a></li>
        <li><a href="#anchor-9-add-formulae-to-a-spreadsheet">Add Formulae to Spreadsheets and more</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <a href="/downloads/assets/excel/tutorials/how-to-read-excel-file-csharp/tutorial-read-excel.pdf" target="_blank">
          <img style="box-shadow: none; width: 308px; height: 320px;" src="/img/tutorials/how-to-read-excel-file-csharp/how-to-read.svg" data-hover-src="/img/tutorials/how-to-read-excel-file-csharp/how-to-read-hover.svg" class="img-responsive learn-how-to-img replaceable-img">
        </a>
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<a name ="video"></a><h4 class="tutorial-segment-title">Overview</h4>
<h2>Read Data from Excel in .NET using IronXL</h2>



<iframe class="lazy" width="100%" height="450" data-src="https://www.youtube.com/embed/sXkcdZWUcWI?rel=0" frameborder="0" allow="accelerometer; encrypted-media; gyroscope picture-in-picture" allowfullscreen></iframe>

IronXL is a comprehensive .NET library designed for managing and manipulating Microsoft Excel files using C#. In this tutorial, you'll learn how to harness C# to read Excel documents efficiently.

1. Begin by installing the IronXL Excel Library, which can be achieved either through our [NuGet package](https://www.nuget.org/packages/IronXL.Excel/) or by directly downloading the [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

2. Utilize the `WorkBook.Load` method to open any XLS, XLSX, or CSV file.

3. Retrieve cell values effortlessly with user-friendly syntax, for example: `sheet["A11"].DecimalValue`.

<h3>IronXL Includes:</h3>

Here's a paraphrased version of the selected section:

-----
- Receive dedicated support from our team of .NET engineering experts.

- Hassle-free setup through Microsoft Visual Studio.

- Complimentary for development purposes. Pricing starts from `$liteLicense`.

Experience the simplicity of processing **Excel** files using C# or VB.NET with the IronXL library, which includes three example spreadsheets.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<h4>Read XLS or XLSX Files: Quick Code</h4>

In this illustration, it's evident that *Excel* files can be effectively accessed without using Interop in C#. Additionally, the Advanced Operations demonstrate compatibility with **Linq** and the capability to perform aggregate calculations across a range.

Here's the paraphrased section of the article with updated code snippets:

```cs
/**
Loading and Reading XLS or XLSX Files
anchor-loading-and-reading-xls-xlsx-files
**/
using IronXL;
using System.Linq;

// The following formats are supported for reading: XLSX, XLS, CSV, and TSV
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.WorkSheets.First();
// Directly access cell values using Excel cell references
int specificCellValue = sheet["A2"].IntValue;
// Iterate over a range of cells and print their values
foreach (var cell in sheet["A2:A10"])
{
    Console.WriteLine("Cell at {0} contains: '{1}'", cell.AddressString, cell.Text);
}


///Complex Operations

// Compute summary statistics for a range
decimal totalSum = sheet["A2:A10"].Sum();
// Utilizing LINQ to find maximum value within a range
decimal maximumValue = sheet["A2:A10"].Max(c => c.DecimalValue);
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## Get Started with IronXL: Free C# Library Download

The initial step to integrate Excel capabilities into your .NET application is to download and install the IronXL.Excel library. This provides the essential tools for handling Excel files in your projects.

<h3>Installing the IronXL NuGet Package</h3>

1. Open Visual Studio, and in the context menu of your project, choose "Manage NuGet Packages..."

2. Look up the `IronXL.Excel` package in the search bar and proceed with the installation.
```

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
  <p><img src="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<br>

Here's the paraphrased section with all relative URL paths resolved to `ironsoftware.com`:

-----

Alternatively, you can begin installation via the following steps:

1. Open the Package Manager Console

2. Input the command: `Install-Package IronXL.Excel`

```shell
Install-Package IronXL.Excel
```

<br>

Moreover, you may also [inspect the package on the NuGet website](https://www.nuget.org/packages/IronXL.Excel/) to explore additional details.

<h3>Direct Download Installation</h3>

Alternatively, you can also initiate by downloading the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually integrating it into Visual Studio.

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Opening an Excel File ##

The [`WorkBook`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) class encapsulates an Excel file. To access an Excel document, utilize the `WorkBook.Load` method and provide the file path of the Excel document (.xlsx).
```

```cs
// Loading the Excel WorkBook
// Anchor for reference: load-workbook-example
var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

Every `WorkBook` can hold several <a href="https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkSheet.html" target="_blank">`WorkSheet`</a> instances, each corresponding to a worksheet within an Excel file. If the workbook includes multiple worksheets, you can fetch them individually by their names using <a href="https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html" target="_blank">`WorkBook.GetWorkSheet`</a>.

```cs
var worksheet = workbook.GetSheet("GDPByCountry");
```

Here's the paraphrased section of the article with resolved relative URL paths:

## Export Data Using Entity Framework

IronXL supports integrating Excel data with a database through various methods, including usage with Entity Framework. Below is an example that demonstrates how you can extract and export data from an Excel sheet and import it directly into an SQLite database using Entity Framework.

Ensure the necessary SQLite Entity Framework NuGet packages are integrated with your project.

![IronXL NuGet](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)

Entity Framework facilitates the creation of data model classes that mirror database structures. Below is an example of a model class:

```csharp
public class Country
{
    [Key]
    public Guid Key { get; set; }
    public string Name { get; set; }
    public decimal GDP { get; set; }
}
```

Next, configure the database context. If you're using a different database system, adapt the connection string and provider:

```csharp
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        Database.EnsureCreated(); // Ensure the database is created
    }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection($"Data Source=Country.db");
        connection.Open(); // Open the connection

        var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery(); // Activate foreign keys

        optionsBuilder.UseSqlite(connection); // Use SQLite

        base.OnConfiguring(optionsBuilder);
    }
}
```

Configure an instance of `CountryContext`, parse each Excel row, and asynchronously save the data into the database.

```csharp
public async Task ProcessAsync()
{
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    using (var countryContext = new CountryContext())
    {
        for (var i = 2; i <= 213; i++)
        {
            var range = worksheet[$"A{i}:B{i}"].ToList();

            var country = new Country
            {
                Name = (string)range[0].Value,
                GDP = (decimal)(double)range[1].Value
            };

            await countryContext.Countries.AddAsync(country); // Add country to the context
        }

        await countryContext.SaveChangesAsync(); // Save changes asynchronously
    }
}
```

This example demonstrates how data from an Excel file can be exported to a database via IronXL's functionality and Entity Framework's data manipulation capacities.

<hr class="separator">

## 3. Generate a New WorkBook ##

To instantiate a new WorkBook in memory, you need to initialize a new `WorkBook` object specifying the type of spreadsheet you intend to work with.

```cs
/**
Initialize a New WorkBook
anchor-initialize-a-new-workbook
**/
var workbook = new WorkBook(ExcelFileFormat.XLSX);
```

# C# Read Excel Files with Examples

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp_old_changed may 2021/>***


This tutorial clarifies how to interface with Excel data in C# and delves into daily tasks using the IronXL library, such as data validation, converting to database formats, preserving data from Web APIs, and amending formulas in spreadsheets. We're utilizing IronXL through sample code provided in a .NET Core Console App.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Steps to Read Excel Files in C# .NET:</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-the-ironxl-c-library-for-free">Acquire the IronXL Library for C#</a></li>
        <li><a href="#anchor-3-create-a-workbook">Initiate a Workbook</a></li>
        <li><a href="#anchor-4-edit-cell-values-within-a-range">Modify Cell Values Within a Range</a></li>
        <li><a href="#anchor-8-export-data-using-entity-framework">Export Data using Entity Framework</a></li>
        <li><a href="#anchor-9-add-formulae-to-a-spreadsheet">Incorporate Formulas into Spreadsheets and beyond</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <a href="https://ironsoftware.com/downloads/assets/excel/tutorials/how-to-read-excel-file-csharp/tutorial-read-excel.pdf" target="_blank">
          <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/how-to-read.svg" data-hover-src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/how-to-read-hover.svg" class="img-responsive learn-how-to-img replaceable-img">
        </a>
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<a name ="video"></a><h4 class="tutorial-segment-title">Overview</h4>
<h2>Utilize IronXL to Read Data from Excel in .NET</h2>


<iframe class="lazy" width="100%" height="450" data-src="https://www.youtube.com/embed/sXkcdZWUcWI?rel=0" frameborder="0" allow="accelerometer; encrypted-media; gyroscope picture-in-picture" allowfullscreen></iframe>

IronXL serves as a robust .NET library supporting C# for reading and manipulating Microsoft Excel files. Follow along this tutorial to learn how to employ C# code to <a href="https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/">read Excel databooks</a>.

1. Setup IronXL by installing it via <a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">NuGet package</a> or downloading the <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">.NET Excel DLL</a>.
2. Open any XLS, XLSX, or CSV document using the `.Load` method on a `WorkBook`.
3. Access individual cells using easy-to-understand syntax: <code>sheet["A11"].DecimalValue</code>

<h3>Key Features of IronXL:</h3>

- Full support provided by our dedicated .NET engineers
- Smooth integration within Microsoft Visual Studio
- Development use is FREE, with various licensing tiers starting from `$liteLicense`.

Experience how straightforward it is to read **Excel** files using C# or VB.NET with IronXL. Included are samples with three different Excel spreadsheets.

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<h4>Quick Example: Read XLS or XLSX Files</h4>

Discover how to process *Excel* files efficiently in C# without the need for Interop. The following advances demonstrate compatibility with **Linq** and how to perform aggregate math across ranges.

```cs
/**
Read XLS or XLSX File Example
anchor-read-an-xls-or-xlsx-file
**/
using IronXL;
using System.Linq;
    
// Supported formats include: XLSX, XLS, CSV, and TSV
WorkBook workbook = WorkBook.Load("example.xlsx");
WorkSheet sheet = workbook.WorkSheets.First();

// Easily select cells in Excel notation to get the computed value
int cellValue = sheet["A2"].IntValue;

// Elegantly read from a range of cells
foreach (var cell in sheet["A2:A10"])
{
    Console.WriteLine("Cell {0} holds value '{1}'", cell.AddressString, cell.Text);
}

///Advanced Operations

// Compute aggregate values like Min, Max, and Sum
decimal sum = sheet["A2:A10"].Sum();
// Compatible with Linq
decimal max = sheet["A2:A10"].Max(c => c.DecimalValue);
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

##  1. Get the IronXL C# Library for FREE

Start by installing the IronXL.Excel library to empower your .NET projects with Excel capabilities.

<h3>Installation via NuGet Package</h3>

1. Right-click on your project in Visual Studio, and select "Manage NuGet Packages..."
2. Type `IronXL.Excel` into the search box and install it.
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
  <p><img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<br>
Alternatively, you can:

1. Go to the Package Manager Console
2. Execute > `Install-Package IronXL.Excel`

```shell
Install-Package IronXL.Excel
```

<br>
Feel free to <a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">view the package on the NuGet website.</a>

<h3>Manual Installation via Direct Download</h3>

Alternatively, begin by downloading the IronXL <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">.NET Excel DLL</a> and manually integrate it into your Visual Studio setup.

<hr class="separator">

## 4. Creating a Worksheet

A "WorkBook" in Excel can contain several "WorkSheets," which serve as individual data sheets. A WorkBook essentially serves as a binder that keeps multiple WorkSheets together. This is demonstrated below through an example of a workbook comprising two distinct worksheets:

<center>
  ![Example Workbook with Two Worksheets](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/work-book.png)
</center>

Each WorkBook organizes its data across these multiple WorkSheets, allowing for better segmentation and data management within Excel.

<center>
  <a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
  </a>
</center>

To generate a new `WorkSheet`, invoke the `WorkBook.CreateWorkSheet` method and provide the desired name of the worksheet.

Here is the paraphrased section of the article, with the relative URL paths resolved to `ironsoftware.com`:

```cs
var sheet = workbook.NewSheet("Countries");
```

<hr class="separator">

## 5. Accessing a Range of Cells

The `Range` class encapsulates a two-dimensional array of `Cell` instances, effectively covering a specified range within an Excel worksheet. To retrieve a specific range, you can utilize the string indexer provided by a `WorkSheet` object. For example, you might specify a single cell like "A1" or a broader range of cells like "B2:E5". Additionally, the method `GetRange` can be directly invoked on a `WorkSheet` to achieve the same result.

```cs
// Obtaining a range of cells from D2 to D101 in the worksheet
var range = worksheet["D2:D101"];
```

```cs
/**
Validate Spreadsheet Data Example
anchor-validate-spreadsheet-data
**/
// Loop through each row starting from row 2
for (var rowIndex = 2; rowIndex <= 101; rowIndex++)
{
    var validationResults = new PersonValidationResult { Row = rowIndex };
    results.Add(validationResults);

    // Fetch all cell data for each person
    var personCells = worksheet [$"A{rowIndex}:E{rowIndex}"].ToList();

    // Validate phone number data located in column B (index 1)
    var phone = personCells[1].Value;
    validationResults.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Validate email data found in column D (index 3)
    validationResults.EmailErrorMessage = ValidateEmailAddress((string)personCells[3].Value);

    // Extract and validate date in text form from column E (index 4)
    var dateText = (string)personCells[4].Value;
    validationResults.DateErrorMessage = ValidateDate(dateText);
}

// This segment generates a new worksheet called "Results"
var resultsSheet = workbook.CreateWorkSheet("Results");

// Setting up the headers in the "Results" sheet
resultsSheet ["A1"].Value = "Row";
resultsSheet ["B1"].Value = "Valid";
resultsSheet ["C1"].Value = "Phone Error";
resultsSheet ["D1"].Value = "Email Error";
resultsSheet ["E1"].Value = "Date Error";

// Populate the "Results" sheet with errors found in validation
for (var index = 0; index < results.Count; index++)
{
    var validationResult = results[index];
    resultsSheet [$"A{index + 2}"].Value = validationResult.Row;
    resultsSheet [$"B{index + 2}"].Value = validationResult.IsValid ? "Yes" : "No";
    resultsSheet [$"C{index + 2}"].Value = validationResult.PhoneNumberErrorMessage;
    resultsSheet [$"D{index + 2}"].Value = validationResult.EmailErrorMessage;
    resultsSheet [$"E{index + 2}"].Value = validationResult.DateErrorMessage;
}

// Saves the workbook with the new "Results" sheet containing validation output
workbook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
```

<hr class="separator">

## 6. Modify Cell Content Within a Specified Range ##

You have numerous options to manipulate or view the content of cells in a specified range. When the number of cells is predetermined, employing a `For` loop is effective.

```cs
/**
Edit Values Within a Cell Range
anchor-edit-values-within-cell-range
**/
// Loop through the specified rows
for (int rowIndex = 2; rowIndex <= 101; rowIndex++)
{
    var validationResult = new PersonValidationResult { Row = rowIndex };
    results.Add(validationResult);

    // Retrieve all cells for a specific person
    var personCells = worksheet[$"A{rowIndex}:E{rowIndex}"].ToList();

    // Validating the phone number (Column B)
    var phone = personCells[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Validating the email address (Column D)
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)personCells[3].Value);

    // Parsing the raw date format of "Month Day [suffix], Year" (Column E)
    var dateField = (string)personCells[4].Value;
    validationResult.DateErrorMessage = ValidateDate(dateField);
}
```

Here is the paraphrased section from the article, with relative URL paths resolved:


## Data Validation Example

Utilize IronXL to assure the accuracy of data within a spreadsheet. The `DataValidation` sample employs the `libphonenumber-csharp` library for phone number validation and standard C# methods for email and date verification.

```cs
/**
Data Validation Example
anchor-data-validation
**/
//Loop through each row starting from the second
for (int i = 2; i <= 101; i++)
{
    var validationResult = new PersonValidationResult { Row = i };
    results.Add(validationResult);

    //Retrieve all cells for an individual entry
    var cells = worksheet [$"A{i}:E{i}"].ToList();

    //Validate phone number in column B
    var phone = cells[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phone.ToString());

    //Check email validity in column D
    validationResult.EmailErrorMessage = ValidateEmailAddress(cells[3].Value.ToString());

    //Verify date format in column E
    var date = cells[4].Value.ToString();
    validationResult.DateErrorMessage = ValidateDate(date);
}
```

This script iterates over each row to collect and validate the fields from the spreadsheet. For each row, it constructs a validation result which includes phone number, email, and date validations, logging any errors encountered.

```cs
var resultsSheet = workbook.CreateWorkSheet("Validation Results");

// Set the headers
resultsSheet ["A1"].Value = "Row";
resultsSheet ["B1"].Value = "Valid";
resultsSheet ["C1"].Value = "Phone Errors";
resultsSheet ["D1"].Value = "Email Errors";
resultsSheet ["E1"].Value = "Date Errors";

// Populate the sheet with validation results
for (int j = 0; j < results.Count; j++)
{
    var result = results[j];
    resultsSheet[$"A{j+2}"].Value = result.Row;
    resultsSheet[$"B{j+2}"].Value = result.IsValid ? "Yes" : "No";
    resultsSheet[$"C{j+2}"].Value = result.PhoneNumberErrorMessage;
    resultsSheet[$"D{j+2}"].Value = result.EmailErrorMessage;
    resultsSheet[$"E{j+2}"].Value = result.DateErrorMessage;
}

// Save the documented errors in a new file
workbook.SaveAs(@"Spreadsheets\\ValidatedData.xlsx");
```

The completion of this process yields a detailed log of validation outcomes for each data entry, making it easy to identify and resolve any issues found within the dataset.
```

This section involves reading through rows of spreadsheet data, performing validations on specific fields, and logging the results in a new worksheet.

<hr class="separator">

## 7. Validate Spreadsheet Data ##

IronXL can be used to validate the contents of a spreadsheet effectively. In the example provided, `libphonenumber-csharp` is utilized to check the validity of phone numbers, while standard C# APIs are employed to authenticate email addresses and date formats.

Here's the paraphrased section with URL and image paths resolved:

```cs
/**
Data Validation within Spreadsheets
anchor-data-validation-in-spreadsheets
**/
// Loop through each row starting from row 2 to row 101
for (var i = 2; i <= 101; i++)
{
    // Initialize validation result for the current row
    var validationResult = new PersonValidationResult { Row = i };
    results.Add(validationResult);

    // Retrieve all cell values for an individual's record
    var individualCells = worksheet [$"A{i}:E{i}"].ToList();

    // Extract the phone number from column B and validate
    var phoneNumber = individualCells [1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    // Extract the email address from column D and validate
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)individualCells [3].Value);

    // Extract the date from column E, formatted as Month Day [suffix], Year, and validate
    var dateOfBirth = (string)individualCells [4].Value;
    validationResult.DateErrorMessage = ValidateDate(dateOfBirth);
}
```

The code iterates through each row in the spreadsheet, collecting the cells into a list for processing. Each validation function evaluates the contents of a cell and generates an error message if the content does not meet the specified criteria.

Furthermore, the script establishes a new spreadsheet, designates headers, and records the error messages. This process ensures that there is a detailed record of any data inconsistencies identified during the validation process.

```cs
// Create a new worksheet named "Results" in the workbook
var resultsWorksheet = workbook.CreateWorkSheet("Results");

// Define column headers for the Results worksheet
resultsWorksheet ["A1"].Value = "Row";
resultsWorksheet ["B1"].Value = "Validity";
resultsWorksheet ["C1"].Value = "Phone Number Error";
resultsWorksheet ["D1"].Value = "Email Error";
resultsWorksheet ["E1"].Value = "Date Error";

// Loop through each entry in the results list
for (int index = 0; index < results.Count; index++)
{
    var validationResult = results[index];
    // Assign row number to column A
    resultsWorksheet [$"A{index + 2}"].Value = validationResult.Row;
    // Indicate validity in column B
    resultsWorksheet [$"B{index + 2}"].Value = validationResult.IsValid ? "Yes" : "No";
    // Record any phone number errors in column C
    resultsWorksheet [$"C{index + 2}"].Value = validationResult.PhoneNumberErrorMessage;
    // Record any email errors in column D
    resultsWorksheet [$"D{index + 2}"].Value = validationResult.EmailErrorMessage;
    // Record any date errors in column E
    resultsWorksheet [$"E{index + 2}"].Value = validationResult.DateErrorMessage;
}

// Save the workbook with the validated data to a file
workbook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
```

<hr class="separator">

## 8. Database Export with Entity Framework ##

Leverage IronXL for exporting spreadsheet content into a database, or for converting entire Excel files into structured databases. The `ExcelToDB` example demonstrates how this is achieved by processing a spreadsheet containing GDP information for various countries, subsequently sending this data to an SQLite database. This process is facilitated using Entity Framework to construct the database and systematically export the data.

Furthermore, initiate the integration by adding SQLite Entity Framework NuGet packages.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

EntityFramework enables the creation of a model object for exporting data into a database.

```cs
// Country class definition
public class Nation
{
    [Key]  // Designates 'ID' as the primary key
    public Guid ID { get; set; }  // Unique identifier for each Nation instance
    public string CountryName { get; set; }  // Name of the country
    public decimal EconomicOutput { get; set; }  // Gross Domestic Product value
}
```

This block of code sets up the database context. If you need to work with another database type, you should install the relevant NuGet package and look for the method that corresponds to `UseSqLite()`.

```cs
/**
Exporting Data with Entity Framework Integration
anchor-export-data-using-entity-framework
**/
public class CountryContext : DbContext
{
    // Set for countries table
    public DbSet<Country> Countries { get; set; }

    // Constructor ensures database creation
    public CountryContext()
    {
        // Reminder: Convert to asynchronous operation
        Database.EnsureCreated();
    }

    /// <summary>
    /// Sets up the database context to utilize Sqlite
    /// </summary>
    /// <param name="optionsBuilder"></param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        // Establish a connection to the Sqlite database
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();

        // Command to check and create the database
        var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();

        // Configure DbContext to use the Sqlite database
        optionsBuilder.UseSqlite(connection);

        // Call base configuration
        base.OnConfiguring(optionsBuilder);
    }
}
```

To manage your data effectively for an Entity Framework operation, begin by establishing a `CountryContext`. Proceed to loop through the necessary data range in order to populate each record accurately. Complete the process by invoking `SaveAsync` to ensure all changes are securely saved to your database.

```cs
public async Task ExecuteAsync()
{
    // Load the initial worksheet
    var workbook = WorkBook.Load(@"Spreadsheets\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    // Establish a database connection
    using (var dbContext = new CountryContext())
    {
        // Traverse through each row
        for (int index = 2; index <= 213; index++)
        {
            // Retrieve the range for columns A and B
            var range = worksheet[$"A{index}:B{index}"].ToList();

            // Construct a Country object to persist in the database
            var country = new Country
            {
                Name = range[0].StringValue,
                GDP = Convert.ToDecimal(range[1].DoubleValue)
            };

            // Insert the new Country object into the database
            await dbContext.Countries.AddAsync(country);
        }

        // Finalize and commit the data entries to the database
        await dbContext.SaveChangesAsync();
    }
}
```

```cs
/**
Export Data using Entity Framework
anchor-export-data-using-entity-framework
**/
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        //Ensure database is created on initialization
        Database.EnsureCreated();
    }

    /// <summary>
    /// Set up the DbContext to utilize Sqlite
    /// </summary>
    /// <param name="optionsBuilder"></param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection($"Data Source=Country.db");
        connection.Open();

        var command = connection.CreateCommand();

        //Activate foreign keys in SQLite if not already enabled
        command.CommandText = $"PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();

        optionsBuilder.UseSqlite(connection);

        base.OnConfiguring(optionsBuilder);
    }

}

//Async function to process and save Country data to the database
public async Task ProcessAsync()
{
    //Load WorkBook and get the GDPByCountry worksheet
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    //Create a new instance of the database context
    using (var countryContext = new CountryContext())
    {
        //Iterate through the worksheet data
        for (var i = 2; i <= 213; i++)
        {
            //Extract range from columns A to B
            var range = worksheet [$"A{i}:B{i}"].ToList();

            //Create a new Country object with Name and GDP values
            var country = new Country
            {
                Name = (string)range [0].Value,
                GDP = (decimal)(double)range [1].Value
            };

            //Add the new Country object to the DbSet
            await countryContext.Countries.AddAsync(country);
        }

        //Save the changes asynchronously to the database
        await countryContext.SaveChangesAsync();
    }
}
```

<hr class="separator">

## 9. Inserting Formulas into a Spreadsheet ##

Incorporate formulas into spreadsheet cells using the `Formula` property of a [`Cell`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html). This feature allows for dynamic calculations within your spreadsheet.

The following code demonstrates a practical use case by cycling through all applicable states and calculating the total percentage in column C:

```cs
// Loop through each state to calculate the percentage total
for (var y = 2; y < i; y++)
{
    // Access the cell in column C for the current row
    var cell = sheet[$"C{y}"].First();

    // Set the formula to calculate the percentage of the total
    cell.Formula = $"=B{y}/B{i}";
}
```
This script executes a loop where each cell in column C receives a formula denoting its proportional relationship to a baseline value found in the spreadsheet, effectively enabling quick and automated financial or statistical assessments.

```cs
/**
Insert Spreadsheet Calculations
anchor-insert-formulae-in-spreadsheets
**/
// Loop through every row that contains data
for (var rowIndex = 2; rowIndex < i; rowIndex++)
{
    // Retrieve the cell in column C
    var targetCell = sheet[$"C{rowIndex}"].First();

    // Assign a formula to calculate the percentage of total
    targetCell.Formula = $"=B{rowIndex}/B{i}";
}
```

```cs
/**
Add Formulas to a Spreadsheet
anchor-add-formulae-to-a-spreadsheet
**/
// Loop through all rows containing data
for (int y = 2; y < i; y++)
{
    // Access the cell in column C
    var cell = sheet[$"C{y}"].First();

    // Assign a formula to calculate the Percentage of Total for the column
    cell.Formula = $"=B{y}/B{i}";
}
```

<hr class="separator">

## 10. Download Data from an API to Spreadsheet ##

The code snippet below demonstrates how to perform a REST API call using [RestClient.Net](https://github.com/MelbourneDeveloper/RestClient.Net). This REST call retrieves JSON data, which is then parsed into a list of `RestCountry` objects. Following this, the data is seamlessly transferred and saved into an Excel file for each country listed from the API response.

```cs
// Download data from an external API into a spreadsheet
// Section ID: download-data-from-api-to-spreadsheet
var apiClient = new Client(new Uri("https://restcountries.eu/rest/v2/"));
List<RestCountry> countryList = await apiClient.GetAsync<List<RestCountry>>();
```

Here's your paraphrased section with resolved URL paths:

-----
Sample: *ApiToExcel*

Below is a representation of the data structure returned by the API in JSON format.

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="API JSON Data Example" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

The code snippet provided demonstrates iterating over a list of country data and populating an Excel spreadsheet with specific attributes for each country, such as Name, Population, Region, NumericCode, and the top three languages spoken.

Here's the paraphrased section of the code:

```cs
for (int index = 2; index < countries.Count; index++)
{
    var currentCountry = countries[index];

    // Assigning basic country details
    worksheet[$"A{index}"].Value = currentCountry.name;
    worksheet[$"B{index}"].Value = currentCountry.population;
    worksheet[$"G{index}"].Value = currentCountry.region;
    worksheet[$"H{index}"].Value = currentCountry.numericCode;

    // Processing languages for each country
    for (int langIndex = 0; langIndex < 3; langIndex++)
    {
        // Ensure there are enough languages in the list
        if (langIndex > (currentCountry.languages.Count - 1)) break;

        var language = currentCountry.languages[langIndex];

        // Calculate the column letter dynamically depending on the language index
        var column = GetColumnLetter(4 + langIndex);

        // Store the language in the corresponding cell
        worksheet[$"{column}{index}"].Value = language.name;
    }
}
```

<hr class="separator">

<h2>API Reference and Resources</h2>

You might find the [IronXL class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/) within the API Reference extremely useful.

Furthermore, additional tutorials are available that explore various functionalities of IronXL.Excel, such as [Creating](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/), [Opening, Writing, Editing, Saving, and Exporting](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/) XLS, XLSX, and CSV files without the need for Excel Interop.

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img alt="" class="img-responsive add-shadow" src="/img/svgs/brand-visual-studio.svg">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>Download this Tutorial as C# Source Code</h3>
      <p>The full free C# for Excel Source Code for this tutorial is available to download as a zipped Visual Studio 2017 project file.</p>
      <a class="btn btn-white3" href="/csharp/excel/tutorials/downloads/How.to.Read.an.Excel.File.in.CSharp.zip">
        <i class="fa fa-cloud-download"></i> Download</a>
      </div>
  </div>
</div>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore this Tutorial on GitHub</h3>
      <p>The source code for this project is available in C# and VB.NET on GitHub.</p>
      <p>Use this code as an easy way to get up and running in just a few minutes. The project is saved as a Microsoft Visual Studio 2017 project, but is compatible with any .NET IDE.</p>
      <a class="doc-link" href="https://github.com/iron-software/tutorials/tree/master/IronXL/How%20to%20Read%20an%20Excel%20File%20in%20C%23" target="_blank">How to Read Excel File in C# on GitHub<i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img alt="" class="img-responsive add-shadow" src="/img/svgs/github-icon.svg">
      </div>
    </div>
  </div>
</div>



<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>View the API Reference</h3>
      <p>Explore the API Reference for IronXL, outlining the details of all of IronXL’s features, namespaces, classes, methods fields and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

