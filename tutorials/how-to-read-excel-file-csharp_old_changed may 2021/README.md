# C# Excel File Handling with Practical Examples

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp_old_changed may 2021/>***


This guide demonstrates how to extract data from Excel files using C# and leverage the library for common operations such as data validation, converting data for database storage, capturing data from Web APIs, and altering formulas within the spreadsheet. It highlights examples from the IronXL library coded within a .NET Core Console Application.

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

IronXL is a comprehensive .NET library designed for manipulating and managing Microsoft Excel files using C#. This guide provides step-by-step instructions on how to utilize C# to [read Excel files](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/).

1. Begin by installing the IronXL Excel library. This can be achieved through the [NuGet package](https://www.nuget.org/packages/IronXL.Excel/) directly, or by manually downloading the [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

2. To open an Excel file such as an XLS, XLSX, or CSV document, utilize the `WorkBook.Load` method.

3. Retrieve cell values effectively using straightforward syntax, for example: `sheet["A11"].DecimalValue`.

<h3>IronXL Includes:</h3>

- Receive tailored assistance directly from our dedicated .NET engineering team.
  
- Simplify setup with straightforward integration into Microsoft Visual Studio.

- Utilize the development version at no cost. Affordable licensing options start from `$liteLicense`.


Experience the simplicity of processing **Excel** files using C# or VB.NET with the IronXL library, inclusive of examples across three different Excel spreadsheets.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<h4>Read XLS or XLSX Files: Quick Code</h4>

In this demonstration, it's evident that reading *Excel* files in C# can be achieved seamlessly without the need for Interop. Additionally, the concluding section on Advanced Operations highlights the **Linq** compatibility and the capacity to perform aggregate computations over a range of data.

```cs
/**
Reading and processing Excel files
anchor-processing-excel-files
**/
using IronXL;
using System.Linq;

// The following Excel file formats are supported: XLSX, XLS, CSV, and TSV
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet sheet = workbook.WorkSheets.First();
// Simply select cells using traditional Excel-style references and obtain their values
int singleCellValue = sheet["A2"].IntValue;
// Elegantly read ranges of cells and retrieve their details.
foreach (var cell in sheet["A2:A10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}

///More Complex Operations

// Perform calculations on ranges, such as minimum, maximum, and sum
decimal totalSum = sheet["A2:A10"].Sum();
// This is also compatible with Linq queries
decimal maximumValue = sheet["A2:A10"].Max(c => c.DecimalValue);
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Obtain the IronXL C# Library at No Cost

To begin, it is essential to incorporate the IronXL.Excel library into your project to enhance your .NET framework with Excel capabilities.

<h3>Installing the IronXL NuGet Package</h3>

Here's the paraphrased section with all relative URL paths resolved to ironsoftware.com:


1. Open Visual Studio, right-click your project in the solution explorer, and choose "Manage NuGet Packages..."

2. Look up the IronXL.Excel package in the search bar and proceed with the installation.
```

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
  <p><img src="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<br>

Here is an alternate method for installation:

1. Open the Package Manager Console in your development environment.
2. Execute the command: `Install-Package IronXL.Excel`

```shell
Install-Package IronXL.Excel
```

<br>

Furthermore, you can also [view the package on the NuGet site](https://www.nuget.org/packages/IronXL.Excel/) for more details.

<h3>Direct Download Installation</h3>

Alternatively, you can begin by downloading the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and proceed with a manual installation into Visual Studio.

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Load a Workbook ##

The [`WorkBook`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) class embodies an Excel spreadsheet document. To initiate a `WorkBook`, employ the `WorkBook.Load` method and designate the file path to the Excel file (.xlsx).

Here's the paraphrased section with resolved URL paths:

```cs
/**
Load an Excel Workbook
anchor-load-excel-workbook
**/
var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

In IronXL, a `WorkBook` instance can consist of several <a href="https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkSheet.html" target="_blank">`WorkSheet`</a> objects, each representing a single worksheet from an Excel document. If your Excel file includes multiple worksheets, you can access a specific worksheet by name using the <a href="https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html" target="_blank">`WorkBook.GetWorkSheet`</a> method.

```cs
// Retrieve the worksheet named "GDPByCountry" from the workbook
var worksheet = workbook.GetWorkSheet("GDPByCountry");
```

In this section titled *ExcelToDB*, we delve into how IronXL can be utilized to migrate data from an Excel spreadsheet to a SQL database using the Entity Framework. This example illustrates the process of reading spreadsheet data and exporting it to an SQLite database.

Initially, we install necessary SQLite Entity Framework packages.

[Click here to see the additional details!](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)

![Entity Framework Integration](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)

Entity Framework facilitates the creation of model objects for the database. Here, the model for a `Country` class is defined with properties for `Key`, `Name`, and `GDP`.

```cs
/**
Use Entity Framework for exporting
anchor-export-data-using-entity-framework
**/
public class Country
{
    [Key]
    public Guid Key { get; set; }
    public string Name { get; set; }
    public decimal GDP { get; set; }
}
```

Next, the database context configuration follows. Replace the `UseSqlite` method with the relevant method based on the database being used if you are not using SQLite.

```cs
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        Database.EnsureCreated();
    }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();
        var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();
        optionsBuilder.UseSqlite(connection);
        base.OnConfiguring(optionsBuilder);
    }
}
```

Following the setup, we proceed with the actual data processing. Here's how we load the workbook, access the specific worksheet, and iterate through the records to populate the database asynchronously.

```cs
public async Task ProcessAsync()
{
    var workbook = WorkBook.Load("Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    using (var countryContext = new CountryContext())
    {
        for (int i = 2; i <= 213; i++)
        {
            var range = worksheet [$"A{i}:B{i}"].ToList();
            var country = new Country
            {
                Name = range[0].StringValue,
                GDP = range[1].DecimalValue
            };

            countryContext.Countries.Add(country);
        }
        await countryContext.SaveChangesAsync();
    }
}
```

This example notably demonstrates how to convert data from an Excel file into a usable database format, leveraging IronXL with Entity Framework for streamlined data manipulation and storage.

<hr class="separator">

## 3. Initialize a WorkBook ##

To instantiate a new `WorkBook` in memory, simply create a new instance while specifying the type of spreadsheet it should handle.

```cs
/**
Initialize a New Workbook
anchor-initialize-new-workbook
**/
var workbook = new WorkBook(ExcelFileFormat.XLSX);
```

```cs
/**
Create WorkBook
anchor-create-a-workbook
**/
var workbook = new WorkBook(ExcelFileFormat.XLSX);
```
Sample: *ApiToExcelProcessor*

For handling older Excel documents (Excel 95 or earlier), use `ExcelFileFormat.XLS`.

<hr class="separator">

## 4. Create a WorkSheet ##

A "WorkBook" in IronXL can contain numerous "WorkSheets," which are essentially individual data sheets. Conversely, the "WorkBook" serves as a container for these numerous "WorkSheets." Below is a visual representation of a workbook containing two distinct worksheets in Excel:

<center>
  <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="" class="img-responsive add-shadow img-margin" style="width:100%; max-width:100%;">
  </a>
</center>

<center>
  <a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
  </a>
</center>

To initiate a new WorkSheet, use the `WorkBook.CreateWorkSheet` method and provide the desired name for the WorkSheet.

```cs
// Creating a new worksheet named "Countries" within the workbook
var worksheet = workbook.CreateWorkSheet("Countries");
```

<hr class="separator">

## 5. Retrieve Cell Ranges ##

The `Range` class embodies a two-dimensional array of `Cell` objects, symbolizing a specified range of cells within an Excel spreadsheet. You can acquire these ranges by employing the string indexer on a `WorkSheet` instance. The input can either be a single cell's coordinates, like "A1", or a continuous block of cells, such as "B2:E5". Additionally, the `GetRange` method available on a `WorkSheet` allows for similar retrieval of cell ranges.

```cs
// Access a specific range of cells from D2 to D101 within the worksheet
var range = worksheet["D2:D101"];
```

## Data Validation Sample

In the `DataValidation` sample within the IronXL tutorial, we utilize the proficient capabilities of IronXL to validate a spreadsheet's data comprehensively. This process utilizes the `libphonenumber-csharp` library for phone number validation and employs standard C# methods to check the validity of email addresses and dates.

```csharp
/**
Data Validation Example
anchor-data-validation-example
**/
// Loop through each rowstarting from row 2
for (var index = 2; index <= 101; index++)
{
    // Instantiate a validation result object for each row
    var validationResult = new PersonValidationResult { Row = index };

    // Store the validation result
    results.Add(validationResult);

    // Retrieve all cells for an individual entry
    var individualCells = worksheet [$"A{index}:E{index}"].ToList();

    // Validate the phone number in column B
    var phone = individualCells [1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Validate the email address in column D
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)individualCells [3].Value);

    // Format and validate the date in column E
    var formatDate = (string)individualCells [4].Value;
    validationResult.DateErrorMessage = ValidateDate(formatDate);
}
```

The script systematically validates each row, retrieving cells from `A` to `E` and applying validation criteria to relevant fields like phone number, email, and date. The `PersonValidationResult` provides a structured format to log and report any discrepancies found during validation.

The results from this validation are then logged in a new worksheet termed "Results" for audit and correction purposes:

```csharp
var resultsSheet = workbook.CreateWorkSheet("Results");

// Set headers
resultsSheet ["A1"].Value = "Row";
resultsSheet ["B1"].Value = "Valid";
resultsSheet ["C1"].Value = "Phone Error";
resultsSheet ["D1"].Value = "Email Error";
resultsSheet ["E1"].Value = "Date Error";

// Populate results sheet from validation
for (var i = 0; i < results.Count; i++)
{
    var result = results [i];
    resultsSheet [$"A{i + 2}"].Value = result.Row;
    resultsSheet [$"B{i + 2}"].Value = result.IsValid ? "Yes" : "No";
    resultsSheet [$"C{i + 2}"].Value = result.PhoneNumberErrorMessage;
    resultsSheet [$"D{i + 2}"].Value = result.EmailErrorMessage;
    resultsSheet [$"E{i + 2}"].Value = result.DateErrorMessage;
}

// Save the workbook to a file
workbook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
```

This method not only validates data but also helps in preemptively correcting errors, ensuring data integrity before further processing or analysis.

<hr class="separator">

## 6. Modifying Cell Values Across a Range ##

There are multiple approaches to either reading or updating the contents of cells within a specified Range. If the number of cells is predetermined, utilizing a For loop is an effective method.

Here's a paraphrased version of the provided code section, where relative paths have also been resolved to the specified domain (ironsoftware.com):

```cs
/**
Modify Range Cell Content
anchor-modify-cell-content-in-range
**/
// Loop through the designated rows
for (int rowIndex = 2; rowIndex <= 101; rowIndex++)
{
    var validationResult = new PersonValidationResult { Row = rowIndex };
    results.Add(validationResult);

    // Retrieve all cells related to a single person
    var personCells = worksheet [$"A{rowIndex}:E{rowIndex}"].ToList();

    // Perform phone number validation (column B signifies index 1)
    var phone = personCells[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phone.ToString());

    // Perform email validation (column D signifies index 3)
    validationResult.EmailErrorMessage = ValidateEmailAddress(personCells[3].Value.ToString());

    // Fetch and verify the date which is structured as Month Day [suffix], Year (from column E, index 4)
    var dateEntry = personCells[4].Value.ToString();
    validationResult.DateErrorMessage = ValidateDate(dateEntry);
}
```

The provided section titled *"DataValidation"* from the article demonstrates the use of IronXL for validating spreadsheet data, specifically focusing on phone numbers, email addresses, and date formats.

```cs
/**
Validate Spreadsheet Data
anchor-validate-spreadsheet-data
**/
//Loop through each row, beginning from the second
for (var i = 2; i <= 101; i++)
{
    var result = new PersonValidationResult { Row = i };
    results.Add(result);

    //Collect all relevant cells for the current person
    var cells = worksheet [$"A{i}:E{i}"].ToList();

    //Validation of phone number at position B (index 1)
    var phoneNumber = cells [1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    //Email validation at position D (index 3)
    result.EmailErrorMessage = ValidateEmailAddress((string)cells [3].Value);

    //Date parsing and validation at position E (index 4)
    var rawDate = (string)cells [4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

This snippet efficiently iterates through each row of a spreadsheet, collecting and validating data related to personal information such as phone numbers, email, and dates using custom validation functions. The results, which include error messages for any invalid entries, are stored and used to further handle data quality assurance within the application.

<hr class="separator">

## 7. Validate Spreadsheet Data ##

IronXL provides robust tools for data validation within spreadsheets. The `DataValidation` example employs the `libphonenumber-csharp` library for phone number validation and standard C# APIs to validate both email addresses and dates.

```cs
/**
Data Validation in Spreadsheets
anchor-validate-spreadsheet-data
**/
// Loop through each row
for (var index = 2; index <= 101; index++)
{
    var validationResult = new PersonValidationResult { Row = index };
    results.Add(validationResult);

    // Retrieve all cell values for an individual
    var cellValues = worksheet[$"A{index}:E{index}"].ToList();

    // Verification of the phone number (1 = B column)
    var phone = cellValues[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Email validation (3 = D column)
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)cellValues[3].Value);

    // Parse the raw date format of Month Day [suffix], Year (4 = E column)
    var rawBirthDate = (string)cellValues[4].Value;
    validationResult.DateErrorMessage = ValidateDate(rawBirthDate);
}
```

The code snippet presented iterates over each row in the spreadsheet, retrieving the cells as a list. Each validation method assesses the cell value and produces an error message if the value does not meet specified criteria.

Furthermore, the code initializes a new spreadsheet, defines headers, and records the error messages. This process ensures that there is a detailed record of any data discrepancies.

```cs
// Initialize the Results worksheet in the workbook
var resultsWorksheet = workbook.CreateWorkSheet("Results");

// Define headers for the Results worksheet
resultsWorksheet["A1"].Value = "Row";
resultsWorksheet["B1"].Value = "Valid";
resultsWorksheet["C1"].Value = "Phone Error";
resultsWorksheet["D1"].Value = "Email Error";
resultsWorksheet["E1"].Value = "Date Error";

// Populate the worksheet with data for each validation result
for (var index = 0; index < results.Count; index++)
{
    var validationResult = results[index];
    // Assigning values to respective columns starting from row 2
    resultsWorksheet[$"A{index + 2}"].Value = validationResult.Row;
    resultsWorksheet[$"B{index + 2}"].Value = validationResult.IsValid ? "Yes" : "No";
    resultsWorksheet[$"C{index + 2}"].Value = validationResult.PhoneNumberErrorMessage;
    resultsWorksheet[$"D{index + 2}"].Value = validationResult.EmailErrorMessage;
    resultsWorksheet[$"E{index + 2}"].Value = validationResult.DateErrorMessage;
}

// Save the workbook with validation results to a file
workbook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
```

<hr class="separator">

## 8. Data Exportation using Entity Framework ##

With IronXL, you can smoothly transfer data from an Excel spreadsheet into a database format. The example `ExcelToDB` demonstrates how to take a spreadsheet that lists countries by their GDP and move that information into an SQLite database using the EntityFramework for both structuring the database and managing the data line-by-line exportation process.

To get started, first ensure to install the necessary SQLite Entity Framework NuGet packages.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

EntityFramework provides the capability to define a model object which can be used for data export to a database.

```cs
public class Nation
{
    [Key]
    public Guid ID { get; set; }
    public string CountryName { get; set; }
    public decimal GrossDomesticProduct { get; set; }
}
```

The provided code snippet is set up to configure the database context. If you need to utilize another database, you should install the suitable NuGet package and locate the appropriate method that corresponds to `UseSqLite()`.

Here's the paraphrased section of the original text, with relative URL paths resolved to `ironsoftware.com` as requested:

```cs
/**
 * Initialization and configuration of the Entity Framework for data export
 * Anchor: export-data-using-entity-framework
 **/
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }  // Set of 'Country' entities

    public CountryContext()
    {
        //TODO: Update to asynchronous operations
        Database.EnsureCreated();  // Ensure the database is created if it doesn't already exist
    }

    /// <summary>
    /// Set configurations for using SQLite
    /// </summary>
    /// <param name="optionsBuilder">The options to modify the database context</param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();  // Open the connection

        var command = connection.CreateCommand();

        // Ensure SQL foreign keys are enabled
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();  // Execute the creation command

        optionsBuilder.UseSqlite(connection);  // Use SQLite with the current connection

        base.OnConfiguring(optionsBuilder);  // Continue with base configuration
    }
}
```

Establish a `CountryContext`, process each record in the provided range, and utilize `SaveAsync` to finalize the updates to the database.

Below is the paraphrased version of the provided C# code snippet, which is designed to automate the process of importing data from an Excel spreadsheet into a database using IronXL and Entity Framework.

```cs
public async Task ImportDataAsync()
{
    // Retrieve the primary worksheet from the file
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    // Establish a new database connection context
    using (var dbContext = new CountryContext())
    {
        // Loop through the designated cells in the worksheet
        for (var index = 2; index <= 213; index++)
        {
            // Extract the data in the range covering columns A and B
            var cellData = worksheet[$"A{index}:B{index}"].ToList();

            // Instantiate a new Country object populated with the data from cells
            var country = new Country
            {
                Name = cellData[0].StringValue,
                GDP = Convert.ToDecimal(cellData[1].DoubleValue)
            };

            // Append the new Country to the database context
            await dbContext.Countries.AddAsync(country);
        }

        // Execute the data transaction to the database
        await dbContext.SaveChangesAsync();
    }
}
```

This revision maintains the original logic for processing and storing each `Country` object but refines the variable naming and comments to enhance readability.

## Exporting Data via Entity Framework Using IronXL

Utilize IronXL to export Excel data into a database or transform an Excel sheet into a database structure. The example involving GDP data from various countries illustrates how an Excel spreadsheet is read and exported to an SQLite database using Entity Framework.

Firstly, ensure the required SQLite Entity Framework NuGet packages are configured in your environment. 

<a href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

Entity Framework supports creation of model objects that facilitate data export into the database.

```cs
public class Country
{
    [Key]
    public Guid Key { get; set; }
    public string Name { get; set; }
    public decimal GDP { get; set; }
}
```

Below is the code snippet to configure the database context. Adapt this snippet by installing the appropriate NuGet package and apply the equivalent to `UseSqlite()` for different databases.

```cs
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        // Ensure database existence synchronously
        Database.EnsureCreated();
    }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();

        var command = connection.CreateCommand();
        
        // Enforce foreign key constraints
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();

        optionsBuilder.UseSqlite(connection);
        base.OnConfiguring(optionsBuilder);
    }
}
```

For the process of exporting data, create a `CountryContext`, iterate through each record in the range, and asynchronously save them to the database.

```cs
public async Task ProcessAsync()
{
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    // Establishing connection with the database
    using (var countryContext = new CountryContext())
    {
        for (var i = 2; i <= 213; i++)
        {
            var range = worksheet[$"A{i}:B{i}"].ToList();
            
            // Creating the Country object to be stored
            var country = new Country
            {
                Name = (string)range[0].Value,
                GDP = (decimal)(double)range[1].Value
            };

            // Adding to context
            await countryContext.Countries.AddAsync(country);
        }

        // Committing changes
        await countryContext.SaveChangesAsync();
    }
}
```

This sample, *ExcelToDB*, exemplifies the ease of integrating IronXL with Entity Framework for streamline data transformations and exports from Excel to databases.

<hr class="separator">

## 9. Input Formulas into Spreadsheet Cells ##

Assign formulas to each cell using the `Cell`'s [`Formula`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) attribute. 

```cs
/**
Add Formulas to Each Row
anchor-add-formulae-to-a-spreadsheet
**/
// Traverse through the rows to apply formula to each cell in column C
for (var rowIndex = 2; rowIndex < totalRows; rowIndex++)
{
    // Access the cell in column C for each state
    var currentStateCell = sheet[$"C{rowIndex}"].First();

    // Apply a formula to calculate the percentage of total
    currentStateCell.Formula = $"=B{rowIndex}/B{totalRows}";
}
```
The script processes each row, setting the cell's formula in column C to compute the percentage contribution to the total, showing the adaptability of the `IronXL` tool for dynamic Excel manipulations.

```cs
/**
Update Spreadsheet Calculations
anchor-update-spreadsheet-calculations
**/
// Loop through rows with existing values
for (var rowIndex = 2; rowIndex < i; rowIndex++)
{
    // Access the cell in column C 
    var targetCell = sheet[$"C{rowIndex}"].First();

    // Assign a formula to calculate the column's percentage total
    targetCell.Formula = $"=B{rowIndex}/B{i}";
}
```

```cs
/**
Add Spreadsheet Formulae
anchor-add-formulae-to-a-spreadsheet
**/
// Loop through rows containing data
for (var y = 2; y < i; y++)
{
    // Access the cell in column C
    var cell = sheet[$"C{y}"].First();

    // Assign a formula to calculate the percentage of the total
    cell.Formula = $"=B{y}/B{i}";
}
```
This segment of code iterates through each populated row, selecting the cell in column C, and sets a formula that calculates and displays the percentage relative to a total value in cell B. Each cell's formula is dynamically set based on its row position to ensure the correct calculation.

<hr class="separator">

## 10. Download Data from an API to Spreadsheet ##

The subsequent example demonstrates how to perform a REST API request using [RestClient.Net](https://github.com/MelbourneDeveloper/RestClient.Net). This process retrieves JSON data and transforms it into a `List` of the `RestCountry` type. From there, you can efficiently loop through each country and transfer the API data directly into an Excel file.

Here's your paraphrased section with resolved URLs:

```cs
/**
Retrieve Data from API and Load into Spreadsheet
anchor-download-data-from-an-api-to-spreadsheet
**/
var httpClient = new Client(new Uri("https://restcountries.eu/rest/v2/"));
List<RestCountry> countryList = await httpClient.GetAsync<List<RestCountry>>();
```

Below is the paraphrased section of the article, with links and image paths resolved to ironsoftware.com:

---
Sample: *ApiToExcel*

Here is a preview of the JSON data obtained from the API:

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

The subsequent code segment cycles through the list of countries and populates columns in the spreadsheet with the Name, Population, Region, NumericCode, and the Top 3 Languages of each country.

Here is the paraphrased version of the provided C# code snippet:

```cs
// Loop through countries starting from the second item
for (var index = 2; index < countries.Count; index++)
{
    var country = countries[index];

    // Assign primary attributes to corresponding worksheet cells
    worksheet[$"A{index}"].Value = country.name;
    worksheet[$"B{index}"].Value = country.population;
    worksheet[$"G{index}"].Value = country.region;
    worksheet[$"H{index}"].Value = country.numericCode;

    // Process the top three languages per country
    for (var langIndex = 0; langIndex < 3; langIndex++)
    {
        // Exit the loop if there aren't enough languages
        if (langIndex > country.languages.Count - 1) break;

        var language = country.languages[langIndex];

        // Determine the corresponding spreadsheet column for this language
        var columnLetter = GetColumnLetter(4 + langIndex);

        // Place the language's name in the correct cell
        worksheet[$"{columnLetter}{index}"].Value = language.name;
    }
}
```

<hr class="separator">

<h2>API Reference and Resources</h2>

You will likely find the [IronXL class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/) within the API Reference to be extremely useful.

Furthermore, you can explore additional tutorials that illuminate various features of IronXL.Excel such as [Creating](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/), [Opening, Writing, Editing, Saving, and Exporting](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/) XLS, XLSX, and CSV files without the need for Excel Interop.

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

