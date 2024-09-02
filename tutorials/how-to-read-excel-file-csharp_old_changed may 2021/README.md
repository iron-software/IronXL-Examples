# C# Excel Reading with Practical Examples

This guide provides instructions on extracting data from Excel files using C# and demonstrates how to utilize the IronXL library for common functionalities such as data validation, database conversions, saving information from Web APIs, and altering formulas within your Excel documents. This document is associated with IronXL example codes, developed as a .NET Core Console application.

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

IronXL is a comprehensive .NET library designed for manipulating and managing Microsoft Excel files using C#. This guide provides detailed instructions on how to utilize C# to efficiently handle Excel documents.

- **Step one**: Begin by installing the IronXL Excel Library. You have the option to integrate it via our [NuGet package](https://www.nuget.org/packages/IronXL.Excel/), accessible directly or by downloading our [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

- **Step two**: Utilize the `WorkBook.Load` method, which enables you to open and read various spreadsheet formats including XLS, XLSX, and CSV.

- **Step three**: Retrieve and manipulate cell data using a clear and concise syntax such as `sheet["A11"].DecimalValue` to fetch the decimal value of a specific cell.

This tutorial will guide you through reading Excel files using a series of easy-to-follow steps, facilitated by helpful links and examples.

<h3>IronXL Includes:</h3>

- Access to committed product assistance from our team of .NET experts

- Straightforward setup through Microsoft Visual Studio

- No cost for development environments. Licensing starts from `$liteLicense`.


Discover the simplicity of reading **Excel** files using C# or VB.NET with the IronXL library. Included are samples featuring three different Excel spreadsheets.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<h4>Read XLS or XLSX Files: Quick Code</h4>

In this illustration, it's evident that *Excel* documents can be efficiently parsed without relying on Interop in C#. The subsequent Advanced Operations demonstrate compatibility with **Linq** and the ability to perform aggregate calculations over ranges.

```cs
/**
Load and Process Excel Spreadsheet
anchor-load-and-process-excel-spreadsheet
**/
using IronXL;
using System.Linq;

// The library supports reading from various formats like XLSX, XLS, CSV, and TSV
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet sheet = workbook.DefaultWorkSheet; // Get the first worksheet
// Easily access cells using Excel-style notations and retrieve values
int valueInCellA2 = sheet["A2"].IntValue;
// Elegant reading from a cell range
foreach (var cell in sheet["A2:A10"])
{
    Console.WriteLine("Cell {0} contains: '{1}'", cell.AddressString, cell.Text);
}

///Additional Operations

// Perform calculations on range values such as Minimum, Maximum, and Sum
decimal totalSum = sheet["A2:A10"].Sum();
// Utilize LINQ to perform calculations on cell range
decimal maxValue = sheet["A2:A10"].Max(cell => cell.DecimalValue);
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Free IronXL C# Library Download

To begin, it's essential to install the IronXL.Excel library, which incorporates Excel capabilities into the .NET framework. This integration allows for enhanced data management and analysis within .NET applications.

<h3>Installing the IronXL NuGet Package</h3>

1. Within Visual Studio, perform a right-click on your project and choose the "Manage NuGet Packages ..." option.

2. Look up the IronXL.Excel package in the search bar and proceed with the installation.
```

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
  <p><img src="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<br>

Here's a paraphrased version of the selected section:

-----

An alternative installation method involves:

1. Opening the Package Manager Console
2. Executing the command: `Install-Package IronXL.Excel`

Here's the paraphrased section with links and images resolved to `ironsoftware.com`:

```shell
Install-Package IronXL.Excel
```

<br>

Furthermore, you have the option to [explore the package on NuGet](https://www.nuget.org/packages/IronXL.Excel/) by following this link.

<h3>Direct Download Installation</h3>

Alternatively, you can begin by obtaining the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually integrating it into Visual Studio.

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Opening a Spreadsheet ##

The [`WorkBook`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) class embodies an Excel file in IronXL. For opening a WorkBook, employ the `WorkBook.Load` method and provide the file path for the Excel document (.xlsx).

Here's the paraphrased section of the code with the relative URL path resolved:

```cs
/**
Load an Excel File
anchor-load-excel-file
**/
var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

Sample: *ExcelToDBProcessor*

Every `WorkBook` instance can support numerous <a href="https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkSheet.html" target="_blank">`WorkSheet`</a> instances, each signifying a distinct sheet within an Excel file. To access any specific worksheet within the document, you can utilize the method <a href="https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html" target="_blank">`WorkBook.GetWorkSheet`</a> by specifying the sheet's name.

```cs
var sheet = workbook.GetWorkSheet("GDPByCountry");
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
        //Ensure the database is created if it does not exist
        Database.EnsureCreated();
    }

    /// Configure the context to utilize SQLite
    /// <param name="optionsBuilder">The options builder for DbContext configuration</param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        // Establish the connection string and open the connection
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();

        // Ensure foreign keys are activated in SQLite
        var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();

        // Configure to use SQLite
        optionsBuilder.UseSqlite(connection);

        base.OnConfiguring(optionsBuilder);
    }
}

/// <summary>
/// Process and save each country into the database asynchronously
/// </summary>
public async Task ProcessAsync()
{
    // Load the workbook and select the specific sheet
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    // Establish a new database context
    using (var countryContext = new CountryContext())
    {
        // Go through each relevant cell in the worksheet
        for (var i = 2; i <= 213; i++)
        {
            // Fetch columns A and B that contain name and GDP
            var range = worksheet[$"A{i}:B{i}"].ToList();

            // Populate the Country object with values
            var country = new Country
            {
                Name = range[0].StringValue,
                GDP = range[1].DecimalValue
            };

            // Add the country object to the database
            await countryContext.Countries.AddAsync(country);
        }

        // Commit the transactions to save changes to the database
        await countryContext.SaveChangesAsync();
    }
}
```
This revised code snippet, part of the *ExcelToDB* example, demonstrates how to export data using Entity Framework, with enhanced comments for clarify and improved scripting standards.

<hr class="separator">

## 3. Constructing a New WorkBook ##

Start by initializing a new WorkBook in memory by specifying the desired spreadsheet format.

Here's the paraphrased section of the article with resolved URLs:

-----
```cs
// Initialize a new WorkBook instance specifying the Excel file format
var workbook = new WorkBook(ExcelFileFormat.XLSX);
```
-----

```cs
// Note for legacy Excel versions (Excel 95 or earlier), you should use the XLS file format.
var workbook = new WorkBook(ExcelFileFormat.XLS);
```

<hr class="separator">

## 4. Constructing a WorkSheet ##

Every `WorkBook` may incorporate multiple `WorkSheets`. Essentially, `WorkSheets` serve as individual sheets containing your data, whereas a `WorkBook` encompasses a compilation of these `WorkSheets`. Here's an illustration of a workbook with two distinct worksheets as viewed in Excel:

<center>
  <a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
  </a>
</center>

To set up a new WorkSheet, utilize the `WorkBook.CreateWorkSheet` method and provide the worksheet's name as the argument.

Here's the paraphrased section:

```cs
var worksheet = workbook.AddWorkSheet("Countries");
```

<hr class="separator">

## 5. Retrieve a Range of Cells ##

The `Range` class encapsulates a two-dimensional array of `Cell` objects, effectively covering a specific area of cells within an Excel spreadsheet. You can acquire ranges by applying the string indexer to a `WorkSheet` instance. The indexer accepts either a singular cell coordinate (for instance, "A1") or a range that spans from one corner to an opposite corner of cells (such as "B2:E5"). Additionally, the `GetRange` method can be employed on a `WorkSheet` to achieve similar outcomes.

```cs
// Selecting a cell range in an Excel worksheet spanning from D2 to D101
var selectedRange = worksheet["D2:D101"];
```

```cs
/**
Data Validation Example
anchor-validate-data
**/
//Loop through each row
for (var i = 2; i <= 101; i++)
{
    var validationResult = new PersonValidationResult { Row = i };
    results.Add(validationResult);

    //Retrieve all cells for this individual
    var cells = worksheet [$"A{i}:E{i}"].ToList();

    //Check the validity of the phone number (column B = 1)
    var phoneNumber = cells [1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    //Verify the email address (column D = 3)
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)cells [3].Value);

    //Extract the raw date in 'Month Day [suffix], Year' format (column E = 4)
    var rawDate = (string)cells [4].Value;
    validationResult.DateErrorMessage = ValidateDate(rawDate);
}

//This script reviews each row in an Excel spreadsheet and evaluates the information in specific cells. Validation methods return error messages if the data doesn’t conform to expected formats.

```cs
var errorLogSheet = workbook.CreateWorkSheet("Error Log");

errorLogSheet ["A1"].Value = "Row";
errorLogSheet ["B1"].Value = "Validation Status";
errorLogSheet ["C1"].Value = "Phone Number Errors";
errorLogSheet ["D1"].Value = "Email Errors";
errorLogSheet ["E1"].Value = "Date Errors";

for (var index = 0; index < results.Count; index++)
{
    var result = results [index];
    errorLogSheet [$"A{index + 2}"].Value = result.Row;
    errorLogSheet [$"B{index + 2}"].Value = result.IsValid ? "Yes" : "No";
    errorLogSheet [$"C{index + 2}"].Value = result.PhoneNumberErrorMessage;
    errorLogSheet [$"D{index + 2}"].Value = result.EmailErrorMessage;
    errorLogSheet [$"E{index + 2}"].Value = result.DateErrorMessage;
}

//Saving the workbook to a file
workbook.SaveAs("Spreadsheets\\ValidatedData.xlsx");
```

<hr class="separator">

## 6. Modify Values Within a Cell Range ##

To access or modify cell values in a specified range, multiple approaches can be utilized. If you know the total number of cells you want to modify, implementing a 'For' loop is an effective method.

```cs
/**
Modify Cell Values Across a Range
anchor-modify-cell-values-in-range
**/
// Loop through each row
for (int y = 2; y <= 101; y++)
{
    // Create a validation result for each row
    var validationResult = new PersonValidationResult { Row = y };
    results.Add(validationResult);

    // Retrieve all cell values for a single row
    var cellGroup = worksheet [$"A{y}:E{y}"].ToList();

    // Validate the phone number located in column B
    var phone = cellGroup [1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phone.ToString());

    // Validate the email address found in column D
    var email = cellGroup [3].Value;
    validationResult.EmailErrorMessage = ValidateEmailAddress(email.ToString());

    // Extract the date from column E in the specified format: Month Day [suffix], Year
    var dateInfo = cellGroup [4].Value;
    validationResult.DateErrorMessage = ValidateDate(dateInfo.ToString());
}
```

Here's the updated section that applies your requested modifications:

-----
```cs
/**
Validate Spreadsheet Data
anchor-validate-spreadsheet-data
**/
//Loop through spreadsheet rows starting from the second one
for (var i = 2; i <= 101; i++)
{
    //Create a result object for the current row
    var result = new PersonValidationResult { Row = i };
    results.Add(result);

    //Retrieve all the cells for this particular row
    var cells = worksheet [$"A{i}:E{i}"].ToList();

    //Examine and authenticate the phone number in column B
    var phoneNumber = cells [1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    //Confirm the email address in column D
    result.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);

    //Fetch and verify date from column E
    var rawDate = (string)cells[4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

This segment iteratively processes each row within a specified range on an Excel worksheet, executing several validation checks. It fetches the entire row of cells, ensuring both phone numbers and email addresses adhere to appropriate formats, and verifies date validity as per specific requirements. Each check either passes and moves to the next or flags an error for further action.

<hr class="separator">

## 7. Validate Spreadsheet Data ##

Leverage IronXL to authenticate the information contained within a spreadsheet. The 'DataValidation' sample incorporates `libphonenumber-csharp` for phone number verification and employs common C# APIs to confirm the integrity of email addresses and date entries.

```cs
/**
 * We validate data in an Excel spreadsheet, focusing on phone numbers, email addresses, and dates.
 * Reference: anchor-validate-spreadsheet-data
**/

// Looping through the designated rows for data validation
for (int i = 2; i <= 101; i++)
{
    var validationResult = new PersonValidationResult { Row = i };
    results.Add(validationResult);

    // Retrieve all cell data for a given row
    var cellData = worksheet[$"A{i}:E{i}"].ToList();

    // Phone number validation is carried out with the second cell (index 1 which is column B)
    var phoneNumber = cellData[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phoneNumber.ToString());

    // Email validation happens with the fourth cell (index 3 which is column D)
    validationResult.EmailErrorMessage = ValidateEmailAddress(cellData[3].Value.ToString());

    // Date validation uses the fifth cell (index 4 which is column E) interpreted as a date string
    var dateEntry = cellData[4].Value.ToString();
    validationResult.DateErrorMessage = ValidateDate(dateEntry);
}
```

The provided code iterates over each row in the Excel sheet, collecting the cells into a list. It employs various validation methods to inspect the content of each cell, issuing an error message when an invalid value is discovered.

Moreover, this script initializes a new spreadsheet, defines the headers, and records any error messages, thereby creating a comprehensive record of any discrepancies found in the data.

Below is the paraphrased section of the article with resolved URL paths for images and links to ironsoftware.com:

```cs
// Initialize a new worksheet to hold validation results
var validationSheet = workbook.CreateWorkSheet("ValidationResults");

// Define column headers for the validation sheet
validationSheet ["A1"].Value = "Row Number";
validationSheet ["B1"].Value = "Validation Status";
validationSheet ["C1"].Value = "Phone Number Errors";
validationSheet ["D1"].Value = "Email Errors";
validationSheet ["E1"].Value = "Date of Birth Errors";

// Loop through each result and populate the worksheet
for (int index = 0; index < results.Count; index++)
{
    var currentResult = results[index];
    validationSheet [$"A{index + 2}"].Value = currentResult.Row;
    validationSheet [$"B{index + 2}"].Value = currentResult.IsValid ? "Yes" : "No";
    validationSheet [$"C{index + 2}"].Value = currentResult.PhoneNumberErrorMessage;
    validationSheet [$"D{index + 2}"].Value = currentResult.EmailErrorMessage;
    validationSheet [$"E{index + 2}"].Value = currentResult.DateErrorMessage;
}

// Save the workbook with validated data to a specific location
workbook.SaveAs(@"Spreadsheets\\ValidatedResults.xlsx");
```

<hr class="separator">

## 8. Data Export with Entity Framework ##

Leverage IronXL to transfer data from an Excel file to a database, or transform an Excel document into a database format. The example provided, `ExcelToDB`, demonstrates this by interpreting a spreadsheet containing GDP information by country and transferring it into an SQLite database. The process uses EntityFramework for database construction and processes the data for export on a row-by-row basis.

Integrate SQLite Entity Framework by installing the necessary NuGet packages.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

EntityFramework enables the creation of a model object for data export to a database.

Here's the paraphrased content for the given C# code section:

```cs
public class Nation
{
    [Key]
    public Guid ID { get; set; }  // Unique identifier for each nation
    public string CountryName { get; set; }  // Name of the country
    public decimal EconomicOutput { get; set; }  // GDP value in decimal format
}
```

This segment sets up the database context configuration. To integrate a different database system, you should first install the appropriate NuGet package and then locate the function corresponding to `UseSqLite()`.

Here is the paraphrased section:

```cs
/**
Entity Framework Data Export
anchor-export-data-using-entity-framework
**/
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        //TODO: Update to asynchronous approach
        Database.EnsureCreated(); // Ensure the database exists
    }

    /// <summary>
    /// Setup the DbContext to utilize a SQLite database
    /// </summary>
    /// <param name="optionsBuilder">Configuration options</param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var dbConnection = new SqliteConnection("Data Source=Country.db");
        dbConnection.Open(); // Open connection to the database file

        var dbCommand = dbConnection.CreateCommand();

        // Enable foreign keys in SQLite database
        dbCommand.CommandText = "PRAGMA foreign_keys = ON;";
        dbCommand.ExecuteNonQuery();

        optionsBuilder.UseSqlite(dbConnection); // Use SQLite with the current connection

        base.OnConfiguring(optionsBuilder);
    }
}
```

---
Instantiate a `CountryContext` and cycle through the specified range to build each record. Then call `SaveAsync` to save the changes to the database.

Here is the paraphrased section with relative URL paths resolved to ironsoftware.com:

```cs
public async Task ExecuteDataExportAsync()
{
    // Load the initial worksheet
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    // Establish a database connection
    using (var countryContext = new CountryContext())
    {
        // Process each cell in the worksheet
        for (var index = 2; index <= 213; index++)
        {
            // Access the range from columns A to B
            var range = worksheet [$"A{index}:B{index}"].ToList();

            // Instantiate a Country object for database insertion
            var country = new Country
            {
                Name = (string)range[0].Value,
                GDP = (decimal)(double)range[1].Value
            };

            // Enqueue the new Country object for database insertion
            await countryContext.Countries.AddAsync(country);
        }

        // Finalize and commit all additions to the database
        await countryContext.SaveChangesAsync();
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
        //TODO: Consider converting to async
        Database.EnsureCreated();
    }

    /// <summary>
    /// Sets up the context to utilize a Sqlite database
    /// </summary>
    /// <param name="optionsBuilder"></param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();

        var command = connection.CreateCommand();

        //Ensure the database is set to utilize foreign keys
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();

        optionsBuilder.UseSqlite(connection);

        base.OnConfiguring(optionsBuilder);
    }

}

// Function to process and save data asynchronously
public async Task ProcessDataAsync()
{
    //Loading the workbook
    var workbook = WorkBook.Load("Spreadsheets/GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");

    //Establishing the database connection
    using (var countryContext = new CountryContext())
    {
        //Looping through the worksheet rows
        for (var row = 2; row <= 213; row++)
        {
            //Selecting the range from column A to B
            var range = worksheet[$"A{row}:B{row}"].ToList();

            //Creating a new Country record
            var country = new Country
            {
                Name = (string)range[0].Value,
                GDP = (decimal)(double)range[1].Value
            };

            //Adding the new Country to the DbContext
            await countryContext.Countries.AddAsync(country);
        }

        //Saving the changes to the database
        await countryContext.SaveChangesAsync();
    }
}
```
In this revised segment, the script facilitates the exportation of data using Entity Framework to manage entries in a SQLite database. It initializes a database connection, configures it to enforce foreign key constraints, and asynchronously adds data from an Excel file to the database.

<hr class="separator">

## 9. Incorporate Formulas into a Spreadsheet ##

Apply formulas in Excel using the [`Cell`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) object, utilizing its [`Formula`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) attribute.

Below, you'll find an example that cycles through each state, calculating a percentage total for each and assigning it to column C.

```cs
/**
Set Spreadsheet Formulas
anchor-set-formulas-in-a-spreadsheet
**/
// Loop over all populated rows
for (var row = 2; row < i; row++)
{
    // Access cell in column C
    var currentCell = sheet [$"C{row}"].First();

    // Assign formula to calculate the percentage of total
    currentCell.Formula = $"=B{row}/B{i}";
}
```

```cs
/**
Insert Spreadsheet Formulae
anchor-insert-formulae-into-spreadsheet
**/
//Loop through each row containing data
for (var rowIndex = 2; rowIndex < i; rowIndex++)
{
    //Access the cell in column C
    var cell = sheet [$"C{rowIndex}"].First();

    //Assign a formula to calculate the Percentage of Total in column C
    cell.Formula = $"=B{rowIndex}/B{i}";
}
```

<hr class="separator">

## 10. Importing Data from an API into a Spreadsheet ##

The code snippet below demonstrates how to perform a REST call using [RestClient.Net](https://github.com/MelbourneDeveloper/RestClient.Net). This function retrieves JSON data and transforms it into a `List` of `RestCountry`. Afterward, the process involves iterating over each country to record the API data directly into an Excel spreadsheet.

```cs
/**
REST API to Excel Conversion
anchor-retrieve-data-from-api-for-spreadsheet
**/
var restClient = new Client(new Uri("https://restcountries.eu/rest/v2/"));
List<RestCountry> countryList = await restClient.GetAsync<List<RestCountry>>();
```

Example: *ApiToExcel*

Below is a visual representation of the JSON data from the API.
```

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

The code snippet below processes a list of countries and populates the spreadsheet with details such as the Name, Population, Region, Numeric Code, and the top three languages for each country.

Here's a paraphrased version of the provided C# code snippet:

```cs
// Loop through countries starting from the second element
for (int index = 2; index < countries.Count; index++)
{
    var currentCountry = countries[index];

    // Assign basic country information to the corresponding cells
    worksheet[$"A{index}"].Value = currentCountry.name;
    worksheet[$"B{index}"].Value = currentCountry.population;
    worksheet[$"G{index}"].Value = currentCountry.region;
    worksheet[$"H{index}"].Value = currentCountry.numericCode;

    // Loop through up to three languages and set them in subsequent columns
    for (int langIndex = 0; langIndex < 3; langIndex++)
    {
        // If the number of languages is less than the loop index, break out of the loop
        if (langIndex >= currentCountry.languages.Count) break;

        // Fetch language information
        var languageDetails = currentCountry.languages[langIndex];

        // Calculate the column letter based on the language index
        var columnLetter = GetColumnLetter(4 + langIndex);

        // Set the language name in the calculated column
        worksheet[$"{columnLetter}{index}"].Value = languageDetails.name;
    }
}
```

<hr class="separator">

<h2>API Reference and Resources</h2>

You might also find immense value in the [IronXL class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/) available in the API Reference section.

Furthermore, there are additional tutorials available that provide insights into various functionalities of IronXL.Excel, such as [Creating](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/), [Opening, Writing, Editing, Saving and Exporting](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/) XLS, XLSX, and CSV files without relying on Excel Interop.

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

