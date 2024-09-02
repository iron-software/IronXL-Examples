# C# Excel File Reading Guide

This guide provides an overview of how to read Excel files in C#, and covers common operations such as data validation, transforming data for databases, integrating with web APIs, and adjusting formulas. The examples provided here make use of the IronXL .NET library for Excel operations.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

IronXL enhances the capabilities of C# for handling Microsoft Excel files, allowing for both reading and editing without the need for Microsoft Excel or [Interop](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia). Indeed, [IronXL offers an API that is both quicker and more user-friendly than `Microsoft.Office.Interop.Excel`](https://ironsoftware.com/csharp/excel/blog/compare-to-other-components/microsoft-office-excel-interop-alternative/).

## What IronXL Offers:

- Full support from our committed .NET engineering team
- Seamless installation through Microsoft Visual Studio
- Complimentary trial for developers, with licensing options starting from `$liteLicense`.

Utilizing the IronXL library simplifies the process of reading and generating Excel files in both C# and VB.NET.

### Reading .XLS and .XLSX Files with IronXL

Here's a step-by-step guide to the procedure for reading Excel files using IronXL:

1. **Install IronXL**: Start by installing the IronXL Excel library. This can be smoothly accomplished using our [NuGet package](https://www.nuget.org/packages/IronXL.Excel/), or by manually downloading the [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

2. **Load the Workbook**: Utilize the `WorkBook.Load` method to open any Excel document, whether it’s in XLS, XLSX, or CSV format.

3. **Retrieve Cell Values**: Access cell values swiftly using straightforward syntax like `sheet["A11"].DecimalValue`.

Here's the paraphrased code section, with improved code comments for clarity and understanding, and relative URL paths resolved:

```cs
using IronXL;
using System;
using System.Linq;

// Load the workbook from a file; supports various formats like XLSX, XLS, CSV, and TSV
WorkBook workbook = WorkBook.Load("test.xlsx");
// Automatically select the first worksheet from the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Easily access cells and retrieve their integer value using Excel-style references
int cellValue = worksheet["A2"].IntValue;

// Read from a range of cells elegantly and display their values
foreach (var cell in worksheet["A2:A10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}

// Perform advanced operations like aggregations on a cell range
// Calculate the sum of values in a range
decimal totalSum = worksheet["A2:A10"].Sum();

// Use LINQ to find the maximum decimal value in a range
decimal maximumValue = worksheet["A2:A10"].Max(c => c.DecimalValue);
```

The following code snippets and accompanying sample projects, detailed in the subsequent parts of this guide, are designed to operate with three example Excel spreadsheets. For a visual reference of these spreadsheets, see the image provided below:

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<hr class="separator">

<p class="main-content__segment-title">Tutorial</p>

## 1. Acquire the IronXL C# Library at No Cost

To get started, the initial step involves incorporating the `IronXL.Excel` library, which enhances the .NET framework with Excel capabilities.

You can conveniently install `IronXL.Excel` through our NuGet package. Alternatively, you have the option to [manually download and install the DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) directly into your project or the global assembly cache.

### How to Install the IronXL NuGet Package

To integrate the IronXL.Excel library using NuGet, follow these steps:

1. Open your project in Visual Studio, right-click on the project name and select "Manage NuGet Packages..."
   
2. In the NuGet package manager, search for `IronXL.Excel`. When the package appears, click on the "Install" button to incorporate it into your project.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
    <p><img src="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
  </a>

Alternatively, you can set up the IronXL library through the NuGet Package Manager Console using the following steps:

1. Open the Package Manager Console
2. Input the command: `> Install-Package IronXL.Excel`

```console
PM > Install-Package IronXL.Excel
```

You can also [check out the package on the NuGet website](https://www.nuget.org/packages/IronXL.Excel/).

### Manual Installation

Alternatively, you can begin by downloading the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually integrating it into Visual Studio.

## 2. Opening an Excel Workbook

The [`WorkBook`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) class signifies an Excel workbook in IronXL. To load an Excel file in C#, utilize the `WorkBook.Load` method and provide the file path as a parameter.

```cs
// Load the Excel Workbook from a specified path
WorkBook workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

Sample: *ExcelToDBProcessor*

A `WorkBook` in IronXL can contain several `WorkSheet` objects, with each one corresponding to an individual worksheet within the Excel document. To access a particular worksheet, utilize the [`WorkBook.GetWorkSheet`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) method which allows you to specify the worksheet by name.

```csharp
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

The `ExcelToDB` example illustrates how to integrate an Excel spreadsheet containing information about different countries' GDP into a SQLite database. This is accomplished by employing the Entity Framework, a robust ORM framework for .NET. 

Here’s how you can replicate this process:

1. **Prepare your environment:** 
   Ensure that you have all necessary SQLite Entity Framework packages added to your project through NuGet.

   ![Entity Framework and SQLite packages](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)

2. **Define your data model:**
   Create a model class called `Country`, which will map to your database table. This class should contain properties such as `Key`, `Name`, and `GDP` which correspond to the columns in your Excel data.

   ```cs
   public class Country
   {
       [Key]
       public Guid Key { get; set; }
       public string Name { get; set; }
       public decimal GDP { get; set; }
   }
   ```

3. **Configure the database context:**
   Create a context class named `CountryContext` derived from `DbContext`. This class manages the database operations. Configure it to connect to the SQLite database using the `UseSqlite` method from `DbContextOptionsBuilder`.

   ```cs
   public class CountryContext : DbContext
   {
       public DbSet<Country> Countries { get; set; }
       public CountryContext() { Database.EnsureCreated(); }

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

4. **Process the Excel data:**
   Load the spreadsheet using the `WorkBook.Load` method. Loop through the spreadsheet starting from the second row (to skip the header row), instantiate `Country` objects for each row, and save them into the database.

   ```cs
   public async Task ProcessAsync()
   {
       var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
       var worksheet = workbook.GetWorkSheet("GDPByCountry");

       using (var countryContext = new CountryContext())
       {
           for (int i = 2; i <= 213; i++) // Assuming there's data up to row 213
           {
               var range = worksheet[$"A{i}:B{i}"].ToList();
               var country = new Country
               {
                   Name = range[0].Value.ToString(),
                   GDP = Convert.ToDecimal(range[1].Value)
               };

               await countryContext.Countries.AddAsync(country);
           }
           await countryContext.SaveChangesAsync();
       }
   }
   ```

This sample demonstrates a practical application of IronXL in the context of data migration from an Excel document to a relational database, leveraging .NET tools and libraries effectively.

### Generating New Excel Files

To initiate a new Excel file, instantiate a fresh `WorkBook` object and specify the desired Excel file format.

Here is the paraphrased section with the resolved URL:

```cs
// Initialize a new WorkBook specifying the file format as XLSX
WorkBook workBook = new WorkBook(ExcelFileFormat.XLSX);
```

Here's the paraphrased section with corrected link paths:

---

Sample: *ApiToExcelProcessor*

Please note: To ensure compatibility with older versions of Microsoft Excel, specifically those from 1995 or earlier, utilize `ExcelFileFormat.XLS`.

### Creating a New Worksheet in an Excel Document

As previously mentioned, an IronXL `WorkBook` is a container that holds one or more `WorkSheet` objects.

<div class="content-img-align-center">
  <a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="This is how one workbook with two worksheets looks in Excel." class="img-responsive add-shadow img-margin" style="max-width:100%;">
  </a>
  <p class="content__image-caption">This is how one workbook with two worksheets looks in Excel.</p>
</div>

To generate a new `WorkSheet`, use the `WorkBook.CreateWorkSheet` method, specifying the desired worksheet name.

```cs
// Retrieve the WorkSheet named "GDPByCountry" from the WorkBook
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

## 3. Accessing Values in Excel Cells

Retrieving individual cell values becomes straightforward with IronXL by leveraging its robust structure for managing cells within a worksheet:

```cs
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;
IronXL.Cell cell = worksheet["B1"].First();
```

IronXL defines each cell in an Excel worksheet as an instance of the [`Cell`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) class, which provides properties and methods to directly manipulate cell data.

Once a cell is referenced, you can easily perform read and write operations to interact with the cell’s content:

```cs
IronXL.Cell cell = worksheet["B1"].First();
string value = cell.StringValue;   // Extracts the cell's string data
Console.WriteLine(value);

cell.Value = "10.3289";           // Updates the cell with a new value
Console.WriteLine(cell.StringValue);
```

Similar operations extend to handling multiple cells via the `Range` class, which denotes a collection of cells:

```cs
Range range = worksheet["D2:D101"];
```

The above example captures a specific set of cells, allowing iterative or collective operations, such as validation. Here's how to loop through the rows and validate each cell within a defined range:

```cs
// Loop through each row
for (var y = 2; y <= 101; y++)
{
    var result = new PersonValidationResult { Row = y };
    results.Add(result);

    // Fetch and validate data from each cell in the row
    var cells = worksheet[$"A{y}:E{y}"].ToList();

    // Validate specific data elements like phone numbers and email addresses
    var phoneNumber = cells[1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    var rawDate = (string)cells[4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

To streamline processes involving spreadsheet formulas, you can assign formulas directly to cells:

```cs
// Apply formula across rows
for (var y = 2; y < i; y++)
{
    // Access cell in column C
    Cell cell = worksheet[$"C{y}"].First();

    // Set formula to calculate percentage of total
    cell.Formula = $"=B{y}/B{i}";
}
```

IronXL thus provides a comprehensive suite for easy manipulation and validation of cell data in Excel files without the need for Microsoft Excel or Interop.

### Accessing and Editing Individual Spreadsheet Cells

To manipulate the data in specific spreadsheet cells, you can easily select the required cell from its respective `WorkSheet`. The process is demonstrated below:

```cs
// Load the Excel workbook
WorkBook workbook = WorkBook.Load("test.xlsx");
// Access the default worksheet in the workbook
WorkSheet worksheet = workbook.DefaultWorkSheet;
// Retrieve the first cell in the "B1" position
IronXL.Cell selectedCell = worksheet["B1"].First();
```

IronXL's [`Cell`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) class symbolizes an individual cell within an Excel spreadsheet. It is equipped with various properties and functions that facilitate both the access and modification of the cell's content.

Each `WorkSheet` maintains a catalog of `Cell` objects, each representing the data within the corresponding cell of an Excel worksheet. In the preceding example, we access a desired cell by utilizing its row and column indicators (in this illustration, cell B1), employing conventional indexing syntax used in arrays.

Once we possess a `Cell` reference, the cell's data can be both read from and written to, providing flexibility in managing spreadsheet content.

```cs
IronXL.Cell selectedCell = workSheet["B1"].First();
string cellValue = selectedCell.StringValue;  // Retrieve the string representation of the cell's value
Console.WriteLine(cellValue);

selectedCell.Value = "10.3289";  // Update the cell with a new value
Console.WriteLine(selectedCell.StringValue);  // Output the updated value
```

### Accessing and Manipulating Multiple Cells

The `Range` class encapsulates a two-dimensional array of `Cell` objects, essentially spanning a specific section of cells within an Excel sheet. To target these cells, you can leverage the string indexer method available on a `WorkSheet` object.

The indexer can accept input in two forms: a single cell coordinate (like "A1") or a continuous block defined by its diagonal corners (such as "B2:E5"). Alternatively, the `GetRange` method on a `WorkSheet` allows for similar access, enabling structured manipulation of cell blocks.

Here's the paraphrased section with the relative URL resolved:

```cs
// Access a sequence of cells vertically from D2 to D101
Range range = workSheet["D2:D101"];
```

For managing the values of cells within a specific `Range`, if the total number of cells is predetermined, leveraging a `for` loop is an effective strategy. This approach allows for systematic access and modification of each cell's contents.

Here is the paraphrased section of the article:

```cs
// Loop over each row starting from the second one up to the 101st
for (var y = 2; y <= 101; y++)
{
    // Instantiate a new result object for each row
    var result = new PersonValidationResult { Row = y };
    results.Add(result);

    // Retrieve all cells for a particular person ranging from A to E columns
    var cells = workSheet[$"A{y}:E{y}"].ToList();

    // Check the validity of the phone number located in the second cell (B column)
    var phoneNumber = cells[1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    // Validate the email address found in the fourth cell (D column)
    result.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);

    // Extract the date in its raw format from the fifth cell (E column), expecting Month Day[suffix], Year
    var rawDate = (string)cells[4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

```cs
// Iterate through each row in the spreadsheet
for (int i = 2; i <= 101; i++)
{
    // Create a container for results related to each person
    var validationResults = new PersonValidationResult { Row = i };
    results.Add(validationResults);

    // Retrieve all cell values for a single row
    var cellRange = worksheet[$"A{i}:E{i}"].ToList();

    // Validate the phone number located at Column B
    var phoneValue = cellRange[1].Value;
    validationResults.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phoneValue.ToString());

    // Check the email address at Column D
    validationResults.EmailErrorMessage = ValidateEmailAddress(cellRange[3].Value.ToString());

    // Validate the date at Column E, expected in "Month Day[suffix], Year" format
    var dateValue = cellRange[4].Value.ToString();
    validationResults.DateErrorMessage = ValidateDate(dateValue);
}

// The results are then compiled into a new worksheet to log validation errors
WorkSheet resultsSheet = workBook.CreateWorkSheet("Results");
resultsSheet["A1"].Value = "Row";
resultsSheet["B1"].Value = "Validation Status";
resultsSheet["C1"].Value = "Phone Number Errors";
resultsSheet["D1"].Value = "Email Errors";
resultsSheet["E1"].Value = "Date Errors";

// Populate the results sheet with validation outputs
for (int j = 0; j < results.Count; j++)
{
    var resultRecord = results[j];
    resultsSheet[$"A{j + 2}"].Value = resultRecord.Row;
    resultsSheet[$"B{j + 2}"].Value = resultRecord.IsValid ? "Yes" : "No";
    resultsSheet[$"C{j + 2}"].Value = resultRecord.PhoneNumberErrorMessage;
    resultsSheet[$"D{j + 2}"].Value = resultRecord.EmailErrorMessage;
    resultsSheet[$"E{j + 2}"].Value = resultRecord.DateErrorMessage;
}

// Save the filled workbook with path including the results sheet
workBook.SaveAs("Spreadsheets\\ValidationResults.xlsx");
```

### Incorporating Formulas into Excel Cells

Utilize the `Formula` property of the `Cell` class to assign formulas to cells. For more in-depth information on this property, refer to the [Formula documentation](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html).

The following snippet demonstrates how to cycle through states and compute the percentage totals for placement in column C:

```cs
// Loop through rows and set the formula for each cell in column C
for (int rowIndex = 2; rowIndex < totalRows; rowIndex++)
{
    // Access the cell in column C for the current row
    Cell currentCell = workSheet[$"C{rowIndex}"].First();

    // Assign a formula to calculate the percentage of the total
    currentCell.Formula = $"=B{rowIndex}/B{totalRows}";
}
```

Here's the paraphrased section of the code with explanations for each step:

```cs
// Loop over every populated row in the spreadsheet
for (var rowIndex = 2; rowIndex < i; rowIndex++)
{
    // Access the cell in column C for the current row
    Cell currentCell = workSheet[$"C{rowIndex}"].First();

    // Assign a formula calculating the percentage relative to the total in column B
    currentCell.Formula = $"=B{rowIndex}/B{i}";
}
```

```cs
// Look through every occupied row
for (var rowIndex = 2; rowIndex < rowCount; rowIndex++)
{
    // Select the cell in column C
    Cell currentCell = workSheet[$"C{rowIndex}"].First();

    // Compute the percentage and set as cell formula
    currentCell.Formula = $"=B{rowIndex}/B{rowCount}";
}
```

### Spreadsheet Data Validation with IronXL

Leverage IronXL for ensuring the accuracy of your spreadsheet data. In the `DataValidation` example, the `libphonenumber-csharp` library is utilized for phone number validation, while conventional C# APIs help in validating email addresses and dates.

Below is the paraphrased section of the article with relative URL paths resolved to ironsoftware.com:

```cs
// Loop through each row from 2 to 101
for (var index = 2; index <= 101; index++)
{
    var validationResult = new PersonValidationResult { Row = index };
    results.Add(validationResult);

    // Retrieve all cell values for a single person from columns A to E
    var personCells = worksheet[$"A{index}:E{index}"].ToList();

    // Extract and validate the phone number from the second column (B)
    var phone = personCells[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Extract and validate the email address from the fourth column (D)
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)personCells[3].Value);

    // Extract the date in the format 'Month Day[suffix], Year' from the fifth column (E)
    var dateInfo = (string)personCells[4].Value;
    validationResult.DateErrorMessage = ValidateDate(dateInfo);
}
```

The code provided iterates over each row in the spreadsheet, gathering the cells into a list. For each cell, validation methods assess the content and generate an error message for any that contain invalid data.

Additionally, this segment of code establishes a new worksheet, defines header labels, and logs the results of the error messages to track discrepancies found in the data.

```cs
// Initialize a new worksheet titled "Results"
var resultsSheet = workBook.CreateWorkSheet("Results");

// Setting header values for the columns
resultsSheet["A1"].Value = "Row";
resultsSheet["B1"].Value = "Validity";
resultsSheet["C1"].Value = "Phone Error Message";
resultsSheet["D1"].Value = "Email Error Message";
resultsSheet["E1"].Value = "Date Error Message";

// Iterating through the list of results to fill in each row
for (int i = 0; i < results.Count; i++)
{
    var result = results[i];
    // Assigning values to the rows starting from the second row (A2, B2, C2...)
    resultsSheet[$"A{i + 2}"].Value = result.Row;
    resultsSheet[$"B{i + 2}"].Value = result.IsValid ? "Yes" : "No";
    resultsSheet[$"C{i + 2}"].Value = result.PhoneNumberErrorMessage;
    resultsSheet[$"D{i + 2}"].Value = result.EmailErrorMessage;
    resultsSheet[$"E{i + 2}"].Value = result.DateErrorMessage;
}

// Saving the workbook to a specified file path
workBook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
```

## 4. Data Migration with Entity Framework

Utilize IronXL for transferring data from spreadsheets into a database or transforming an Excel file into a database structure. In the `ExcelToDB` example, we process a spreadsheet containing GDP data by country and subsequently migrate this data into an SQLite database.

This process leverages `EntityFramework` to construct the database incrementally and manage the data transfer seamlessly.

To begin, install the necessary SQLite Entity Framework packages via NuGet.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

`EntityFramework` provides the functionality to develop a model object that is capable of exporting data into a database.

```cs
public class Nation
{
    [Key]
    public Guid Identifier { get; set; }
    public string CountryName { get; set; }
    public decimal GrossDomesticProduct { get; set; }
}
```

To work with an alternative database, you should first obtain and install the appropriate NuGet package. Once installed, seek out the method within the library that functions similarly to `UseSqLite()`. This will help configure your application to interface with the chosen database.

```cs
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        // Initialize the database if not already created
        Database.EnsureCreated();
    }

    /// <summary>
    /// Sets up the database configuration to utilize SQLite.
    /// </summary>
    /// <param name="optionsBuilder">Used to configure the database options.</param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        // Establish a connection to the database
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();

        // Command to enable foreign key constraints
        var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();

        // Configure DbContext to use the SQLite database connection
        optionsBuilder.UseSqlite(connection);

        // Call the base method to complete configuration
        base.OnConfiguring(optionsBuilder);
    }
}
```

Create an instance of `CountryContext`, loop through the provided range to generate each record, and subsequently use `SaveAsync` to finalize the data entries in the database.

The following C# code asynchronously processes data for insertion into a database by iterating through specific cells in an Excel workbook:

```cs
public async Task ProcessDataAsync()
{
    // Load the workbook and select the appropriate sheet
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");
    // Create a new database connection within a using block for automatic disposal
    using (var context = new CountryContext())
    {
        // Loop through specified cells to gather data
        for (int index = 2; index <= 213; index++)
        {
            // Access the data in columns A and B for each row
            var cells = worksheet[$"A{index}:B{index}"].ToList();
            // Initialize a new Country object with data from the worksheet
            var country = new Country
            {
                Name = (string)cells[0].Value,
                GDP = (decimal)(double)cells[1].Value
            };
            // Add the new Country object to the database context
            await context.Countries.AddAsync(country);
        }
        // Save all changes made to the database
        await context.SaveChangesAsync();
    }
}
``` 

This modified section enhances readability and clarifies the actions performed by the code, such as loading data from an Excel file and using it to populate a database with `Country` entities.

# Export Data to a Database Using IronXL

With IronXL, converting an Excel spreadsheet into database entities is seamless. In the `ExcelToDB` example, the process begins by extracting GDP data by country from an Excel file to then be transferred into an SQLite database.

### Implementation Steps

Initialize the IronXL library and load the Excel workbook:

```cs
var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
var worksheet = workbook.GetWorkSheet("GDPByCountry");
```

Set up the database connection using Entity Framework for SQLite operations:

```cs
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }

    public CountryContext()
    {
        // Ensure the database is created
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

Define the `Country` entity:

```cs
public class Country
{
    [Key]
    public Guid Key { get; set; }
    public string Name { get; set; }
    public decimal GDP { get; set; }
}
```

Process each row in the worksheet, transform them into `Country` entities, and save to the database asynchronously:

```cs
public async Task ProcessAsync()
{
    using (var countryContext = new CountryContext())
    {
        for (var i = 2; i <= 213; i++) // Assuming rows start at 2 and end at 213
        {
            var cells = worksheet[$"A{i}:B{i}"].ToList();
            var country = new Country
            {
                Name = (string)cells[0].Value,
                GDP = (decimal)(double)cells[1].Value
            };

            await countryContext.Countries.AddAsync(country);
        }

        await countryContext.SaveChangesAsync();
    }
}
```

In this example, IronXL proves its efficiency not only in reading spreadsheet data but also in integrating with databases via Entity Framework, streamlining the data export process.

## 5. Download Data from an API to Spreadsheet

The code snippet below demonstrates how to execute a REST request using [RestClient.Net](https://github.com/MelbourneDeveloper/RestClient.Net) to retrieve JSON data. This data is then converted into a list of `RestCountry` objects. Following this, you can easily loop through each country's data and populate an Excel spreadsheet with the details fetched from the REST API.

```cs
// Create a new REST client instance pointing to the REST Countries V2 API
Client httpApiClient = new Client(new Uri("https://restcountries.eu/rest/v2/"));

// Asynchronously retrieve a list of countries using the REST client
List<RestCountry> listOfCountries = await httpApiClient.GetAsync<List<RestCountry>>();
```

Here's a visualization of the JSON data obtained from the API call:
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="Sample API JSON data" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>
```

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

The provided code snippet sequentially cycles through each country, extracting and logging individual data points such as Name, Population, Region, NumericCode, and the top three languages, recording these details into an Excel worksheet.

Here's a rephrased version of the provided section:

```cs
// Iterate starting from the second item in the countries list
for (int index = 2; index < countries.Count; index++)
{
    var country = countries[index];

    // Populate the basic details of the country
    workSheet[$"A{index}"].Value = country.name;
    workSheet[$"B{index}"].Value = country.population;
    workSheet[$"G{index}"].Value = country.region;
    workSheet[$"H{index}"].Value = country.numericCode;

    // Loop through the top three languages spoken in the country
    for (int languageIndex = 0; languageIndex < 3; languageIndex++)
    {
        // Exit loop if there are fewer than three languages
        if (languageIndex >= country.languages.Count) break;

        // Retrieve language details
        var language = country.languages[languageIndex];

        // Determine the column letter based on language index
        string columnLetter = GetColumnLetter(4 + languageIndex);

        // Assign the language to the corresponding cell in the worksheet
        workSheet[$"{columnLetter}{index}"].Value = language.name;
    }
}
```

<hr class="separator">

## Object Reference and Resources

For an in-depth understanding of IronXL's capabilities, the [IronXL class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/) found in the Object Reference section is highly informative.

Moreover, several tutorials offer insights into additional functionalities of `IronXL.Excel`, such as [Creating](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/), [Opening, Writing, Editing, Saving, and Exporting](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/) XLS, XLSX, and CSV files without the need for *Excel Interop*.

## Summary

IronXL.Excel stands out as a unique .NET library capable of handling numerous spreadsheet formats without the need for installing [Microsoft Excel](https://products.office.com/en-us/excel) or relying on Interop.

Should you find the .NET library beneficial for altering Excel documents, consider checking out the [Google Sheets API Client Library](https://developers.google.com/api-client-library/dotnet/apis/sheets/v4) for .NET, which facilitates the modification of Google Sheets.

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
      <h3>Download this Tutorial as C&num; Source Code</h3>
      <p>The full free C&num; for Excel Source Code for this tutorial is available to download as a zipped Visual Studio 2017 project file.</p>
      <a class="btn btn-white3" href="/csharp/excel/tutorials/downloads/How.to.Read.an.Excel.File.in.CSharp.zip">
        <i class="fa fa-cloud-download"></i> Download</a>
      </div>
  </div>
</div>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore this Tutorial on GitHub</h3>
      <p>The source code for this project is available in C&num; and VB.NET on GitHub.</p>
      <p>Use this code as an easy way to get up and running in just a few minutes. The project is saved as a Microsoft Visual Studio 2017 project, but is compatible with any .NET IDE.</p>
      <a class="doc-link" href="https://github.com/iron-software/tutorials/tree/master/IronXL/How%20to%20Read%20an%20Excel%20File%20in%20C%23" target="_blank">How to Read Excel File in C&num; on GitHub<i class="fa fa-chevron-right"></i></a>
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

