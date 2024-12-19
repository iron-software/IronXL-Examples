# C# Tutorial on Reading Excel Files

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp/>***


This guide provides a comprehensive overview of how to read Excel documents using C#, covering common tasks such as data validation, converting databases, integrating with Web APIs, and altering formulas. It includes practical code examples that demonstrate the use of the IronXL .NET Excel library.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

IronXL empowers the handling and modification of Microsoft Excel documents using C#. It operates independently of Microsoft Excel and does not utilize [Interop](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia). Moreover, [IronXL offers a more rapid and user-friendly API compared to `Microsoft.Office.Interop.Excel`](https://ironsoftware.com/csharp/excel/blog/compare-to-other-components/microsoft-office-excel-interop-alternative/).

## What IronXL Provides:

- Specialized support from our experienced .NET engineering team.
- Hassle-free setup through Microsoft Visual Studio.
- An opportunity to trial the software at no cost. Licensing starts at `$liteLicense`.

Utilizing IronXL, handling Excel files in both C# and VB.NET becomes straightforward and efficient.

### How to Read .XLS and .XLSX Files with IronXL

Here's a quick guide to reading Excel files with the IronXL library:

1. **Install IronXL**: Begin by installing the IronXL Excel Library. This can be accomplished by adding it through our [NuGet package](https://www.nuget.org/packages/IronXL.Excel/) or by directly downloading the [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

2. **Loading Documents**: Open your XLS, XLSX, or CSV files by employing the `WorkBook.Load` method.

3. **Accessing Cell Values**: Retrieve values from specific cells with a straightforward syntax, for example, `sheet["A11"].DecimalValue` to get the numeric value of cell A11.

Here is a paraphrased version of the provided C# code section, ensuring all relative links are resolved to `ironsoftware.com`:

```cs
using IronXL;
using System;
using System.Linq;

// The IronXL library supports reading spreadsheet formats such as XLSX, XLS, CSV, and TSV.
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet worksheet = workbook.GetFirstSheet();

// Access specific cells easily using Excel-like references and retrieve their integer values.
int valueAtCellA2 = worksheet["A2"].IntValue;

// Elegantly iterate through a range of cells and print their values.
foreach (var cell in worksheet["A2:A10"])
{
    Console.WriteLine("Cell {0} holds the value '{1}'", cell.AddressString, cell.Text);
}

// Perform advanced operations like calculating aggregate values, including minimum, maximum, and sum.
decimal totalSum = worksheet["A2:A10"].Sum();

// The worksheet is compatible with LINQ queries for even more complex data manipulation.
decimal maximumValue = worksheet["A2:A10"].Max(c => c.DecimalValue);
```

This rewrite maintains the original functionality and instructional purpose, applies the requested URL resolutions, and modifies the comments and variables for a fresh perspective.

In the upcoming sections of this guide (including the accompanying sample project code), you’ll be able to apply what you learn on three example Excel spreadsheets. Below is a visual representation of these spreadsheets:

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<hr class="separator">

<p class="main-content__segment-title">Tutorial</p>

## 1. Download the IronXL C# Library for FREE

To get started, you'll need to install the `IronXL.Excel` library to add Excel capabilities to your .NET framework.

The easiest way to install `IronXL.Excel` is through our NuGet package, but you also have the option to directly download and install the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) into your project or your global assembly cache.

### How to Install the IronXL NuGet Package

To incorporate IronXL into your project using Visual Studio, follow these simple steps:

1. Right-click on your project within Visual Studio and choose "Manage NuGet Packages..." from the context menu.

2. In the NuGet Package Manager, type `IronXL.Excel` in the search box and press Enter. Once the package appears, click the 'Install' button to add IronXL.Excel to your project.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
    <p><img src="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
  </a>

You can also add the IronXL library to your project via the NuGet Package Manager Console:

1. Open the Package Manager Console.
2. Run the command `> Install-Package IronXL.Excel`.

```console
PM > Install-Package IronXL.Excel
```

Additionally, you can [explore the package on the NuGet website](https://www.nuget.org/packages/IronXL.Excel/).

### Manual Setup Procedure

As an alternative, you have the option to manually download the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and proceed with its manual installation within Visual Studio.

## 2. Load an Excel Workbook

The [`WorkBook`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) class signifies an Excel workbook. To initiate the opening of an Excel file in C#, utilize the method `WorkBook.Load`, where the file's directory should be provided.

The following line of C# code demonstrates how to open an Excel file named "GDP.xlsx" located in the "Spreadsheets" directory on your computer using the IronXL library:

```cs
// Load the Excel file "GDP.xlsx" from the "Spreadsheets" folder
WorkBook workBook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

Sample: *ExcelToDBProcessor*

Multiple [`WorkSheet`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkSheet.html) objects can exist within a single `WorkBook`, with each representing an individual sheet within the Excel document. To access a particular `WorkSheet`, employ the [`WorkBook.GetWorkSheet`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) method which allows you to specify and retrieve any sheet by its name.

```csharp
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

In the *ExcelToDB* example, the process involves reading data from a spreadsheet and exporting it to an SQLite database, leveraging the capabilities of *Entity Framework* for database operations. This example is geared towards handling data for a collection of countries, each represented by a unique `GUID`, a name, and their respective GDP values.

```cs
public class Country
{
    [Key]
    public Guid Key { get; set; }
    public string Name { get; set; }
    public decimal GDP { get; set; }
}
```

Setting up the `DbContext` is crucial for facilitating database operations. Below is a setup that configures the context to utilize an SQLite database, ensuring that foreign keys are actively enforced:

```cs
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }
    public CountryContext()
    {
        Database.EnsureCreated();  // Make sure the database is created
    }
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();
        var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";  // Enforce foreign key relationships
        command.ExecuteNonQuery();
        optionsBuilder.UseSqlite(connection);  // Use SQLite as the database provider
        base.OnConfiguring(optionsBuilder);
    }
}
```

The operation to populate the database is executed asynchronously and iterates through each row from the worksheet, starting from the second row up to the 213th:

```cs
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
                Name = (string)range[0].Value,  // Assign the country name from Excel
                GDP = (decimal)(double)range[1].Value  // Convert and assign the GDP value
            };
            await countryContext.Countries.AddAsync(country);  // Add the country entity to the context
        }
        await countryContext.SaveChangesAsync();  // Commit the transaction to the database
    }
}
```

This example showcases an effective way to migrate Excel data into a structured database using C# and Entity Framework, simplifying data management tasks.

### Generating New Excel Documents

To initiate a new Excel document, instantiate a new `WorkBook` object and specify the desired file format.

```cs
// Create a new Excel workbook with the XLSX file format
WorkBook workbook = new WorkBook(ExcelFileFormat.XLSX);
```

```
Sample: *ApiToExcelProcessor*

Note: For compatibility with older versions of Microsoft Excel (95 and earlier), utilize `ExcelFileFormat.XLS`.
```

### Adding a Worksheet to an Excel Document

As mentioned earlier, an IronXL `WorkBook` encapsulates a collection containing one or several `WorkSheet` objects.

<div class="content-img-align-center">
  <a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="This is how one workbook with two worksheets looks in Excel." class="img-responsive add-shadow img-margin" style="max-width:100%;">
  </a>
  <p class="content__image-caption">This is how one workbook with two worksheets looks in Excel.</p>
</div>

To generate a new `WorkSheet`, use the `WorkBook.CreateWorkSheet` method and specify the worksheet's name.

```cs
// Retrieve a specific worksheet from an Excel workbook named "GDPByCountry"
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

## 3. Accessing Cell Values

### Read and Edit Individual Cells

To handle specific cell values within a spreadsheet, you can easily retrieve cells by leveraging IronXL's concise API. Here's an illustration:

```csharp
WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;
IronXL.Cell cell = workSheet["B1"].First();
```

In IronXL, the `Cell` class symbolizes a unique cell in the Excel sheet and is equipped with properties and methods that allow direct manipulation of the cell’s content.

To read or update the spreadsheet cells, consider the following example where cell values are accessed and modified using simple methods:

```csharp
IronXL.Cell cell = workSheet["B1"].First();
string currentValue = cell.StringValue;  // Reading the cell's value
Console.WriteLine(currentValue);

cell.Value = "10.3289";  // Updating the cell's value
Console.WriteLine(cell.StringValue);
```

### Handling a Range of Cells

A `Range` is a collection of `Cell` objects referring to a block of adjacent cells within a worksheet. You can acquire cell ranges using index notation:

```csharp
Range range = workSheet["D2:D101"];
```

To manipulate the cell contents within a Range, looping structures like `for` loops can be very helpful, especially if the size of the Range is predetermined. Below is an example on how to manage cell values across a range:

```csharp
// Iteration over the rows
for (var y = 2; y <= 101; y++)
{
    var result = new PersonValidationResult { Row = y };
    results.Add(result);

    // Fetch all cells associated with an individual
    var cells = workSheet[$"A{y}:E{y}"].ToList();

    // Assume validating phone number in column B
    var phoneNumber = cells[1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    // Validate the email in column D
    result.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);

    // Extract raw date from column E
    var rawDate = (string)cells[4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

### Incorporating Formulas in a Spreadsheet

To add Excel formulas into cells via IronXL, you can assign a formula string directly to the `Formula` property of a `Cell`:

```csharp
// Loop through rows with values
for (var y = 2; y < i; y++)
{
    // Access cell in column C
    Cell cell = workSheet[$"C{y}"].First();

    // Assign a formula calculating the percentage of the total
    cell.Formula = $"=B{y}/B{i}";
}
```

### Validate Spreadsheet Data

IronXL supports the validation of spreadsheet data. The following example demonstrates data validation using both internal and external libraries:

```csharp
// Iterate through spreadsheet rows
for (var i = 2; i <= 101; i++)
{
    var result = new PersonValidationResult { Row = i };
    results.Add(result);

    // Retrieve cells for validation for each row
    var cells = workSheet[$"A{i}:E{i}"].ToList();

    // Executing validation for phone number at column B
    var phoneNumber = cells[1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    // Email validation for column D
    result.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);

    // Date validation from the content of column E
    var rawDate = (string)cells[4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

This method aggregates validation errors and could ideally output a log to monitor spreadsheet data accuracy.

By navigating these functionalities, you can effectively manage and verify data in Excel spreadsheets with IronXL, making data handling tasks much simpler in .NET applications.

### Single Cell Access and Modification

Interacting with individual cell values within a spreadsheet involves fetching the specific cell from the `WorkSheet`, as demonstrated below:

Here's the paraphrased section of the code:

```cs
// Load the workbook from an existing file
WorkBook workbook = WorkBook.Load("test.xlsx");

// Access the default worksheet in the workbook
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Fetch the cell at position B1 from the worksheet
IronXL.Cell cell = worksheet["B1"].First();
```

The `Cell` class in IronXL represents a single cell within an Excel spreadsheet, complete with properties and methods that allow users to directly access or alter the cell's value.

Each `WorkSheet` in IronXL contains a collection of `Cell` objects, each corresponding to a cell in the spreadsheet. By using standard array indexing, such as in the demonstration above for cell B1, you can easily locate and reference any specific cell by its position.

Once a `Cell` object is selected, it becomes straightforward to both retrieve data from and write data to any cell in the spreadsheet.

The given C# code snippet illustrates how to manipulate the contents of a specific cell in an Excel worksheet using the IronXL library. Here is a paraphrased version of the code, which retains the same functionality:

```cs
// Access the cell at position B1 from the worksheet
IronXL.Cell selectedCell = workSheet["B1"].First();
// Retrieve and display the string content of the cell
string cellContent = selectedCell.StringValue;
Console.WriteLine(cellContent);

// Update the value of the cell to a new number
selectedCell.Value = "10.3289";
// Output the updated value of the cell
Console.WriteLine(selectedCell.StringValue);
```

### Handling Cell Value Ranges in Excel Documents

The `Range` class in IronXL is designed for handling a group of `Cell` objects laid out in a two-dimensional structure, mapping directly to a collection of cells in an Excel file. You can retrieve a range of cells by using the string indexer feature available on the `WorkSheet` object.

When specifying the range, you can either reference a single cell (for instance, "A1") or define a block of cells spanning from one corner to another (like "B2:E5"). Additionally, it's feasible to use the `GetRange` method on a `WorkSheet` to achieve the same result.

```cs
Range selectedRange = workSheet["D2:D101"];
```

```cs
// Example demonstrating cell manipulation within a Range in IronXL
// This approach assumes a pre-determined number of cells for iteration.

// Initialize a loop to iterate through a specific number of rows.
for (int rowIndex = 2; rowIndex <= 101; rowIndex++)
{
    // Example entity to hold validation results for easy tracking.
    var validationResults = new PersonValidationResult { Row = rowIndex };
    results.Add(validationResults);

    // Fetch all relevant cells for a person from columns A to E
    var personCells = workSheet[$"A{rowIndex}:E{rowIndex}"].ToList();

    // Validate phone number from column B (zero-based index 1)
    var phoneValue = personCells[1].Value;
    validationResults.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneValue);

    // Validate email from column D (zero-based index 3)
    validationResults.EmailErrorMessage = ValidateEmailAddress((string)personCells[3].Value);

    // Validate a custom formatted date from column E (zero-based index 4)
    var dateInText = (string)personCells[4].Value;
    validationResults.DateErrorMessage = ValidateDate(dateInText);
}
```
This revised script skillfully navigates the validation process for data within specified cell ranges, ensuring each step is clear and concise for effective data management within a spreadsheet. It demonstrates practical use of a loop to iterate through cell values for validation in IronXL, adhering to the best practices of .NET programming.

```cs
// Loop through each row
for (var rowIndex = 2; rowIndex <= 101; rowIndex++)
{
    var validationResult = new PersonValidationResult { Row = rowIndex };
    results.Add(validationResult);

    // Retrieve all cells for an individual person
    var cells = workSheet[$"A{rowIndex}:E{rowIndex}"].ToList();

    // Extract and validate the phone number from column B
    var phone = cells[1].Value;
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Extract and validate the email address from column D
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);

    // Extract the raw date in "Month Day[suffix], Year" format from column E
    var date = (string)cells[4].Value;
    validationResult.DateErrorMessage = ValidateDate(date);
}
```

In this section of the C# tutorial for using IronXL, we explore data validation through a practical example labeled *DataValidation*. We'll be using the IronXL library along with the `libphonenumber-csharp` library to validate phone numbers, and standard C# functionalities for validating email addresses and dates.

```cs
// Iterate over each row starting from the second
for (var i = 2; i <= 101; i++)
{
    var result = new PersonValidationResult { Row = i };
    results.Add(result);

    // Gather all cells for an individual from columns A to E
    var cells = worksheet[$"A{i}:E{i}"].ToList();

    // Validate the phone number, stored in column B
    var phoneNumber = cells[1].Value;
    result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);

    // Validate the email address, found in column D
    result.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);

    // Fetch and validate the date from column E, formatted as "Month Day[suffix], Year"
    var rawDate = (string)cells[4].Value;
    result.DateErrorMessage = ValidateDate(rawDate);
}
```

This code demonstrates how to loop through rows in the spreadsheet to collect individual entries contained between columns A and E. For each row, the script validates phone numbers, email addresses, and date entries, logging any errors encountered into a `PersonValidationResult` object.

The next part of the code creates a new Worksheet and logs the results, ensuring there is a record of all data with validation issues.

```cs
var resultsSheet = workBook.CreateWorkSheet("Results");
resultsSheet["A1"].Value = "Row";
resultsSheet["B1"].Value = "Valid";
resultsSheet["C1"].Value = "Phone Error";
resultsSheet["D1"].Value = "Email Error";
resultsSheet["E1"].Value = "Date Error";
for (var i = 0; i < results.Count; i++)
{
    var result = results[i];
    resultsSheet[$"A{i + 2}"].Value = result.Row;
    resultsSheet[$"B{i + 2}"].Value = result.IsValid ? "Yes" : "No";
    resultsSheet[$"C{i + 2}"].Value = result.PhoneNumberErrorMessage;
    resultsSheet[$"D{i + 2}"].Value = result.EmailErrorMessage;
    resultsSheet[$"E{i + 2}"].Value = result.DateErrorMessage;
}
workBook.SaveAs("https://ironsoftware.com/csharp/excel/Spreadsheets/PeopleValidated.xlsx");
```

By orchestrating the `IronXL.WorkBook` and `IronXL.WorkSheet` classes effectively, the code not only validates data but also logs detailed results in a newly created `WorkSheet` named "Results". The validated data is then saved onto a new Excel file, ensuring that all data handling is clear and accountable.

### Implementing Formulas in a Spreadsheet

To assign formulas to cells, the `Formula` property of the `Cell` class is used, as detailed in the [IronXL Cell class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html).

The ensuing code demonstrates how to loop through each state in a list and calculate percentage totals in column C.

```csharp
// Cycle through each row setting the formula in column C
for (int y = 2; y < i; y++)
{
    // Access the cell in column C for the current row
    Cell cell = workSheet[$"C{y}"].First();

    // Assign a formula to calculate the percentage of the total
    cell.Formula = $"=B{y}/B{i}";
}
```

```cs
// Loop through every row that contains data, starting from the second row
for (var rowIndex = 2; rowIndex < i; rowIndex++)
{
    // Access the cell in column C for the current row
    Cell currentCell = workSheet[$"C{rowIndex}"].First();

    // Assign a formula to calculate the percentage of the total value
    currentCell.Formula = $"=B{rowIndex}/B{i}";
}
```

# Adding Formulas in a Spreadsheet

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp/>***


When adding formulas to a cell within a spreadsheet using IronXL, you can easily compute values dynamically based on other cell values. Below, you'll find an example to help you understand how to utilize formulas effectively. This example focuses on updating column C with the percentage totals based on values in column B.

```cs
// Loop through rows that contain values
for (var rowIndex = 2; rowIndex < i; rowIndex++)
{
    // Access the cell in column C
    Cell cell = workSheet[$"C{rowIndex}"].First();

    // Set the formula to calculate the percentage of the total
    cell.Formula = $"=B{rowIndex}/B{i}";
}
```

This method illustrates handling cell formulas programmatically, aiming to update each cell in column C with a formula reflecting the proportion of the value in the corresponding row of column B to the total in cell B(i).

### Spreadsheet Data Validation

IronXL excels in ensuring the accuracy of your data within spreadsheets. In the `DataValidation` example, we utilize `libphonenumber-csharp` for phone number verification along with standard C# APIs to check the validity of email addresses and dates. These tools work in tandem to ensure that the data adheres to specified formats and standards.

Here's the paraphrased section with actionable comments and improved clarity, and all relative URL paths resolved to `ironsoftware.com`:

```cs
// Loop through each row from 2 to 101
for (var rowIndex = 2; rowIndex <= 101; rowIndex++)
{
    // Initialize validation result for the current row
    var validationResult = new PersonValidationResult { Row = rowIndex };
    results.Add(validationResult);

    // Retrieve cell data for a single person spanning columns A to E
    var personCells = worksheet[$"A{rowIndex}:E{rowIndex}"].ToList();

    // Extract and validate the phone number from column B
    var phone = personCells[1].Value; // Column index 1 corresponds to column B
    validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phone);

    // Validate the email address from column D
    var email = personCells[3].Value; // Column index 3 corresponds to column D
    validationResult.EmailErrorMessage = ValidateEmailAddress((string)email);

    // Parse and validate date from column E, expected format: Month Day(suffix), Year
    var dateInfo = (string)personCells[4].Value; // Column index 4 corresponds to column E
    validationResult.DateErrorMessage = ValidateDate(dateInfo);
}
```

In this revision, variable names and comments have been enhanced for better readability and maintenance.

The code snippet provided iterates over each row in the spreadsheet, retrieving the cells into a list format. For each cell, it applies validation methods that inspect the cell's content and return an error message if the content does not meet the specified criteria.

In addition, this segment of code constructs a new worksheet, sets up column headers, and compiles the results of the error messages. This results in a structured record of all data that failed to validate correctly, ensuring a thorough log of any discrepancies or issues found within the spreadsheet data.

```cs
// Create a new worksheet for results
var resultsSheet = workBook.CreateWorkSheet("Results");

// Define the headers for result columns
resultsSheet["A1"].Value = "Row";
resultsSheet["B1"].Value = "Validated";
resultsSheet["C1"].Value = "Phone Issue";
resultsSheet["D1"].Value = "Email Issue";
resultsSheet["E1"].Value = "Date Issue";

// Iterate over the validation results and populate the sheet
for (int index = 0; index < results.Count; index++)
{
    var result = results[index];
    // A results start from row 2 in the spreadsheet
    resultsSheet[$"A{index + 2}"].Value = result.Row;
    resultsSheet[$"B{index + 2}"].Value = result.IsValid ? "Yes" : "No";
    resultsSheet[$"C{index + 2}"].Value = result.PhoneNumberErrorMessage;
    resultsSheet[$"D{index + 2}"].Value = result.EmailErrorMessage;
    resultsSheet[$"E{index + 2}"].Value = result.DateErrorMessage;
}

// Save the workbook with results
workBook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
```

## 4. Exporting Data with Entity Framework

IronXL is powerful for exporting data directly from an Excel file to a database or transforming an Excel sheet into a database format. For example, the `ExcelToDB` sample demonstrates how IronXL takes a spreadsheet containing GDP data by country and exports it into an SQLite database.

The process relies on `EntityFramework` to construct the database structure and progressively exports the data row by row.

You will need to include SQLite Entity Framework NuGet packages into your project for this functionality.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

`EntityFramework` enables the creation of a model object capable of exporting data into a database.

```cs
public class Nation
{
    [Key]
    public Guid ID { get; set; } // Unique identifier for each record
    public string CountryName { get; set; } // Name of the country
    public decimal GrossDomesticProduct { get; set; } // Economic measure of a nation's total value of production
}
```

For integrating a different database, you'll need to download and install the appropriate NuGet package and then seek the method that functions similarly to `UseSqLite()`.

```cs
public class CountryContext : DbContext
{
    public DbSet<Country> Countries { get; set; }
    public CountryContext()
    {
        // Set this operation to be performed asynchronously later
        Database.EnsureCreated();  // Ensures that the database is created upon initialization
    }
    /// <summary>
    /// Sets up the context to utilize an SQLite database
    /// </summary>
    /// <param name="optionsBuilder">The options builder to configure</param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var connection = new SqliteConnection("Data Source=Country.db");
        connection.Open();  // Open the connection to the Sqlite Database
        var command = connection.CreateCommand();
        // Ensure that foreign keys are enforced in the SQLite database
        command.CommandText = "PRAGMA foreign_keys = ON;";
        command.ExecuteNonQuery();  // Execute the SQL command to modify database settings
        optionsBuilder.UseSqlite(connection);  // Use SQLite with the current settings
        base.OnConfiguring(optionsBuilder);
    }
}
```

Construct a `CountryContext`, traverse through the specified range to generate each record, and subsequently utilize `SaveAsync` to finalize the entries in the database.

Here is the paraphrased section of the article with the resolved URL paths for any links or images to `ironsoftware.com`:

```cs
public async Task ExecuteDatabaseExportAsync()
{
    // Load the initial worksheet
    var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
    var worksheet = workbook.GetWorkSheet("GDPByCountry");
    // Establish connection to the database
    using (var db = new CountryContext())
    {
        // Loop through each cell in the workbook
        for (var index = 2; index <= 213; index++)
        {
            // Extract data from column A to B
            var dataRange = worksheet[$"A{index}:B{index}"].ToList();
            // Instantiate a country object for database insertion
            var countryRecord = new Country
            {
                Name = (string)dataRange[0].Value,
                GDP = (decimal)(double)dataRange[1].Value
            };
            // Queue the country entity for insertion
            await db.Countries.AddAsync(countryRecord);
        }
        // Persist changes to the database
        await db.SaveChangesAsync();
    }
}
```

## Exporting Data from Excel to a Database Using IronXL

This task involves reading data from an Excel spreadsheet and transferring it to an SQLite database by leveraging the IronXL library coupled with the Entity Framework. The dataset example uses a spreadsheet with GDP data by country.

1. **Installation of Necessary Packages**: Begin by adding the required SQLite Entity Framework packages to your project. You can find these on your respective package management console or NuGet gallery.

   ![NuGet Package Installation](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)

2. **Model Definition**: Define a model that mirrors the data structure in the database. This is done via a class in C# with properties that represent the database fields.

   ```cs
   public class Country
   {
       [Key]
       public Guid Key { get; set; }
       public string Name { get; set; }
       public decimal GDP { get; set; }
   }
   ```

3. **Database Context Setup**: Configure your database context for using SQLite. This includes setting up connection strings and ensuring the database is created if it does not exist.

   ```cs
   public class CountryContext : DbContext
   {
       public DbSet<Country> Countries { get; set; }

       public CountryContext()
       {
           Database.EnsureCreated();  // Ensure database is created on initialization
       }

       protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
       {
           var connectionString = "Data Source=Country.db";
           optionsBuilder.UseSqlite(connectionString); // Use SQLite database
       }
   }
   ```

4. **Reading and Importing Data**: Load the spreadsheet, read the data, and populate the database.

   ```cs
   public async Task ImportGDPData()
   {
       var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx"); // Load the workbook
       var worksheet = workbook.GetWorkSheet("GDPByCountry"); // Get the specific worksheet

       using (var countryContext = new CountryContext()) // Instantiate the database context
       {
           for (int rowIndex = 2; rowIndex <= 213; rowIndex++) // Start from row 2 to skip headers
           {
               var range = worksheet[$"A{rowIndex}:B{rowIndex}"].ToList();
               var country = new Country
               {
                   Name = range[0].StringValue,
                   GDP = Convert.ToDecimal(range[1].DoubleValue)
               };

               await countryContext.Countries.AddAsync(country); // Add country data to the context
           }

           await countryContext.SaveChangesAsync(); // Save changes asynchronously
       }
   }
   ```

In this scenario, every row in the Excel spreadsheet from the indicated range is read, converted into a `Country` object, and saved to the database. This method efficiently transfers large datasets from Excel to a structured database format.

## 5. Importing Data from an API into a Spreadsheet

The process initiates with a REST API call using [RestClient.Net](https://github.com/MelbourneDeveloper/RestClient.Net), which retrieves JSON data. This data is subsequently transformed into a "List" of `RestCountry` objects. Following this transformation, each country's data is seamlessly transferred and stored into an Excel file.

```cs
// Create a new REST client instance pointing to the REST Countries API
var restClient = new Client(new Uri("https://restcountries.eu/rest/v2/"));

// Asynchronously fetch a list of countries from the REST API
List<RestCountry> countries = await restClient.GetAsync<List<RestCountry>>();
```

Sample: *ApiToExcel* 

Here's a visual representation of the JSON data obtained from the API.

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>
```

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

The code snippet below cycles through each country's data, assigning values for Name, Population, Region, NumericCode, as well as the top three languages to the corresponding columns in the Excel spreadsheet.

Below is the rephrased code snippet from the provided section:

```cs
for (var index = 2; index < countries.Count; index++)
{
    var currentCountry = countries[index];
    // Assign the fundamental attributes
    workSheet[$"A{index}"].Value = currentCountry.name;
    workSheet[$"B{index}"].Value = currentCountry.population;
    workSheet[$"G{index}"].Value = currentCountry.region;
    workSheet[$"H{index}"].Value = currentCountry.numericCode;
    // Loop through the top 3 languages, if available
    for (var langIndex = 0; langIndex < 3; langIndex++)
    {
        if (langIndex >= currentCountry.languages.Count) break;
        var language = currentCountry.languages[langIndex];
        // Determine the corresponding column letter
        var columnLetter = GetColumnLetter(4 + langIndex);
        // Place the language name in the appropriate cell
        workSheet[$"{columnLetter}{index}"].Value = language.name;
    }
}
```

In this version, variable names are more descriptive, enhancing the readability of what each part of the loop does and making it clearer for other developers who might work with this code in the future.

<hr class="separator">

## Object Reference and Resources

The [IronXL class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/) available in the Object Reference section is a useful resource.

Furthermore, you can explore additional tutorials that provide insights into various functionalities of `IronXL.Excel`. These include guides on [Creating](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/), [Opening, Writing, Editing, Saving, and Exporting](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/) XLS, XLSX, and CSV files all without the need for *Excel Interop*.

## Summary

IronXL.Excel stands alone as a .NET software library, supporting a vast range of spreadsheet formats. It operates independently without the need for installing [Microsoft Excel](https://products.office.com/en-us/excel) or relying on Interop.

Should you find the .NET library beneficial for altering Excel documents, consider delving into the [Google Sheets API Client Library](https://developers.google.com/api-client-library/dotnet/apis/sheets/v4) designed for .NET, enabling the modification of Google Sheets.

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

