# C# Excel File Reading Guide

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp/>***


This guide provides instructions on how to read Excel files in C#, and covers common operations such as data validation, database transformations, Web API integrations, and altering formulas. The examples shown here make use of the IronXL .NET Excel library.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

IronXL simplifies the process of reading and manipulating Microsoft Excel files using C#. It operates independently without the need for Microsoft Excel or [Interop](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia). Moreover, IronXL offers [a more efficient and user-friendly API compared to `Microsoft.Office.Interop.Excel`](https://ironsoftware.com/csharp/excel/blog/compare-to-other-components/microsoft-office-excel-interop-alternative/).

## What IronXL Offers:

- Access to specialized support from our .NET engineering team.
- Simple setup through Microsoft Visual Studio.
- Complimentary trial for development purposes, with licensing starting from `$liteLicense`.

Utilizing the IronXL software library simplifies the process of reading and generating Excel documents in both C&num; and VB.NET.

### How to Read .XLS and .XLSX Files with IronXL

Here's a step-by-step guide for using the IronXL to read Excel files efficiently:

1. Begin by installing the IronXL Excel Library. This can be accomplished either via our [NuGet package](https://www.nuget.org/packages/IronXL.Excel/) or by downloading the [.Net Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) directly.

2. Utilize the `WorkBook.Load` function to open any XLS, XLSX, or CSV file.

3. Retrieve cell values with an easy-to-use syntax: `sheet["A11"].DecimalValue`.

Here's the paraphrased version of the specified section:

```cs
using System.Linq;
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section1
    {
        public void Run()
        {
            // Load the workbook. Supported file types are XLSX, XLS, CSV, and TSV
            WorkBook workbook = WorkBook.Load("test.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();

            // Easily access cells using Excel style cell references and fetch their integer value
            int valueAtCellA2 = worksheet["A2"].IntValue;

            // Gracefully read from a cell range
            foreach (var cell in worksheet["A2:A10"])
            {
                Console.WriteLine("Cell {0} value: '{1}'", cell.AddressString, cell.Text);
            }

            // Perform calculations on cell ranges like obtaining Sum, Min or Max values
            // Summation of range values
            decimal totalSum = worksheet["A2:A10"].Sum();

            // Calculation of maximum value using LINQ
            decimal maxValue = worksheet["A2:A10"].Max(cell => cell.DecimalValue);
        }
    }
}
```

The upcoming portions of this tutorial, including the sample project code, are designed to operate using three example Excel spreadsheets, as illustrated below:

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<hr class="separator">

<p class="main-content__segment-title">Tutorial</p>

## 1. Get Started with the IronXL C# Library for FREE

Before anything else, it's essential to integrate the `IronXL.Excel` library, which brings Excel capabilities within the .NET framework environment.

The simplest method to incorporate `IronXL.Excel` into your projects is through our NuGet package. Alternatively, you might consider downloading the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) directly and including it either in your project or the global assembly cache for broader use.
```

### How to Install the IronXL NuGet Package

To integrate the IronXL Excel library into your .NET project using Visual Studio, follow these steps:

1. Open your project in Visual Studio, right-click on the project node in the Solution Explorer, and choose "Manage NuGet Packages..." from the context menu.
2. In the NuGet Package Manager, use the search bar to look for `IronXL.Excel`. Once located, select the package and click the `Install` button to incorporate it into your project.

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" target="_blank">
    <p><img src="/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
  </a>

Alternatively, you can set up the IronXL library through the NuGet Package Manager Console:

1. Open the Package Manager Console in Visual Studio.
2. Execute the command:

   ```console
   PM> Install-Package IronXL.Excel
   ```
```

Here's the section paraphrased with the relative URL paths resolved:

```console
PM > Install-Package IronXL.Excel
```

Additionally, you can [explore the package on the NuGet website](https://www.nuget.org/packages/IronXL.Excel/).

### Manual Setup

If you prefer a hands-on approach, begin by acquiring the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and proceed with manually integrating it into your Visual Studio project.

## 2. Load an Excel Workbook

The [`WorkBook`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkBook.html) class embodies an Excel workbook. To initiate the opening of an Excel file in C#, the `WorkBook.Load` method is employed, wherein the file path is specified.

Here's the paraphrased section with updated paths:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section2
    {
        public void OpenWorkbook()
        {
            // Load an Excel workbook from a specified file
            WorkBook myWorkbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
        }
    }
}
``` 

In this snippet, the method name has been changed to `OpenWorkbook` to describe the action more precisely, and comments have been added for clarity.

Sample: *ExcelToDBProcessor*

Each instance of the `WorkBook` class can contain numerous `WorkSheet` objects, with each one corresponding to a single worksheet in your Excel file. To access a particular worksheet within an Excel workbook, you can use the `WorkBook.GetWorkSheet` method.

```csharp
using IronXL.Excel;

namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section3
    {
        public void Run()
        {
            WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
        }
    }
}
```
In this example, the `GetWorkSheet` method retrieves the worksheet named `"GDPByCountry"` from the existing workbook. For further details and methods regarding the `WorkSheet` class, you can refer to the full IronXL API documentation here: [IronXL WorkSheet API](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.WorkSheet.html).

# C# Read Excel File Tutorial

***Based on <https://ironsoftware.com/tutorials/how-to-read-excel-file-csharp/>***


This guide presents the method for reading an Excel file using C#, including common operations such as data validation, turning spreadsheets into databases, integrating with web APIs, and altering formulas. Throughout, we leverage the IronXL .NET Excel library for code illustrations.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

IronXL offers a robust solution for manipulating Microsoft Excel files in C#. It operates independently of Microsoft Excel and doesn't rely on [Interop](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia). Instead, [IronXL provides an API that is both faster and more user-friendly than `Microsoft.Office.Interop.Excel`](https://ironsoftware.com/csharp/excel/blog/compare-to-other-components/microsoft-office-excel-interop-alternative/).

## What IronXL Offers:

- Expert support from our .NET professionals
- Seamless integration with Microsoft Visual Studio
- A complimentary trial for development. Paid licenses starting from `$liteLicense`.

IronXL simplifies the processes of reading and constructing Excel files in both C# and VB.NET.

### Using IronXL to Read .XLS and .XLSX Files

Here’s a breakdown of how to read Excel files using IronXL:

1. Get the IronXL Excel Library using the NuGet package from [here](https://www.nuget.org/packages/IronXL.Excel/) or download the [.Net Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) directly.
2. Employ the `WorkBook.Load` function to open XLS, XLSX, or CSV files.
3. Fetch cell values using simple syntax such as `sheet ["A11"].DecimalValue`.

```cs
using System.Linq;
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section1
    {
        public void Run()
        {
            // Supports reading various spreadsheet formats: XLSX, XLS, CSV, and TSV
            WorkBook workbook = WorkBook.Load("test.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            
            // Access cells using Excel-style notation, retrieve the computed value
            int cellValue = worksheet["A2"].IntValue;
            
            // Elegant reading from cell ranges
            foreach (var cell in worksheet["A2:A10"])
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
            }
            
            // Execute advanced operations like calculating Min, Max, and Sum
            decimal sum = worksheet["A2:A10"].Sum();
            
            // Compatibility with Linq
            decimal max = worksheet["A2:A10"].Max(c => c.DecimalValue);
        }
    }
}
```

Further down in this tutorial, we use sample projects and code examples that apply to three sample Excel files:

[View the Excel samples](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png)

<hr class="separator">

<p class="main-content__segment-title">Tutorial</p>

## 1. Free Download of the IronXL C# Library

Start by installing the `IronXL.Excel` library, adding robust Excel capabilities to your .NET applications.

You can install `IronXL.Excel` directly via our NuGet package, or opt to download the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) manually to your project.

### How to Install the IronXL NuGet Package

1. Right-click on your project in Visual Studio, then choose "Manage NuGet Packages ..."
2. Search for the IronXL.Excel package and install it by clicking the Install button.

  [See NuGet installation guide](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png)

Alternatively, use the NuGet Package Manager Console:

1. Open the Package Manager Console.
2. Type `> Install-Package IronXL.Excel`.

  ```console
  PM > Install-Package IronXL.Excel
  ```

For more details, [visit the package on the NuGet site](https://www.nuget.org/packages/IronXL.Excel/).

### Manual Installation Steps

You can also start by downloading the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually adding it to your Visual Studio project.

## 2. Opening an Excel Workbook

The `WorkBook` class symbolizes an Excel document. To open an Excel File in C#, utilize the `WorkBook.Load` method and specify the Excel file's path.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
        }
    }
}
```

Example: *ExcelToDBProcessor*

Each `WorkBook` can contain multiple `WorkSheet` instances, each representing a different sheet within the Excel file. Use the `WorkBook.GetWorkSheet` method to acquire a specific sheet.

```c
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section3
    {
        public void Run()
        {
            WorkSheet worksheet = workBook.GetWorkSheet("GDPByCountry");
        }
    }
}
```

Example: *ExcelToDB*

### Creating New Excel Documents

To initiate a new Excel document, instantiate a `WorkBook` object specifying the desired file format.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workbook = new WorkBook(ExcelFileFormat.XLSX);
        }
    }
}
```

Example: *ApiToExcelProcessor*

Opt for `ExcelFileFormat.XLS` for older Excel versions (95 and earlier).

### Including a Worksheet in an Excel Document

Previously discussed, an IronXL `WorkBook` includes several `WorkSheet`s.

Here's a visual representation of a workbook with two worksheets in Excel:

[View Workbook Structure](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/work-book.png)

To create a new `WorkSheet`, call `WorkBook.CreateWorkSheet` with the worksheet's name.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section5
    {
        public void Run()
        {
            WorkSheet worksheet = workBook.GetWorkSheet("GDPByCountry");
        }
    }
}
```

## 3. Accessing Cell Values

### Read and Modify a Single Cell

Accessing individual cell values involves retrieving the desired cell from its `WorkSheet`, as demonstrated below:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section6
    {
        public you

### Generating New Excel Documents

To initiate a new Excel file, simply instantiate a new `WorkBook` class, specifying the desired Excel file format.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section4
    {
        public void Run()
        {
            // Initialize a new WorkBook object with the default Excel file format, XLSX.
            WorkBook workBook = new WorkBook(ExcelFileFormat.XLSX);
        }
    }
}
```

Here's the paraphrased section with relative URL paths resolved:

---
Sample: *ApiToExcelProcessor*

Please note, to accommodate older versions of Microsoft Excel (95 and earlier), specify `ExcelFileFormat.XLS` when creating new Excel documents. This ensures compatibility with legacy file formats.
---

### Creating a New Worksheet in an Excel Document

As we've discussed, an IronXL `WorkBook` encapsulates one or more `WorkSheet` objects. Each `WorkSheet` in the `WorkBook` represents a single spreadsheet page.

<div class="content-img-align-center">
  <a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" target="_blank">
    <img src="/img/tutorials/how-to-read-excel-file-csharp/work-book.png" alt="This is how one workbook with two worksheets looks in Excel." class="img-responsive add-shadow img-margin" style="max-width:100%;">
  </a>
  <p class="content__image-caption">This is how one workbook with two worksheets looks in Excel.</p>
</div>

To generate a new `WorkSheet`, use the `WorkBook.CreateWorkSheet` method and specify the desired name for the worksheet.

Below is the paraphrased content of the specified section with resolved URL paths:

```cs
using IronXL.Excel;
namespace ironxl.AccessExcelSheetCsharp
{
    public class RetrieveWorksheet
    {
        public void Execute()
        {
            // Accessing the worksheet named 'GDPByCountry' from the workbook
            WorkSheet worksheet = workBook.GetWorkSheet("GDPByCountry");
        }
    }
}
```

## 3. Accessing Cell Data

### Single Cell Access and Modification

The process of accessing and modifying the values of individual cells in a spreadsheet involves retrieving the specific cell using its associated `WorkSheet`. This can be done as outlined below:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section6
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            IronXL.Cell cell = workSheet["B1"].First();
        }
    }
}
```

The `Cell` class in IronXL represents a single cell in an Excel spreadsheet. It offers properties and methods that facilitate direct access and modification of the cell's content. Each `WorkSheet` handles a collection of `Cell` objects, each corresponding to a cell in the spreadsheet. In the example above, we access the cell located at B1 using standard array indexing syntax.

To modify or read data from the spreadsheet:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section7
    {
        public void Run()
        {
            IronXL.Cell cell = workSheet["B1"].First();
            string value = cell.StringValue;  // Retrieving the cell's content as a string
            Console.WriteLine(value);
            
            cell.Value = "10.3289";           // Updating the cell's content
            Console.WriteLine(cell.StringValue);
        }
    }
}
```

### Manipulating a Range of Cell Values

The `Range` class denotes a two-dimensional collection of `Cell` entities, representing a specific range of cells within an Excel document. You can obtain these ranges by using the string indexer on a `WorkSheet`:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section8
    {
        public void Run()
        {
            Range range = workSheet["D2:D101"];
        }
    }
}
```

For operations involving multiple cells within a known count, a simple loop can be effective:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section9
    {
        public void Run()
        {
            // Loop through the rows
            for (var y = 2; y <= 101; y++)
            {
                var result = new PersonValidationResult { Row = y };
                results.Add(result);
            
                // Extract all cells for a person
                var cells = workSheet[$"A{y}:E{y}"].ToList();
            
                // Validate phone number at column B
                var phoneNumber = cells[1].Value;
                result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);
            
                // Validate email address at column D
                result.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);
            
                // Extract and validate the date in a specific format at column E
                var rawDate = (string)cells[4].Value;
                result.DateErrorMessage = ValidateDate(rawDate);
            }
        }
    }
}
```

This method efficiently validates and processes multiple cell values in a spreadsheet dataset.

### Single Cell Access and Editing

Retrieving and editing values from individual cells within a spreadsheet is achieved by accessing the specific cell from its corresponding `WorkSheet`, as illustrated in the following example:

Here's the paraphrased section of the code using the IronXL library to load an Excel file and access a specific cell:

```cs
using IronXL.Excel;
namespace IronXLExample
{
    public class ExcelCellAccessDemo
    {
        public void Execute()
        {
            // Load an existing Excel workbook
            WorkBook workbook = WorkBook.Load("test.xlsx");
            // Access the default worksheet
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            // Retrieve the first cell in column B (B1)
            Cell selectedCell = worksheet["B1"].First();
        }
    }
}
```

This modified code snippet continues to demonstrate how to use the IronXL library to load an Excel file, access the default worksheet, and retrieve a specific cell within that worksheet.

The IronXL [`Cell`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) class is dedicated to representing individual cells within an Excel sheet. This class offers various properties and methods that facilitate both the retrieval and modification of cell values effectively.

Each `WorkSheet` within IronXL maintains a registry of `Cell` objects. This registry maps each `Cell` to a specific location in the spreadsheet, allowing for precise data manipulation. When accessing a cell, such as cell B1 in this instance, standard index notation (row and column) is used to specify the cell's position.

Once you obtain a `Cell` reference, it's straightforward to manipulate the data stored within the cell, either by reading or updating its contents.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section7
    {
        public void Run()
        {
            // Access the cell at position B1
            IronXL.Cell cell = workSheet["B1"].First();
            
            // Extract the string content of the cell
            string extractedValue = cell.StringValue;
            Console.WriteLine(extractedValue); // Display the string value

            // Update the cell's content to a new value
            cell.Value = "10.3289";
            
            // Output the updated value
            Console.WriteLine(cell.StringValue);
        }
    }
}
```

### Access and Modify Multiple Excel Cells Concurrently

The `Range` class in IronXL functions as a two-dimensional array of `Cell` instances. This configuration precisely maps a section of cells within an Excel spreadsheet. You can acquire these ranges by employing the string indexer associated with a `WorkSheet` object.

A range can be specified using single cell coordinates like "A1" or by defining a more expansive span of cells encompassing several rows and columns, such as "B2:E5". Additionally, one can utilize the `GetRange` method available on the `WorkSheet` to achieve similar results.

Here's the paraphrased section with resolved relative URL paths:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section8
    {
        public void Run()
        {
            // Initiating a Range object to hold cells from D2 to D101 in the worksheet
            Range range = workSheet["D2:D101"];
        }
    }
}
```

Here are some alternative methods for extracting or modifying the contents of cells within a `Range`. When you are aware of the number of cells in the range, a `for` loop can be utilized effectively. 

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section9 
    {
        public void Run() 
        {
            // Loop through specified rows
            for (int rowIndex = 2; rowIndex <= 101; rowIndex++) 
            {
                var validationResult = new PersonValidationResult { Row = rowIndex };
                results.Add(validationResult);
                
                // Accessing all cells in the row
                var cells = workSheet[$"A{rowIndex}:E{rowIndex}"].ToList();
                
                // Validate the phone number - column B (index 1)
                var phoneNumber = cells[1].Value;
                validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phoneNumber as string);
          
                // Validate the email address - column D (index 3)
                validationResult.EmailErrorMessage = ValidateEmailAddress(cells[3].Value as string);
                
                // Validate the date - column E (index 4), expected format: Month Day[suffix], Year
                var dateString = cells[4].Value as string;
                validationResult.DateErrorMessage = ValidateDate(dateString);
            }
        }
    }
}
```

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section9
    {
        public void Execute()
        {
            // Loop through each row in the selected range
            for (var rowIndex = 2; rowIndex <= 101; rowIndex++)
            {
                var validationResult = new PersonValidationResult { Row = rowIndex };
                results.Add(validationResult);

                // Retrieve the cells for the current individual
                var individualCells = workSheet[$"A{rowIndex}:E{rowIndex}"].ToList();

                // Perform phone number validation for the value in the second column (B)
                var phone = individualCells[1].Value;
                validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phone.ToString());

                // Perform email address validation for the value in the fourth column (D)
                var email = individualCells[3].Value;
                validationResult.EmailErrorMessage = ValidateEmailAddress(email.ToString());

                // Extract and validate the date from the fifth column (E)
                var dateText = individualCells[4].Value.ToString();
                validationResult.DateErrorMessage = ValidateDate(dateText);
            }
        }
    }
}
```

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section9
    {
        public void Run()
        {
            // Loop through the specified rows
            for (var y = 2; y <= 101; y++)
            {
                var validationResult = new PersonValidationResult { Row = y };
                results.Add(validationResult);
            
                // Retrieve all cells concerning an individual
                var cells = workSheet[$"A{y}:E{y}"].ToList();
            
                // Check validity of the phone number found in the second column
                var phoneNumber = cells[1].Value;
                validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);
            
                // Verify the email address situated in the fourth column
                validationResult.EmailErrorMessage = ValidateEmailAddress((string)cells[3].Value);
            
                // Extract and validate date from column five using a custom format
                var rawDate = (string)cells[4].Value;
                validationResult.DateErrorMessage = ValidateDate(rawDate);
            }
        }
    }
}
```

### Incorporating Formulas into Excel Cells

You can define formulas for cells using the `Formula` property on the [`Cell`](https://ironsoftware.com/csharp/excel/object-reference/api/IronXL.Cell.html) class.

Below is a walkthrough of the code that cycles through each state recording and assigns a percentage total to column C:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section10
    {
        public void Run()
        {
            // Iterate through all rows with a value
            for (var y = 2 ; y < i ; y++)
            {
                // Retrieve the cell in column C
                Cell cell = workSheet[$"C{y}"].First();
            
                // Apply a formula calculating the percentage of the total
                cell.Formula = $"=B{y}/B{i}";
            }
        }
    }
}
```

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section10
    {
        public void Run()
        {
            // Loop through each row containing data
            for (var rowIndex = 2; rowIndex < i; rowIndex++)
            {
                // Access the cell in column C for the current row
                Cell targetCell = workSheet[$"C{rowIndex}"].First();

                // Assign a formula to calculate the percentage relative to the total
                targetCell.Formula = $"=B{rowIndex}/B{i}";
            }
        }
    }
}
```

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section10
    {
        public void Run()
        {
            // Loop through all populated rows
            for (var y = 2 ; y < i ; y++)
            {
                // Access the cell in column C
                Cell cell = workSheet[$"C{y}"].First();
            
                // Assign a formula to calculate the percentage over total
                cell.Formula = $"=B{y}/B{i}";
            }
        }
    }
}
```

### Spreadsheet Data Validation

Leverage IronXL for data validation within spreadsheets. In the `DataValidation` example, we utilize `libphonenumber-csharp` for phone number validation along with standard C# APIs to ensure email addresses and dates are correctly formatted.

Below is the paraphrased section of the article you provided, with URL paths resolved to `ironsoftware.com`:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class ValidationResultProcessor
    {
        public void Execute()
        {
            // Loop through the specified range of rows
            for (int index = 2; index <= 101; index++)
            {
                var validationResult = new PersonValidationResult { Row = index };
                results.Add(validationResult);

                // Retrieve all relevant cells for the given person
                var personCells = worksheet[$"A{index}:E{index}"].ToList();

                // Phone number is located in the second cell (B column)
                var phone = personCells[1].Value;
                validationResult.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, phone.ToString());

                // Email address is in the fourth cell (D column)
                var email = personCells[3].Value;
                validationResult.EmailErrorMessage = ValidateEmailAddress(email.ToString());

                // Extract the date from the fifth cell (E column) and validate
                var date = personCells[4].Value;
                validationResult.DateErrorMessage = ValidateDate(date.ToString());
            }
        }
    }
}
```

The preceding code iterates over each row in the spreadsheet and collects the cells into a list. Each validation method scrutinizes the cell's value and yields an error message if the data is found to be incorrect.

This segment of code sets up a new worksheet, defines header titles, and records the error outcomes to produce a record of any data discrepancies.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section12
    {
        public void Run()
        {
            // Create a new worksheet called "Results" in the workbook
            var resultsSheet = workBook.CreateWorkSheet("Results");
            // Define headers for the results worksheet
            resultsSheet["A1"].Value = "Row";
            resultsSheet["B1"].Value = "Valid";
            resultsSheet["C1"].Value = "Phone Error";
            resultsSheet["D1"].Value = "Email Error";
            resultsSheet["E1"].Value = "Date Error";
            // Loop through each validation result and populate the worksheet
            for (var i = 0; i < results.Count; i++)
            {
                var result = results[i];
                resultsSheet[$"A{i + 2}"].Value = result.Row; // Set the row number
                resultsSheet[$"B{i + 2}"].Value = result.IsValid ? "Yes" : "No"; // Indicate if the row is valid
                resultsSheet[$"C{i + 2}"].Value = result.PhoneNumberErrorMessage; // Set the phone number error message
                resultsSheet[$"D{i + 2}"].Value = result.EmailErrorMessage; // Set the email error message
                resultsSheet[$"E{i + 2}"].Value = result.DateErrorMessage; // Set the date validation error message
            }
            // Save the workbook with validated data to a specified path
            workBook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");
        }
    }
}
```

## 4. Migrate Data from Spreadsheets to Databases Using Entity Framework

The IronXL library is adept at migrating data from Excel spreadsheets to databases seamlessly. For instance, the `ExcelToDB` example illustrates how to read GDP data from various nations from a spreadsheet and transfer it into an SQLite database.

This process employs the `EntityFramework` for constructing the database and systematically exporting the data.

Include the required SQLite Entity Framework NuGet packages to start this operation.

[![SQLite Entity Framework NuGet packages](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png)

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/ironxl-nuget.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

`EntityFramework` enables the creation of a model entity that facilitates the export of data to a database.

Here is the paraphrased section with resolved URL paths:

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section13
    {
        public void Run()
        {
            public class CountryData
            {
                [Key]
                public Guid ID { get; set; }  // Unique identifier for each country entry
                public string CountryName { get; set; }  // Name of the country
                public decimal GrossDomesticProduct { get; set; }  // GDP value stored as decimal
            }
        }
    }
}
```

To integrate a different database, you should install the appropriate NuGet package and look for the method analogous to `UseSqLite()`.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section14
    {
        public void Run()
        {
            // Define the DataContext for Country Entity
            public class CountryContext : DbContext
            {
                public DbSet<Country> Countries { get; set; } // DBSet for Country

                // Ensure database creation asynchronously
                public CountryContext()
                {
                    Database.EnsureCreated();
                }

                /// <summary>
                /// Setup SQLite configuration for the context
                /// </summary>
                /// <param name="optionsBuilder">Configures the database connection</param>
                protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
                {
                    // Establish the connection with the SQLite database
                    var connection = new SqliteConnection($"Data Source=Country.db");
                    connection.Open();

                    // Ensure foreign keys are enforced
                    var command = connection.CreateCommand();
                    command.CommandText = "PRAGMA foreign_keys = ON;";
                    command.ExecuteNonQuery();

                    // Utilize SQLite as the database engine
                    optionsBuilder.UseSqlite(connection);
                    base.OnConfiguring(optionsBuilder);
                }
            }
        }
    }
}
```

Set up a `CountryContext`, process each record within the designated range, and subsequently use the `SaveAsync` method to finalize and store the changes in the database.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section15
    {
        public void InitiateProcess()
        {
            public async Task ImportGDPDataAsync()
            {
                // Load the initial worksheet
                var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
                var worksheet = workbook.GetWorkSheet("GDPByCountry");
                // Establish a connection to the database
                using (var context = new CountryContext())
                {
                    // Traverse through each cell in the predefined range
                    for (var index = 2; index <= 213; index++)
                    {
                        // Capture the range for each country's GDP data
                        var cellRange = worksheet[$"A{index}:B{index}"].ToList();
                        // Instantiate a Country object to store in the database
                        var country = new Country
                        {
                            Name = (string)cellRange[0].Value,
                            GDP = (decimal)(double)cellRange[1].Value
                        };
                        // Queue the country for insertion
                        await context.Countries.AddAsync(country);
                    }
                    // Execute insert operations in the database
                    await context.SaveChangesAsync();
                }
            }
        }
    }
}
```

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section15
    {
        public async Task ProcessAsync()
        {
            // Load the worksheet
            var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
            var worksheet = workbook.GetWorkSheet("GDPByCountry");
            // Establish the database connection
            using (var countryContext = new CountryContext())
            {
                // Loop through the cells
                for (var i = 2; i <= 213; i++)
                {
                    // Access the range from columns A to B
                    var range = worksheet[$"A{i}:B{i}"].ToList();
                    // Construct a new Country entity for saving
                    var country = new Country
                    {
                        Name = (string)range[0].Value,
                        GDP = (decimal)(double)range[1].Value
                    };
                    // Prepare to save the entity
                    await countryContext.Countries.AddAsync(country);
                }
                // Save the changes to the database
                await countryContext.SaveChangesAsync();
            }
        }
    }
}
```

This code snippet illustrates a process for reading data from an Excel file and saving it to a database. It starts by loading a workbook and selecting a specific worksheet. Within a new database context, it iterates through specified cells, creating objects from the cell data, and adds them to the database. The process completes by committing these changes to the database. This example is a demonstration of integrating IronXL with a database using Entity Framework, focusing on an Excel to database conversion task.

## 5. Retrieve Data from an API and Save to a Spreadsheet

In this process, a REST request is made using the [RestClient.Net](https://github.com/MelbourneDeveloper/RestClient.Net) library. This action retrieves JSON data and transforms it into a list of `RestCountry` objects. Following this, iterating over each country to record the REST API data into an Excel file is straightforward.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section16
    {
        public async Task FetchCountriesData()
        {
            // Initialize the RestClient with the API endpoint
            var client = new Client(new Uri("https://restcountries.eu/rest/v2/"));
            
            // Make an asynchronous request to obtain a list of countries from the REST API
            List<RestCountry> countries = await client.GetAsync<List<RestCountry>>();
        }
    }
}
```

Here is the paraphrased section:

-----

Sample: *ApiToExcel*

The image below displays the JSON data received from the API. 

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

-----

<a rel="nofollow" href="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" target="_blank">
  <img src="/img/tutorials/how-to-read-excel-file-csharp/country-data.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

The subsequent code snippet cycles through a list of countries, assigning values for Name, Population, Region, NumericCode, and the top three Languages to corresponding cells in an Excel worksheet.

```cs
using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section17
    {
        public void Execute()
        {
            // Loop through each country in the list starting from the second element
            for (var index = 2; index < countries.Count; index++)
            {
                var country = countries[index];
                // Populate basic country details into the worksheet
                workSheet[$"A{index}"].Value = country.name;
                workSheet[$"B{index}"].Value = country.population;
                workSheet[$"G{index}"].Value = country.region;
                workSheet[$"H{index}"].Value = country.numericCode;
                
                // Process up to three languages per country, if available
                for (var languageIndex = 0; languageIndex < 3; languageIndex++)
                {
                    if (languageIndex >= country.languages.Count) break; // Exit loop if no more languages
                    
                    var language = country.languages[languageIndex];
                    // Calculate the Excel column letter dynamically
                    var columnLetter = GetColumnLetter(4 + languageIndex);
                    // Assign the language name to the appropriate cell
                    workSheet[$"{columnLetter}{index}"].Value = language.name;
                }
            }
        }
    }
}
```

<hr class="separator">

## Object Reference and Additional Learning Materials

Discover the comprehensive details and functionalities of IronXL by exploring the [IronXL class documentation](https://ironsoftware.com/csharp/excel/object-reference/api/), which can be an indispensable resource in your programming toolkit.

Further enhance your understanding of `IronXL.Excel` through our series of tutorials that cover a range of topics such as [Creating](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/), [Opening, Writing, Editing, Saving, and Exporting](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/) XLS, XLSX, and CSV files—all without the need for *Excel Interop*. These guides provide valuable insights into effective file management and manipulation using IronXL.

## Summary

IronXL.Excel stands as a versatile .NET library dedicated to reading various spreadsheet formats. It operates independently without the need for [Microsoft Excel](https://products.office.com/en-us/excel) installation and does not rely on Interop.

For those who appreciate the flexibility of the .NET library in editing Excel files, exploring the [Google Sheets API Client Library](https://developers.google.com/api-client-library/dotnet/apis/sheets/v4) for .NET could further enhance your capabilities, allowing seamless modifications to Google Sheets.

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

