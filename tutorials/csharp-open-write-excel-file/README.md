# C# Write to Excel [Interop-Free] Code Example Tutorial

Discover a series of detailed guides on how to create, open, save, and manipulate Excel files using C#, covering operations like calculating sums, averages, counts, and more. IronXL.Excel is a self-contained .NET library capable of handling various spreadsheet formats without the need for Microsoft Excel or reliance on Interop services to be installed.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

<h2>Use IronXL to Open and Write Excel Files</h2>

Open, write, save, and modify Excel documents with the user-friendly [IronXL C# library](https://ironsoftware.com/csharp/excel/).

Download a [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or integrate it into your own project and follow this guide:

1. Obtain the IronXL Excel Library via [NuGet](https://www.nuget.org/packages/IronXL.Excel) or by downloading the DLL directly.
2. Utilize the `WorkBook.Load` method to access content from XLS, XLSX, or CSV files.
3. Retrieve cell values easily with syntax like `sheet["A11"].DecimalValue`.

Throughout this tutorial, you will learn about:

- **Installing IronXL.Excel**: Incorporating the IronXL.Excel library into your existing project.
- **Basic Operations**: Steps for creating or opening a workbook, selecting sheets and cells, and saving your work.
- **Advanced Sheet Operations**: Mastering advanced features such as inserting headers and footers, performing mathematical operations, and other advanced functionalities.

<h4>Open an Excel File : Quick Code</h4>

```cs
using IronXL;

// Load the workbook from the file system
WorkBook workbook = WorkBook.Load("test.xlsx");
// Access the default worksheet
WorkSheet worksheet = workbook.DefaultWorkSheet;
// Define a range of cells to work with
IronXL.Range cellRange = worksheet["A2:A8"];
decimal sumOfValues = 0;

// Loop through each cell in the defined range
foreach (var cell in cellRange)
{
    // Output the row index and cell value
    Console.WriteLine("Cell {0} holds the value '{1}'", cell.RowIndex, cell.Value);
    // Check if the cell contains a numeric value
    if (cell.IsNumeric)
    {
        // Accumulate the decimal value of the cell
        sumOfValues += cell.DecimalValue;
    }
}

// Validate the sum against a predefined cell value for consistency
if (worksheet["A11"].DecimalValue == sumOfValues)
{
    Console.WriteLine("Validation Succeeded");
}
```

<h4>Write and Save Changes to the Excel File : Quick Code</h4>

Here's a paraphrased version of the provided code section from the article:

```cs
// Assign a decimal value to a specific cell within the Excel spreadsheet
workSheet["B1"].Value = 11.54;

// Persist modifications by saving the workbook to a file
workBook.SaveAs("test.xlsx");
```

<hr class="separator">
<p class="main-content__segment-title">Step 1</p>

## 1. Download and Install the Free IronXL C# Library

IronXL.Excel offers a robust and versatile library, designed for handling Excel files in .NET environments. It supports integration across various .NET project types including Windows applications, ASP.NET MVC, and .NET Core applications. This allows developers to read, edit, save, and create Excel documents without needing Excel installed on their machine.

<h3>Install the Excel Library to your Visual Studio Project with NuGet</h3>

The initial step involves incorporating the IronXL.Excel library into your project. You have two options to achieve this: by using the NuGet Package Manager or through the NuGet Package Manager Console.

To integrate the IronXL.Excel library via the NuGet Package Manager in a user-friendly way, follow these steps:

1. Navigate with your mouse, right-click on the project name in your solution, and select 'Manage NuGet Packages'.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <p><img src="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

2. Navigate to the 'Browse' tab, enter `IronXL.Excel` in the search field, and select the 'Install' option.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" target="_blank">

```html
<p><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
```

</a>

3. Installation Complete

The IronXL.Excel library is now successfully added to your project, wrapping up the installation process.

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>



<h3>Install Using NuGet Package Manager Console</h3>

1. Navigate to the `NuGet Package Manager` by selecting `Package Manager Console` from the `tools` menu.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

Execute the command below to install the `IronXL.Excel` package, specifying the version `2019.5.2`.

```plaintext
PM > Install-Package IronXL.Excel -Version 2019.5.2
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<h3>Manually Install with the DLL</h3>

You also have the option to directly install the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) into your project or into the global assembly cache manually.

Here is the paraphrased version:

```
 PM > Install-Package IronXL.Excel
```

# C# Writing to Excel Files: A Code Example Tutorial without Interop

Explore the step-by-step guide on how to create, open, and save Excel files using C# without installing Microsoft Excel or relying on Interop. Discover the versatility of IronXL.Excel, a comprehensive .NET library designed for managing various spreadsheet formats.

---

### Introduction

#### Leveraging IronXL for Excel Manipulations

Easily manage Excel documents such as creating, saving, and manipulating contents using the intuitive IronXL C# library. Get started by [downloading a sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or integrating IronXL into your existing projects.

To begin, add the IronXL Excel Library to your project via [NuGet](https://www.nuget.org/packages/IronXL.Excel) or by directly implementing the DLL file. Then, engage with documents using the following:
- Load documents with `WorkBook.Load` supporting XLS, XLSX, and CSV formats.
- Access cell values with concise syntax, e.g., `sheet["A11"].DecimalValue`.

This guide covers:
- **Setting up IronXL.Excel:** Adding the IronXL library to your application.
- **Basic Operations:** From opening workbooks to reading and writing cells.
- **Enhanced Sheet Functionalities:** Dive into advanced features like adding headers and performing complex calculations.

#### Opening an Excel File: Quick Start Example

```csharp
using IronXL;

WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;
Range range = worksheet["A2:A8"];
decimal total = 0;

// Loop through cells in the range
foreach (var cell in range)
{
    Console.WriteLine($"Cell {cell.RowIndex} has value '{cell.Value}'");
    if (cell.IsNumeric)
    {
        // Add up numeric values
        total += cell.DecimalValue;
    }
}

// Verify the correctness of calculations
if (worksheet["A11"].DecimalValue == total)
{
    Console.WriteLine("Validation successful");
}
```

#### Writing to and Saving an Excel File: Quick Example

```csharp
worksheet["B1"].Value = 11.54;

// Commit changes to the file
workbook.SaveAs("test.xlsx");
```

---

### Step-by-Step Implementation

## Installation of IronXL C# Library for Free

IronXL.Excel is a robust and flexible library that can be easily implemented across various .NET project types, including desktop applications, ASP.NET MVC, and .NET Core applications.

#### Adding IronXL via NuGet Package Manager

1. Install the library using the NuGet Package Manager which provides a graphical interface:
   - Right-click the project name.
   - Click 'Manage NuGet Packages'.
   - Search for "IronXL.Excel" and click 'Install'.

![Manage NuGet Packages](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg)

2. Alternatively, use the Package Manager Console:
   - Navigate to 'Tools' -> 'NuGet Package Manager' -> 'Package Manager Console'.
   - Run the command `Install-Package IronXL.Excel -Version 2019.5.2`.

![Package Manager Console](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg)

#### Manual DLL Integration

For manual setup, the DLL can be directly added to your project or the global assembly cache. To install using the DLL, download it from [here](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

```plaintext
PM > Install-Package IronXL.Excel
```

Continue to learn more about basic and advanced operations by following the steps outlined in this tutorial, or explore additional resources provided by Iron Software.

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## Basic Excel Operations: Creation, Opening, and Saving with IronXL ##

### Getting Started with a HelloWorld Project ###

<p class="list-description">Starting a HelloWorld Project</p>

<p class="list-decimal">2.1.1. Launch Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Opt to Create a New Project</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Select Console App (.NET framework) as the project type</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Name your sample as “HelloWorld” and launch it</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. Visual Studio has now created your console application</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Include IronXL.Excel by clicking install</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Initialize by reading the first cell of the first sheet and displaying it</p>

```csharp
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
string cell = workSheet["A1"].StringValue;
Console.WriteLine(cell);
```

### Creating a New Excel File ###

<p class="list-description">Initiate a fresh Excel file using IronXL</p>

```csharp
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "IronXL New File";
WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
```

### Open (CSV, XML, JSON List) as Workbook ###

<p class="list-decimal">2.3.1. Open a CSV file</p>

<p class="list-decimal">2.3.2 Prepare a new text document, enter a list of names and ages, then save it as CSVList.csv</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p>Your code should resemble this:</p>

```csharp
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
WorkSheet workSheet = workBook.WorkSheets.First();
string cell = workSheet["A1"].StringValue;
Console.WriteLine(cell);
```

<p class="list-decimal">
    2.3.3. Load an XML File
    <span class="list-description">Design an XML file containing a list of countries, with each country holding attributes defining it such as code and continent.</span>
</p>

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">2.3.4. Apply the subsequent piece of code to manage the XML as a workbook</p>

```csharp
DataSet xmldataset = new DataSet();
xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();
```

<p class="list-decimal">2.3.5. Open a JSON list of countries</p>

```json
[
    {
        "name": "United Arab Emirates",
        "code": "AE"
    },
    {
        "name": "United Kingdom",
        "code": "GB"
    },
    {
        "name": "United States",
        "code": "US"
    },
    {
        "name": "United States Minor Outlying Islands",
        "code": "UM"
    }
]
```

### 2.1. Example Project: HelloWorld Console Project ###

<p class="details-list">Initiating a HelloWorld Project</p>

<p class="sequence-step">2.1.1. Launch Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="image-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="sequence-step">2.1.2. Opt for Create New Project</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="image-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="sequence-step">2.1.3. Select Console App (.NET framework)</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="image-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="sequence-step">2.1.4. Name the sample “HelloWorld” and confirm creation</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="image-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="sequence-step">2.1.5. Console application is now created</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="image-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="sequence-step">2.1.6. Incorporate IronXL.Excel => initiate install</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="image-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="sequence-step">2.1.7. Introduce your initial code to read the first cell in the first sheet of the Excel file, then display it</p>

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
string cell = workSheet["A1"].StringValue;
Console.WriteLine(cell);
```

<p class="list-description">Create a HelloWorld Project</p>

<p class="list-decimal">2.1.1. Open Visual Studio</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Choose Create New Project</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Choose Console App (.NET framework)</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Give our sample the name “HelloWorld” and click create</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. Now we have console application created</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Add IronXL.Excel => click install</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Add our first few lines that reads 1st cell in 1st sheet in the Excel file, and print</p>

Here is the paraphrased section of the article:

```cs
// Load the workbook from the current directory
WorkBook workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\HelloWorld.xlsx");
// Get the first worksheet
WorkSheet worksheet = workbook.GetFirstSheet();
// Retrieve the value of the first cell
string value = worksheet["A1"].StringValue;
// Output the value to the console
Console.WriteLine(value);
```

### 2.2. Generating a New Excel Document ###

Craft a fresh Excel file effortlessly using IronXL:

```cs
// Instantiate a new WorkBook
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
// Set the metadata title of the workbook
workBook.Metadata.Title = "IronXL New File";

// Create a worksheet titled '1stWorkSheet'
WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");

// Assigning value to cell A1
workSheet["A1"].Value = "Hello World";

// Customize the style of cell A2
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
``` 

This snippet sets up a new Excel workbook with IronXL, establishes a worksheet, and styles the cells with just a few lines of code.

<p class="list-description">Create a new Excel file using IronXL</p>

Here is the paraphrased content:

```cs
// Initialize a new workbook with XLSX format
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
// Set the title in the metadata of the workbook
workbook.Metadata.Title = "IronXL New File";

// Create a new worksheet and name it
WorkSheet worksheet = workbook.CreateWorkSheet("FirstSheet");

// Set the value of cell A1
worksheet["A1"].Value = "Hello World";

// Customize the style of cell A2
worksheet["A2"].Style.BottomBorder.SetColor("#ff6600");  // Set the color of the bottom border
worksheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;  // Set the border type to dashed
```

### 2.3. Workbook Creation from CSV, XML, and JSON

This section focuses on how to load workbooks from different file types such as CSV, XML, and JSON using IronXL.

#### 2.3.1. Opening a CSV File
Load existing CSV files to start your workbook.

#### 2.3.2. Create a CSV File 
First, you'll create a text file containing a list of names and ages, and save it with a `.csv` extension. Here's an example of how your data should be organized in the file:

[![CSV Data](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg)](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg)

#### Code Example for Loading a CSV File
Here's how you can load the CSV file into a workbook:
```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
WorkSheet workSheet = workBook.WorkSheets.First();
string cell = workSheet["A1"].StringValue;
Console.WriteLine(cell);
```

#### 2.3.3. Open an XML File
Create and define an XML file that lists countries, each with attributes such as code, continent, etc. Here's an example structure for the XML file:
```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <country code="ae" continent="asia" handle="united-arab-emirates" iso="784">United Arab Emirates</country>
  <country code="gb" continent="europe" handle="united-kingdom" iso="826">United Kingdom</country>
  <!-- Add more countries here -->
</countries>
```

#### Code to Load XML as a Workbook
Once your XML is ready, use this snippet to read it into a workbook:
```cs
DataSet xmldataset = new DataSet();
xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();
```

#### 2.3.4. Open a JSON List as Workbook
You can also initialize workbooks with JSON data. Below is a JSON structure for a list of countries:
```json
[
    {"name": "United Arab Emirates", "code": "AE"},
    {"name": "United Kingdom", "code": "GB"},
    {"name": "United States", "code": "US"},
    {"name": "United States Minor Outlying Islands", "code": "UM"}
]
```

#### Define a Model to Match the JSON Structure
First, create a simple C# model that matches the JSON data:
```cs
public class CountryModel {
    public string name { get; set; }
    public string code { get; set; }
}
```

#### Convert JSON to DataSet and Load as Workbook
Convert the JSON into a DataSet and then load it into the workbook:
```cs
StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
var xmldataset = countryList.ToDataSet();
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();
```

This process allows you to flexibly import data from several file formats directly into IronXL, enabling advanced data manipulation and analysis within your .NET applications.

<p class="list-decimal">2.3.1. Open CSV file</p>

<p class="list-decimal">2.3.2 Create a new text file and add to it a list of names and ages (see example) then save it as CSVList.csv</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Your code snippet should look like this</p>

Here's the rewritten version of the provided code snippet:

```cs
// Load the workbook from a CSV file located in the current directory's "Files" subfolder
WorkBook currentWorkbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\CSVList.csv");

// Retrieve the first worksheet within our loaded workbook
WorkSheet primarySheet = currentWorkbook.WorkSheets.First();

// Access the value of the first cell in the first worksheet and assign it to a string variable
string firstCellValue = primarySheet["A1"].StringValue;

// Output the value of the first cell to the console
Console.WriteLine(firstCellValue);
```

<p class="list-decimal">

### 2.3.3. Open an XML File

To begin working with XML files in your C# application, first, you need to create an XML file. Here's a sample structure for your XML file which includes a list of countries. Each country is represented with details like its code, name, and attributes related to its geographic and political information.

```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

Next, you can load this XML data directly into an IronXL `WorkBook` by reading the XML data into a `DataSet` and then initializing the `WorkBook` using this dataset. Here's how to execute this:

```csharp
DataSet xmldataset = new DataSet();
xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();
```

This technique allows you to manipulate and work with XML data as if it were a typical Excel file, taking full advantage of IronXL's features to handle complex data transformation and analysis right within your .NET applications.

<span class="list-description">Create an XML file that contains a countries list: the root element “countries”, with children elements “country”, and each country has properties that define the country like code, continent, etc.</span>
</p>

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">2.3.4. Copy the following code snippet to open XML as a workbook</p>

```cs
// Initialize a new DataSet instance.
DataSet xmlDataSet = new DataSet();
// Load the XML file into the DataSet from the specified directory.
xmlDataSet.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
// Create a WorkBook instance by loading data from the DataSet.
WorkBook excelWorkbook = IronXL.WorkBook.Load(xmlDataSet);
// Retrieve the first WorkSheet from the WorkBook.
WorkSheet excelSheet = excelWorkbook.WorkSheets.First();
```

<p class="list-decimal">

### 2.3.5. Load JSON List as a Workbook

This step explains how to create and import a JSON list representing a collection of countries into an Excel workbook using IronXL.

- First, you will need to create a JSON file that includes a list of countries. Here's a sample JSON structure you can use:

```json
[
    {
        "name": "United Arab Emirates",
        "code": "AE"
    },
    {
        "name": "United Kingdom",
        "code": "GB"
    },
    {
        "name": "United States",
        "code": "US"
    },
    {
        "name": "United States Minor Outlying Islands",
        "code": "UM"
    }
]
```

- You must define a model in C# that matches the JSON structure. This class will map the JSON properties to C# object properties. Here’s an example of the country model class:

```csharp
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

- To handle the JSON data, you'll need to include the Newtonsoft.Json library in your project. This can be installed via NuGet. This library helps in deserializing the JSON data into C# objects.

- Convert this list of country model objects into a `DataSet`. To accomplish this, create a utility class that extends list functionality:

```csharp
public static class ListConvertExtension
{
    public static DataSet ToDataSet<T>(this IList<T> list)
    {
        Type elementType = typeof(T);
        DataSet ds = new DataSet();
        DataTable t = new DataTable();
        ds.Tables.Add(t);

        // Add a column in the DataTable for each public property on T
        foreach (var propInfo in elementType.GetProperties())
        {
            Type colType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
            t.Columns.Add(propInfo.Name, colType);
        }

        // Populate the DataTable with values from the list
        foreach (T item in list)
        {
            DataRow row = t.NewRow();
            foreach (var propInfo in elementType.GetProperties())
            {
                row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
            }
            t.Rows.Add(row);
        }
        return ds;
    }
}
```

- Finally, load the `DataSet` into an IronXL `WorkBook` and access it through a `WorkSheet`:

```csharp
using (StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json"))
{
    var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
    var dataset = countryList.ToDataSet();
    WorkBook workbook = IronXL.WorkBook.Load(dataset);
    WorkSheet worksheet = workbook.WorkSheets.First();
}
```

In this way, you can easily transform JSON data into an Excel workbook using IronXL, providing a straightforward method for managing structured data within .NET applications.

<span class="list-description">Create JSON country list</span>
</p>

Here's a paraphrased version of the given section:

```cs
[
    {
        "name": "United Arab Emirates",
        "isoCode": "AE"
    },
    {
        "name": "United Kingdom",
        "isoCode": "GB"
    },
    {
        "name": "United States",
        "isoCode": "US"
    },
    {
        "name": "United States Minor Outlying Islands",
        "isoCode": "UM"
    }
]
```

<p class="list-decimal"></p>
<p class="list-decimal">2.3.6. Create a country model that will map to JSON</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Here is the class code snippet</p>

Here's the paraphrased section of the code with the URL paths resolved to `ironsoftware.com`:

```cs
public class NationDetails
{
    public string Name { get; set; }
    public string CountryCode { get; set; }
}
```

<p class="list-decimal"></p>
<p class="list-decimal">2.3.8. Add Newtonsoft library to convert JSON to the list of country models</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/add-newtonsoft-library-to-convert-json.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/add-newtonsoft-library-to-convert-json.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">2.3.9 To convert the list to dataset, we have to create a new extension for the list. Add extension class with the name “ListConvertExtension”</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/convert-list-to-dataset.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/convert-list-to-dataset.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Then add this code snippet</p>

Here's the paraphrased section of the code, with comments enhanced for better understanding and clarity:

```cs
public static class ListConvertExtension
{
    // Method to convert a list of type T to a DataSet
    public static DataSet ToDataSet<T>(this IList<T> list)
    {
        // Get the type of the elements in the list
        Type elementType = typeof(T);
        // Create a new DataSet
        DataSet ds = new DataSet();
        // Create a new DataTable
        DataTable t = new DataTable();
        // Add the DataTable to the DataSet
        ds.Tables.Add(t);

        // Add a column to the DataTable for each public property of type T
        foreach (var propInfo in elementType.GetProperties())
        {
            // Determine the type of the column
            Type colType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
            // Add the column to the DataTable
            t.Columns.Add(propInfo.Name, colType);
        }

        // Populate the DataTable with values from the list
        foreach (T item in list)
        {
            // Create a new row in the DataTable
            DataRow row = t.NewRow();
            // Set each column value to the value of the corresponding property in the object
            foreach (var propInfo in elementType.GetProperties())
            {
                // If the property value is null, use DBNull.Value
                row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
            }
            // Add the completed row to the DataTable
            t.Rows.Add(row);
        }
        // Return the filled DataSet
        return ds;
    }
}
``` 

This rewritten version maintains the functionality while enhancing the clarity and detail of the comments to make it easier for other developers to understand and use this code segment.

<p class="list-decimal"></p>
<p class="list-decimal">And finally load this dataset as a workbook</p>

Here's the paraphrased section with resolved URL paths:

```cs
StreamReader countriesJsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
var listOfCountries = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(countriesJsonFile.ReadToEnd());
var dataSetOfCountries = listOfCountries.ToDataSet();
WorkBook workbookInstance = IronXL.WorkBook.Load(dataSetOfCountries);
WorkSheet initialWorkSheet = workbookInstance.WorkSheets.First();
```

### Saving and Exporting Files ###

This section demonstrates how to store or convert your Excel document to several formats such as XLSX, CSV, JSON, and XML using IronXL. Here are specific examples to help you.

#### 2.4.1. Save as XLSX ####

To save as an `.xlsx` file, utilize the `SaveAs` function. This example involves creating an Excel workbook and saving it in the XLSX format.

```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "IronXL New File";

WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
```

#### 2.4.2. Save as CSV ####

To save your workbook as a `.csv` file, you can use the `SaveAsCsv` method and specify the delimiter.

```cs
workBook.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter: "|");
```

#### 2.4.3. Save as JSON ####

Saving the workbook in JSON format can be done using the `SaveAsJson` method. This makes your data readable in a wide range of applications that support JSON.

```cs
workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

Here is how your JSON output will appear:

```json
[
    [
        "Hello World"
    ],
    [
        ""
    ]
]
```

#### 2.4.4. Save as XML ####

Finally, to store your data as `.xml`, use the `SaveAsXml` method. This format is beneficial for configurations or data sharing across different platforms.

```cs
workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
```

This is what your XML output will look like:

```xml
<?xml version="1.0" standalone="yes"?>
<_x0031_stWorkSheet>
  <_x0031_stWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Column1>
  </_x0031_stWorkSheet>
  <_x0031_stWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </_x0031_stWorkSheet>
</_x0031_stWorkSheet>
```

These examples highlight the versatility of the IronXL library in handling various file formats, making your workflow smoother and more efficient.

<span class="list-description">We can save or export the Excel file to multiple file formats like (“.xlsx”,”.csv”,”.html”) using one of the following commands.</span>

<p class="list-decimal">

### 2.4.1. Saving to XLSX Format

Utilize the following command to store the Excel document in the XLSX format. Here’s how you can do it:

```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "IronXL New File";

WorkSheet workSheet = workBook.CreateWorkSheet("FirstSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
```

This example illustrates the creation of a new workbook with a single worksheet. It demonstrates setting the content of cell `A1` to "Hello World" and applying a dashed bottom border with a specific color to cell `A2`. Finally, it saves the workbook in the `.xlsx` format to the specified file path.

<span class="list-description">To Save to “.xlsx” use saveAs function</span>
</p>

Here's the paraphrased section of the code:

```cs
// Instantiate a new Workbook of Excel format
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
workbook.Metadata.Title = "IronXL New File";

// Add a new Worksheet named 'FirstSheet'
WorkSheet worksheet = workbook.CreateWorkSheet("FirstSheet");
worksheet["A1"].Value = "Hello World";
worksheet["A2"].Style.BottomBorder.SetColor("#ff6600"); // Setting the color of the bottom border
worksheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed; // Style set as Dashed

// Save the workbook in the current directory
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
```

This code snippet demonstrates initializing a new Excel workbook using the `IronXL` library, creating a worksheet, setting values and styles for cells, and then saving the workbook into the specified directory.

<p class="list-decimal">

### 2.4.2 Save to CSV Format

<span class="list-description">Utilize the `SaveAsCsv` method to store the workbook in CSV format. This function requires two arguments: the first is the file's path and name, and the second is the delimiter character, which can be a comma (`,`), a pipe (`|`), or a colon (`:`).</span>
```

</p>

```cs
// Save the workbook as a CSV file in the current directory with a custom delimiter
workBook.SaveCsvFile(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "HelloWorld.csv"), delimiter: ",");
```

<p class="list-decimal">

### 2.4.3. Export as JSON Format

<span class="list-description">For converting an Excel sheet to the JSON format:</span>

```cs
workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

<span class="list-description">Here’s what the JSON output file will look like:</span>

```json
[
    [
        "Hello World"
    ],
    [
        ""
    ]
]
```

<span class="list-description">To save to Json “.json” use SaveAsJson as follow</span>
</p>

Here's a paraphrased version of the suggested section while resolving the relative URL paths:

-----
```cs
// Save the workbook to a JSON file in the current directory
workBook.SaveAsJson($"{Environment.CurrentDirectory}/Files/HelloWorldJSON.json");
```

<p class="list-decimal">
  <span class="list-description">The result file should look like this</span>
</p>

Certainly! Here's the paraphrased section you requested:

```cs
[
    [
        "Hello World"
    ],
    [
        // Empty cell representation
    ]
]
```

<p class="list-decimal">

### Saving to XML Files using IronXL

IronXL provides a seamless method to export Excel sheets directly to XML format. Below, we detail the process of saving an Excel workbook as an XML file using IronXL:

```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "IronXL New File";

WorkSheet workSheet = workBook.CreateWorkSheet("FirstWorkSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

// Using the SaveAsXml method to save the workbook as an XML file
workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.xml");
```

The output XML will be structured as follows, showing a simple example of exported data from an Excel file into XML format:

```xml
<?xml version="1.0" standalone="yes"?>
<_x0031_stWorkSheet>
  <_x0031_stWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Column1>
  </_x0031_stWorkSheet>
  <_x0031_stWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </_x0031_stWorkSheet>
</_x0031_stWorkSheet>
```

This approach allows developers to take advantage of IronXL's capabilities to create, modify, and save Excel data as XML, which is beneficial for data sharing and interoperability with other applications that process XML formats.

<span class="list-description">To save to xml use SaveAsXml as follow</span>
</p>

Certainly! Here's a paraphrased version of the specified section with updated paths:

```cs
// This command saves the current workbook into an XML file within the 'Files' directory.
workBook.SaveAsXml($"{Directory.GetCurrentDirectory()}/Files/HelloWorldXML.XML");
```

<p class="list-decimal">
  <span class="list-description">Result should be like this</span>
</p>

The following is a paraphrased version of the XML section provided:

```html
<?xml version="1.0" encoding="utf-8"?>
<FirstWorkSheet>
  <FirstWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Column1>
  </FirstWorkSheet>
  <FirstWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></Column1>
  </FirstWorkSheet>
</FirstWorkSheet>
```

<hr class="separator">

## 3. Advanced Excel Functions: Sum, Average, Count, and More ##

Explore the utilization of popular Excel functionalities such as SUM, AVERAGE, and COUNT through practical code examples. Let's dive into them one by one.

### Example of Calculating Sum ###

In this example, we demonstrate how to compute the sum of a list of numbers using IronXL. I previously created an Excel document named "Sum.xlsx" and populated it with a series of numerical values.

#### Original Excel Visualization ####
For a visual representation of our file setup, click on the following link:
![Sum Example](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png)

#### Code to Calculate Sum ####

Below is the code snippet that loads the "Sum.xlsx" file and computes the sum of numbers from cells `A2` to `A4`.

```cs
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet worksheet = workbook.WorkSheets.First();
decimal totalSum = worksheet["A2:A4"].Sum();
Console.WriteLine(totalSum);
```
This will output the sum of the numbers in the specified range to the console.

<p class="list-description">Let’s find the sum for this list. I created an Excel file and named it “Sum.xlsx” and added this list of numbers manually</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

Here is the paraphrased section of the code example provided:

```cs
// Load the workbook containing sum data from the specified directory
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the loaded workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Calculate the sum of the values in cells A2 through A4
decimal totalSum = worksheet["A2:A4"].Sum();

// Output the calculated sum to the console
Console.WriteLine(totalSum);
```

### 3.2. Example: Calculating the Average ###

In this section, we will demonstrate how to calculate the average value from a list of numbers using IronXL. To facilitate this, an Excel file named "Sum.xlsx" has been prepopulated with numerical data.

```cs
// Load the workbook from the specified location
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
WorkSheet workSheet = workBook.WorkSheets.First();

// Calculate the average of values in the range A2 to A4
decimal avg = workSheet["A2:A4"].Avg();

// Output the average value to the console
Console.WriteLine(avg);
```

<p class="list-description">Using the same file, we can get the average:</p>

Here's the paraphrased version of the provided code snippet:

```cs
// Load the Excel workbook
WorkBook excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
WorkSheet mainSheet = excelWorkbook.WorkSheets.First();

// Calculate the average value of the cells from A2 to A4
decimal averageValue = mainSheet["A2:A4"].Avg();

// Output the average value to the console
Console.WriteLine(averageValue);
```

### 3.3. Example: Counting Elements ###

This example demonstrates how to determine the number of elements in a specific range of cells within an Excel file using IronXL:

```cs
// Load the workbook
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet
WorkSheet workSheet = workBook.WorkSheets.First();

// Count the number of elements in the range A2 to A4
decimal elementCount = workSheet["A2:A4"].Count();

// Output the count to the console
Console.WriteLine(elementCount);
```

This code snippet efficiently counts the elements between cells A2 and A4 in the `Sum.xlsx` Excel file, displaying the result in the console. This operation is useful for quickly assessing the quantity of data entries in a specified range.

<p class="list-description">Using the same file, we can also get the number of elements in a sequence:</p>

```cs
// Load the workbook from a specific file
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the loaded workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Use the Count method to count the elements in the specified cell range
decimal elementCount = worksheet["A2:A4"].Count();

// Output the count to the console
Console.WriteLine(elementCount);
```

### 3.4. Example: Finding the Maximum Value ###

This section demonstrates how to identify the maximum value within a range of cells using IronXL. We'll utilize an Excel file named "Sum.xlsx" and focus specifically on the cells from A2 to A4.

Here’s how to accomplish this:

```cs
// Load the workbook and select the first worksheet
WorkBook workBook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Calculate the maximum value from the specified cell range
decimal max = workSheet["A2:A4"].Max();

// Display the maximum value in the output
Console.WriteLine(max);
```

In addition, you can apply functions to filter or alter the output during the max value retrieval:

```cs
// Load the workbook and select the first worksheet
WorkBook workBook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Check if any cell within the range contains a formula and retrieve the maximum value based on a condition
bool hasFormula = workSheet["A1:A4"].Max(cell => cell.IsFormula);

// Display the result in the console
Console.WriteLine(hasFormula);  // Outputs 'false' if no formulas are detected
```

These examples not only show how to retrieve the highest numerical value from a selected range but also demonstrate conditional checks within the range values in IronXL.

<p class="list-description">Using the same file, we can get the max value of range of cells:</p>

Here's the paraphrased section with URLs resolved to `ironsoftware.com`:

```cs
// Load the Workbook from a specified path
WorkBook excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Retrieve the first Worksheet from the Workbook
WorkSheet activeSheet = excelWorkbook.WorkSheets.First();

// Calculate the maximum value from cell range A2 to A4
decimal maximumValue = activeSheet["A2:A4"].Max();

// Display the maximum value on the console
Console.WriteLine(maximumValue);
```

<p class="list-description">– We can apply the transform function to the result of max function:</p>

Here is a paraphrased version of the provided code snippet, including enhanced comments for clearer understanding:

```cs
// Load the workbook from a file named 'Sum.xlsx' located in the Files directory of the current working directory.
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Retrieve the first worksheet from the workbook.
WorkSheet worksheet = workbook.WorkSheets.First();

// Evaluate the maximum condition on a cell range, checking if any formula is present in the cells from A1 to A4.
bool hasFormula = worksheet["A1:A4"].Max(cell => cell.IsFormula);

// Output the result to the console. It prints 'true' if any cell in the range contains a formula, otherwise 'false'.
Console.WriteLine(hasFormula);
```

This updated snippet provides more explanations through comments and uses variable naming that enhances readability, while maintaining the functionality of the original code.

<p class="list-description">This example writes “false” in the console.</p>

### 3.5 Example of Calculating Minimum ###

Below, we explore how to determine the minimum value within a specific range of cells in an existing Excel spreadsheet titled "Sum.xlsx".

```cs
// Load the workbook from a file
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
// Get the first worksheet in the workbook
WorkSheet workSheet = workBook.WorkSheets.First();
// Calculate the minimum value from a specified cell range
decimal min = workSheet["A1:A4"].Min();
// Output the minimum value to the console
Console.WriteLine(min);
```

<p class="list-description">Using the same file, we can get the min value of range of cells:</p>

Here's the paraphrased code snippet:

```cs
// Load the workbook using the file path
WorkBook excelWorkbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\Sum.xlsx");
// Access the first worksheet from the workbook
WorkSheet excelSheet = excelWorkbook.WorkSheets.First();
// Calculate the minimum value from the specified cell range
decimal minimumValue = excelSheet["A1:A4"].Min();
// Output the minimum value to the console
Console.WriteLine(minimumValue);
```

### 3.6. Example: Ordering Cells ###

<p class="list-description">Taking the same file "Sum.xlsx", let's explore how we can organize cells in either ascending or descending order:</p>

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
// For ascending order:
workSheet["A1:A4"].SortAscending();
// Uncomment the following line to sort in descending order:
// workSheet["A1:A4"].SortDescending();
// Save the sorted Excel file
workBook.SaveAs("SortedSheet.xlsx");
```

<p class="list-description">Using the same file, we can order cells by ascending or descending:</p>

Here's the paraphrased section of the article, with relative URLs resolved:

```cs
// Load the workbook from a specified path
WorkBook excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet from the workbook
WorkSheet excelSheet = excelWorkbook.WorkSheets.First();

// Sort the range of cells from A1 to A4 in ascending order
excelSheet["A1:A4"].SortAscending();
// Uncomment the next line to sort in descending order instead
// excelSheet["A1:A4"].SortDescending();

// Save the workbook to a new file
excelWorkbook.SaveAs("SortedSheet.xlsx");
```

### 3.7. Example: Using IF Conditions in Formulas ###

In this section, we will demonstrate how to incorporate IF conditions into cell formulas to evaluate specific criteria. Here's how to apply IF statements to manage data in your Excel workbook using IronXL:

1. **Setting Up the Workbook**: Start by loading an existing workbook or creating a new one. For this example, we will use a previously created file named `Sum.xlsx`.

```cs
WorkBook workbook = WorkBook.Load(@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet worksheet = workbook.WorkSheets.First();
```

2. **Applying IF Conditions**: Navigate through specific cells and assign conditions that will return "Pass" if the value is greater than or equal to 20, and "Fail" otherwise.

```cs
int rowIndex = 1;
foreach (var cell in worksheet["B1:B4"])
{
    cell.Formula = $"=IF(A{rowIndex}>=20, \"Pass\", \"Fail\")";
    rowIndex++;
}
```

3. **Save the Changes**: Once all formulas are set, save the workbook to preserve the changes, specifying the desired location and file name.

```cs
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
```

4. **Retrieve Formulas**: After saving, if you wish to check the formulas applied in the cells, you can reload the workbook and output the formulas to the console:

```cs
WorkBook loadedWorkbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
WorkSheet loadedWorksheet = loadedWorkbook.WorkSheets.First();
foreach (var cell in loadedWorksheet["B1:B4"])
{
    Console.WriteLine(cell.Formula);
}
Console.ReadKey();
```

By following these steps, you can efficiently utilize conditional logic to manage and manipulate data in Excel files using the IronXL library, all without needing Excel installed on your machine.

<p class="list-description">Using the same file, we can use the Formula property to set or get a cell’s formula:</p>

<p class="list-decimal">3.7.1. Save to XML “.xml”</p>

Here is the paraphrased section of the article involving IronXL usage for working with Excel formulas and saving the file:

```cs
// Load an existing workbook
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
// Access the first worksheet
WorkSheet worksheet = workbook.WorkSheets.First();
int index = 1;
// Loop through cells from B1 to B4
foreach (var cell in worksheet["B1:B4"])
{
    // Assign conditional formula to each cell
    cell.Formula = $"=IF(A{index} >= 20, \" Pass\", \" Fail\")";
    index++;
}
// Save the changes to a new Excel file
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
```

<p class="list-decimal">7.2. Using the generated file from the previous example, we can get the Cell’s Formula:</p>

Here's the paraphrased content for the given C# code section:

```cs
// Load the workbook from the directory where the application is running
WorkBook workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\NewExcelFile.xlsx");

// Get the first worksheet from the workbook
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Iterate over cells from B1 to B4
foreach (var cell in worksheet["B1:B4"])
{
    // Output the formula within each cell to the console
    Console.WriteLine(cell.Formula);
}

// Wait for a user key press before closing the console window
Console.ReadKey();
```

### 3.8. Example: Trimming Cells ###

In this section, we demonstrate how to use the `trim` function to remove extraneous spaces from cell values. This is particularly useful for ensuring that data is neatly formatted before carrying out operations such as data analysis or reporting.

Below is a practical guide:

1. **Visual Example**: First, we incorporated a new column in our `sum.xlsx` file specifically to illustrate the impact of the `trim` function.
   
   ![A visual depiction of cell values before and after applying the trim function](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png)

2. **Implementation Code**: 
   To apply the `trim` function, we added the following code to manipulate cells within the `NewExcelFile.xlsx`:

   ```csharp
   WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
   WorkSheet workSheet = workBook.WorkSheets.First();
   int i = 1;
   foreach (var cell in workSheet["f1:f4"])
   {
       cell.Formula = "=trim(D" + i + ")";
       i++;
   }
   workBook.SaveAs("editedFile.xlsx");
   ```

This demonstrates how you can effectively eliminate unnecessary spaces from data entries. It’s a common technique used to standardize data before analysis or reporting.

<p class="list-description">To apply trim function (to eliminate all extra spaces in cells), I added this column to the sum.xlsx file</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description">And use this code</p>

Here's the paraphrased section of the code you provided, ensuring references to images and links are resolved to `ironsoftware.com`:

```cs
// Load the workbook from the specified file
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");

// Access the first worksheet in the loaded workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Initialize a counter
int counter = 1;

// Iterate through a specific range of cells
foreach (var cell in worksheet["f1:f4"])
{
    // Set the formula to trim the content of another cell
    cell.Formula = $"=trim(D{counter})";
    counter++; // Increment the counter
}

// Save the modifications to a new Excel file
workbook.SaveAs("editedFile.xlsx");
```

<p class="list-description">Thus, you can apply formulas in the same way.</p>

<hr class="separator">

## Multisheet Workbooks in C#

In this section, we explore how to manage Excel workbooks containing multiple sheets using IronXL.

### 4.1 Reading Data from Various Sheets within a Workbook

For Excel files that consist of multiple sheets, such as "Sheet1" and "Sheet2," you have the flexibility to choose which sheet to work with, instead of defaulting to the first sheet. Here's an example of how to specify and handle data from different sheets:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
var range = workSheet["A2:D2"];
foreach (var cell in range)
{
    Console.WriteLine(cell.Text);
}
```

### 4.2 Adding a New Sheet to an Existing Workbook

IronXL also allows you to easily add new sheets to an already existing workbook. This can be particularly useful for organizing different types of data logically separated within the same file. Here's how you can add a new sheet and introduce some data:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet newSheet = workBook.CreateWorkSheet("new_sheet");
newSheet["A1"].Value = "Hello World";
workBook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

By mastering these functionalities, you can enhance how you manipulate multi-sheet Excel workbooks, making your data processing tasks more structured and efficient.

### 4.1 Accessing Data Across Multiple Sheets in a Workbook ###

Explore how to read data from several sheets within the same workbook effectively.

```cs
// Load the workbook from a specified file
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
// Access a specific worksheet by its name
WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
// Define a range of cells to read
var range = workSheet["A2:D2"];
// Iterate through each cell in the range and output their text content
foreach (var cell in range)
{
    Console.WriteLine(cell.Text);
}
```

This snippet demonstrates the process of loading a workbook, selecting a particular sheet, and then reading and displaying a specific range of cells.

<p class="list-description">I created an xlsx file that contains two sheets: “Sheet1”,” Sheet2”</p>
<p class="list-description">Until now we used WorkSheets.First() to work with the first sheet. In this example we will specify the sheet name and work with it</p>

```cs
// Load the workbook from the current directory
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Access the worksheet named 'Sheet2'
WorkSheet worksheet = workbook.GetWorkSheet("Sheet2");

// Define a range of cells to be read; in this case from A2 to D2
IronXL.Range cellRange = worksheet["A2:D2"];

// Iterate through the cells in the defined range
foreach (var cell in cellRange)
{
    // Output the text content of each cell to the Console
    Console.WriteLine(cell.Text);
}
```

### 4.2. Insert a New Worksheet into an Existing Workbook ###

This tutorial section demonstrates how to enhance a workbook by adding a new worksheet to it.

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet newSheet = workBook.CreateWorkSheet("new_sheet");
newSheet["A1"].Value = "Hello World";
workBook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

Include this straightforward script in your .NET project to seamlessly expand your workbook capabilities by incorporating additional sheets. Whether for accumulating data sets or segregating different data forms, this method provides a robust foundation for managing diverse content within a single workbook.

<p class="list-description">We can also add new sheet to a workbook:</p>

Here's the paraphrased section with the relative URL paths resolved:

```cs
// Load an existing workbook
WorkBook existingWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create a new worksheet named 'new_sheet'
WorkSheet newWorksheet = existingWorkbook.CreateWorkSheet("new_sheet");

// Assign a value to cell A1 in the new worksheet
newWorksheet["A1"].Value = "Hello World";

// Save the workbook with the new worksheet to a specified path
existingWorkbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

<hr class="separator">

## 5. Database Integration with Excel ##

Explore how to import and export data between a database and Excel.

I established a database called "TestDb" which has a table named Country. This table contains two columns: Id (integer, identity) and CountryName (string).

### 5.1. Populate an Excel Sheet with Database Data ###

This segment demonstrates how to populate an Excel sheet with data extracted from a database.

```cs
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.CreateWorkSheet("FromDb");
List<Country> countryList = dbContext.Countries.ToList();
sheet.SetCellValue(0, 0, "Id");
sheet.SetCellValue(0, 1, "Country Name");
int row = 1;
foreach (var item in countryList)
{
    sheet.SetCellValue(row, 0, item.id);
    sheet.SetCellValue(row, 1, item.CountryName);
    row++;
}
workbook.SaveAs("FilledFile.xlsx");
```

This example illustrates creating a new worksheet titled "FromDb" in an existing workbook and filling it with data from a 'Countries' table in a database. The code utilizes the `IronXL.WorkBook` class to load the Excel file and `WorkSheet.SetCellValue` method to insert database data into the sheet. Each row corresponds to a record in the database, populating columns "Id" and "Country Name" with data. The updated workbook is then saved with the name "FilledFile.xlsx", preserving the changes.

<p class="list-description">Here we will create a new sheet and fill it with data from the Country Table</p>

Here's a paraphrased version of the section provided, with absolute URLs resolved using `ironsoftware.com` and slight modifications to the code for a clearer understanding:

```cs
// Initialize the database context for accessing country data
TestDbEntities databaseContext = new TestDbEntities();

// Load an existing workbook or create one if it doesn't exist
WorkBook excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create a new worksheet titled 'FromDb' in the workbook
WorkSheet dataSheet = excelWorkbook.CreateWorkSheet("FromDb");

// Retrieve the list of countries from the database
List<Country> listofCountries = databaseContext.Countries.ToList();

// Setup column headers for the new worksheet
dataSheet.SetCellValue(0, 0, "Id");
dataSheet.SetCellValue(0, 1, "Country Name");

// Starting from the second row, populate the worksheet with country data
int currentRow = 1;
foreach (Country country in listofCountries)
{
    dataSheet.SetCellValue(currentRow, 0, country.id);
    dataSheet.SetCellValue(currentRow, 1, country.CountryName);
    currentRow++;
}

// Save the populated workbook as a new file named 'FilledFile.xlsx'
excelWorkbook.SaveAs("FilledFile.xlsx");
```

### 5.2. Populate Database with Data from an Excel Sheet ###

<p class="list-description">Insert data into the TestDb Database's Country table</p>

```cs
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
System.Data.DataTable dataTable = sheet.ToDataTable(true);
foreach (DataRow row in dataTable.Rows)
{
    Country c = new Country();
    c.CountryName = row[1].ToString();
    dbContext.Countries.Add(c);
}
dbContext.SaveChanges();
```

<p class="list-description">Insert the data to the Country table in TestDb Database</p>

The following code snippet loads an Excel file and transfers the data from a specific worksheet to a database. The `TestDbEntities` represents the database context, and `WorkBook.Load` function from IronXL loads the `testFile.xlsx` from the current directory. The `GetWorkSheet` method selects "Sheet3" for processing. The worksheet data is converted into a `DataTable`, and each row is iterated over to extract the "CountryName" and add it to the database before saving the changes.

```cs
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
System.Data.DataTable dataTable = sheet.ToDataTable(true);
foreach (DataRow row in dataTable.Rows)
{
    Country countryEntry = new Country();
    countryEntry.CountryName = row[1].ToString();
    dbContext.Countries.Add(countryEntry);
}
dbContext.SaveChanges();
```

<hr class="separator">

### Additional Resources

For those interested in deepening their knowledge of IronXL, it is beneficial to check out more tutorials in this series as well as the examples featured on our homepage, which are typically sufficient to get most developers up to speed.

For detailed information on the `WorkBook` class, refer to our [API Reference](https://ironsoftware.com/csharp/excel/object-reference/), which provides specific details.

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
      <a class="btn btn-white3" href="/csharp/excel/tutorials/downloads/Use.CSharp.to.Open.&.Write.an.Excel.File.zip">
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
      <a class="doc-link" href="https://github.com/iron-software/tutorials/tree/master/IronXL/Use%20C%23%20to%20Open%20%26%20Write%20an%20Excel%20File" target="_blank">How to Open and Write Excel File in C# on GitHub<i class="fa fa-chevron-right"></i></a>
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
      <h3>API Reference for IronXL</h3>
      <p>Explore the API Reference for IronXL, outlining the details of all of IronXL’s features, namespaces, classes, methods fields and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

