# C# Excel File Manipulation Tutorial

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file_old_changed may 2021/>***


Discover through guided examples how to generate, access, and store Excel documents using C#, executing fundamental functions such as obtaining totals, averages, counts, and more. IronXL.Excel serves as an independent .NET library capable of handling numerous spreadsheet types. Importantly, it functions independently of [Microsoft Excel](https://products.office.com/en-us/excel) and does not rely on Interop.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>How To Open & Write Excel Files in C# .NET:</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-install-the-ironxl-c-library-free">Install the IronXL C# Library</a></li>
        <li><a href="#anchor-2-2-create-a-new-excel-file">Create a New Excel File using CSV, XML, or JSON</a></li>
        <li><a href="#anchor-4-working-with-workbooks-with-sheets">Read and Write Multisheet Workbooks</a></li>
        <li><a href="#anchor-3-advanced-operations-sum-avg-count-etc">Apply Functions for SUM, AVG, Count, conditionals and more</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <a href="/downloads/assets/excel/tutorials/csharp-open-write-excel-file/tutorial-open-and-write-excel.pdf" target="_blank">
          <img style="box-shadow: none; width: 308px; height: 320px;" src="/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write.svg" data-hover-src="/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write-hover.svg" class="img-responsive learn-how-to-img replaceable-img">
        </a>
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h4 class="tutorial-segment-title">Overview</h4>
<h2>Use IronXL to Open and Write Excel Files</h2>

Explore the functionalities of reading, writing, modifying, and saving Excel files effortlessly using the [IronXL C# library](https://ironsoftware.com/csharp/excel/).

To get started, either download a [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or proceed with your existing project by following this guide:

1. Add the IronXL Excel Library to your project from [NuGet](https://www.nuget.org/packages/IronXL.Excel) or install using the DLL.
2. Open any XLS, XLSX, or CSV document by leveraging the `WorkBook.Load` method.
3. Retrieve cell values eloquently: `sheet["A11"].DecimalValue` is all it takes to obtain a decimal value from a cell.

This tutorial will guide you through the process of:

- **Installing IronXL to Your Project:** Instructions on how to integrate the IronXL library into your existing project.
- **Basic Operations with Excel Files:** Steps to create or open a workbook, select sheets and cells, and save your work.
- **Advanced Sheet Manipulations:** Enhance your spreadsheets by adding headers, footers, performing mathematical operations, and more to maximize your Excel file interactions.

<h4>Open an Excel File : Quick Code</h4>

```cs
using IronXL;
using System;

// Load the workbook
WorkBook workbook = WorkBook.Load("test.xlsx");

// Access the default worksheet
WorkSheet sheet = workbook.DefaultWorkSheet;

// Select a range of cells
Range range = sheet["A2:A8"];

decimal total = 0;

// Iterate through each cell in the range
foreach (var cell in range)
{
    Console.WriteLine($"Cell {cell.RowIndex} contains '{cell.Value}'");

    // Ensure the cell contains numeric data
    if (cell.IsNumeric)
    {
        // Accumulate the decimal value to maintain precision
        total += cell.DecimalValue;
    }
}

// Verify the sum calculated against a value in the worksheet
if (sheet["A11"].DecimalValue == total)
{
    Console.WriteLine("Verification Successful: Basic Test Passed");
}
```

<h4>Write and Save Changes to the Excel File : Quick Code</h4>

Here's the paraphrased version of the given C# code snippet, with adjusted comments and slight modifications for clarity:

```cs
// Assign value 11.54 to cell B1
sheet["B1"].Value = 11.54;

// Persist modifications by saving to a file
workbook.SaveAs("test.xlsx");
```

<hr class="separator">




<h4 class="tutorial-segment-title">Step 1</h4>

## Installing the Free IronXL C# Library

The IronXL.Excel library offers robust and adaptable solutions for managing Excel files in .NET environments. It supports a variety of .NET project types including Windows applications, ASP.NET MVC, and .NET Core applications, allowing for easy integration across different development setups.

<h3>Install the Excel Library to your Visual Studio Project with NuGet</h3>

The initial phase involves integrating the IronXL.Excel library into your project. There are two primary methods available for this integration: either through the NuGet Package Manager UI or via the NuGet Package Manager Console.

To incorporate IronXL.Excel into your project using the graphical NuGet Package Manager interface, follow these steps:

1. Using the mouse, right-click on the name of your project and choose "Manage NuGet Packages".

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <p><img src="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

Under the "Browse" tab, locate `IronXL.Excel` by searching for it and proceed to install it.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" target="_blank">

<p><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
```

</a>

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
  <img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>
```
The image above shows the completion of the library installation. Click on the image for a full view.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>



<h3>Install Using NuGet Package Manager Console</h3>

1. Navigate from the "Tools" menu to "NuGet Package Manager" and then select "Package Manager Console".

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

Below is the paraphrased section with adjusted URL paths:

```
2. Execute the following command to install IronXL:
```
Install-Package IronXL.Excel -Version 2019.5.2
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<h3>Manually Install with the DLL</h3>

You also have the option to manually add the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) to your project or the global assembly cache.

```
PM> Install-Package IronXL.Excel
```

# C# Tutorial for Opening and Writing to Excel Files

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file_old_changed may 2021/>***


Embark on a comprehensive guide on how to generate, access, and store Excel documents using C#. The IronXL.Excel library, a robust standalone .NET library, enables manipulation of various spreadsheet formats without the need for Microsoft Excel or Interop.

## Getting Started with IronXL in C# .NET:
Learn the fundamental steps to open and write Excel files by exploring the following points:
- [Install IronXL C# Library](#anchor-1-install-the-ironxl-c-library-free)
- [Create a New File using CSV, XML, or JSON](#anchor-2-2-create-a-new-excel-file)
- [Manipulate Multisheet Workbooks](#anchor-4-working-with-workbooks-with-sheets)
- [Utilize Advanced Functions like SUM, AVG, Count, etc.](#anchor-3-advanced-operations-sum-avg-count-etc)

![How to Open and Write Excel File in C#](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write.svg)

<div style="text-align: center;">[Download PDF Tutorial](https://ironsoftware.com/downloads/assets/excel/tutorials/csharp-open-write-excel-file/tutorial-open-and-write-excel.pdf)</div>

<hr class="divider">

## Comprehensive Overview
### Master Excel File Operations with IronXL

Utilize the IronXL C# library to seamlessly open, write, modify, and save Excel files. Start by either downloading the [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or use your personal project as a base to follow along this tutorial.

Steps involved:
1. Acquire the IronXL Excel library either through [NuGet](https://www.nuget.org/packages/IronXL.Excel) or direct DLL download.
2. Load an Excel file (XLS, XLSX, or CSV) with the `WorkBook.Load` method.
3. Retrieve cell values using simple commands like `sheet["A11"].DecimalValue`.

Throughout this tutorial, we'll detail every step needed, from installation and basic manipulations, to advanced techniques involving multi-sheet operations and data manipulation.

### Rapid Tutorial on Opening Excel Files

```csharp
// Integrate the IronXL namespace
using IronXL;
using System;

// Load an existing workbook
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet sheet = workbook.DefaultWorkSheet;

// Define a range of cells for operations
Range range = sheet["A2:A8"];
decimal total = 0;

// Loop through the cells in the specified range
foreach (var cell in range)
{
    Console.WriteLine($"Cell {cell.RowIndex} contains: '{cell.Value}'");

    // Summing only numeric cells
    if (cell.IsNumeric)
    {
        total += cell.DecimalValue; // Ensures precision
    }
}

// Testing formula evaluation
if (sheet["A11"].DecimalValue == total)
{
    Console.WriteLine("Validation Successful");
}
```

### Steps to Write and Persist Modifications

```csharp
// Assign a new value to a specific cell
sheet["B1"].Value = 11.54;

// Commit changes by saving the workbook
workbook.SaveAs("test.xlsx");
```

<hr class="divider">

By the end of this tutorial, you will be able to leverage IronXL to manipulate Excel files efficiently using C#. For further learning, explore additional operations and functions provided by IronXL.

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Key Operations: Creating, Opening, and Saving Excel Files ##

Learn the fundamental operations of handling Excel files, including creation, opening, and saving:
  
### 2.1. Tutorial Example: Creating a "HelloWorld" Console Application ###

**Step-by-step creation of a HelloWorld Project:**

1. **Start your project in Visual Studio**  
   Initiate a new project by launching Visual Studio.  
   ![Open Visual Studio](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png)

2. **Create a New Project**  
   Click on 'Create New Project'.  
   ![Choose to create a new project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png)

3. **Select the Console App (.NET framework)**  
   Choose this option for your project type.  
   ![Select Console App](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg)

4. **Name your project “HelloWorld” and create it**  
   This step will establish your project in the system.  
   ![Name and create your project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg)

5. **Verify the creation of your console application**  
   Ensure your console application is ready.  
   ![Confirm the creation of the application](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg)

6. **Install IronXL.Excel by adding it through NuGet Package Manager**  
   Click on 'Install' to add IronXL.Excel to your project.  
   ![Add and install IronXL.Excel](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg)

7. **Implement your initial code to read and display the first cell from a worksheet**:  
   Using the code below, read the first cell from the 'HelloWorld.xlsx' file and print its content.

```cs
static void Main(string[] args)
{
    var workbook = IronXL.WorkBook.Load(System.IO.Directory.GetCurrentDirectory() + @"\Files\HelloWorld.xlsx");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet["A1"].StringValue;
    Console.WriteLine(cell);
}
```

### 2.2. Creating a New Excel File ###

**Generating a new Excel file with IronXL**:

```cs
// Create a new Excel file
static void Main(string[] args)
{
    var newXLFile = WorkBook.Create(IronXL.ExcelFileFormat.XLSX);
    newXLFile.Metadata.Title = "IronXL New File";
    var newWorkSheet = newXLFile.CreateWorkSheet("1stWorkSheet");
    newWorkSheet["A1"].Value = "Hello World";
    newWorkSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
    newWorkSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

### 2.3. Opening Various Formats as Workbook ###

**Handling different file formats (CSV, XML, JSON) as Workbooks**: 

Each code snippet below demonstrates opening different file formats within a C# environment, adapting them to function as Excel workbooks using the IronXL library.

- **Opening a CSV file**:
  
```cs
// Open a CSV file as a Workbook
static void Main(string[] args)
{
    var workbook = IronXL.WorkBook.Load(System.IO.Directory.GetCurrentDirectory() + @"\Files\CSVList.csv");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet["A1"].StringValue;   
    Console.WriteLine(cell);
}
```

- **Opening an XML file**:

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
  <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
  <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
  <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

- **Opening a JSON file list as workbook**:

```cs
/**
 * Open JSON as Workbook
 */
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
        "code":
        "UM"
    }
]
```
Each of these code examples guides developers in incorporating data from various file formats into their .NET applications seamlessly using IronXL's extensive library functionalities.

### 2.1. Example Project: HelloWorld Console App ###

<p class="list-description">Initiate a HelloWorld Project</p>

<p class="list-decimal">2.1.1. Launch Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Select 'Create New Project'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Opt for 'Console App (.NET Framework)'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Name your sample project 'HelloWorld' and then click 'Create'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. Your console application is now set up</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Integrate IronXL.Excel into your project and proceed to install</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Add your initial code to read the first cell from the first sheet in the Excel file, then display it</p>

```cs
static void Main(string [] args)
{
    var workbook = IronXL.WorkBook.Load($@"{System.IO.Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet ["A1"].StringValue;
    Console.WriteLine(cell);
}
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

Here's the paraphrased section of the C# code you provided:

```cs
static void Main(string [] args)
{
    // Load the workbook from a specified path
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");

    // Retrieve the first worksheet in the workbook
    var sheet = workbook.DefaultWorkSheet;

    // Access the value in cell A1 as a string
    var cellValue = sheet ["A1"].StringValue;

    // Print the value of cell A1 to the console
    Console.WriteLine(cellValue);
}
```

### 2.2. Excel File Creation ###

<p class="list-description">Initiate a new Excel file using the IronXL library.</p>

```cs
/**
Create a Fresh Excel Document
anchor-create-a-new-excel-file
**/
static void Main(string [] args)
{
    var newExcelDocument = WorkBook.Create(ExcelFileFormat.XLSX);
    newExcelDocument.Metadata.Title = "Fresh IronXL Document";
    var initialSheet = newExcelDocument.CreateWorkSheet("FirstSheet");
    initialSheet ["A1"].Value = "Hello World";
    initialSheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    initialSheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

<p class="list-description">Create a new Excel file using IronXL</p>

Here's the paraphrased code section:

```cs
// Initialize a new Excel file using IronXL
static void Main(string[] args)
{
    // Create a new Excel file in XLSX format
    var excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "New IronXL Document";

    // Add a worksheet named 'FirstSheet'
    var sheet = excelFile.CreateWorkSheet("FirstSheet");

    // Set the value of cell A1
    sheet["A1"].Value = "Hello World";

    // Style cell A2 with a bottom border
    sheet["A2"].Style.BottomBorder.SetColor("#ff6600");
    sheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

This version maintains the same operations as the original code but uses slightly different variable names and comments to enhance clarity.

### 2.3. Loading (CSV, XML, JSON List) into a Workbook ###

#### 2.3.1. Opening a CSV File ####

Start by opening a CSV file format:

```cs
static void Main(string [] args)
{
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet["A1"].StringValue;	
    Console.WriteLine(cell);
}
```

#### 2.3.2. Managing an XML File ####
Create and manage an XML file by defining elements and attributes to represent countries:

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
		<country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
		<country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
		<country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
		<country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

Load the XML as a workbook:

```cs
static void Main(string [] args)
{
    var xmldataset = new DataSet();
    xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
    var workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

#### 2.3.3. Opening a JSON List ####
Create a JSON structure to represent a list of countries:

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

Construct the `CountryModel` class to map the JSON structure:

```cs
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

Utilize the Newtonsoft library to convert the JSON list into a set of `CountryModel` objects and load it into a workbook:

```cs
static void Main(string [] args)
{
    var jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
    var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
    var dataset = countryList.ToDataSet();
    var workbook = IronXL.WorkBook.Load(dataset);
    var sheet = workbook.WorkSheets.First();
}
```

This process translates your data from various structured formats into a fully functional Excel workbook using IronXL.

<p class="list-decimal">2.3.1. Open CSV file</p>

<p class="list-decimal">2.3.2 Create a new text file and add to it a list of names and ages (see example) then save it as CSVList.csv</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Your code snippet should look like this</p>

```cs
// Load a CSV file into an IronXL.WorkBook object
static void Main(string [] args)
{
    // Load the CSV file into a workbook
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");

    // Access the first worksheet in the workbook
    var sheet = workbook.WorkSheets.First();

    // Retrieve the value from the first cell (A1) and convert it to a string
    var cell = sheet ["A1"].StringValue;

    // Output the string value of the cell to console
    Console.WriteLine(cell);
}
```

<p class="list-decimal">

### 2.3.3. Load an XML File ###

<p class="list-decimal">2.3.3. Loading an XML File</p>
<p class="list-description">Start by crafting an XML file containing a list of countries. Each country entry in the file should be structured with specific attributes like code, continent name, and other relevant details.</p>

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
  <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
  <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
  <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">Utilize the following script to read the XML content into a workbook:</p>

```cs
/**
Load XML into a Workbook
anchor-open-xml-file
**/
static void Main(string [] args)
{
   var xmldataset = new DataSet();
   xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
   var workbook = IronXL.WorkBook.Load(xmldataset);
   var sheet = workbook.WorkSheets.First();
}
```

This approach lays out a comprehensive method for transforming XML data into an efficiently manageable Excel format, leveraging IronXL technology.

<span class="list-description">Create an XML file that contains a countries list: the root element “countries”, with children elements “country”, and each country has properties that define the country like code, continent, etc.</span>
</p>

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England, Scotland, Wales, GB, UK, Great Britain, Britain, Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US, America, USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">2.3.4. Copy the following code snippet to open XML as a workbook</p>

Here's the paraphrased section with updated relative URL paths resolved to `ironsoftware.com`:

```cs
// Loading an XML file into an IronXL Workbook
static void Main(string [] args)
{
    // Creating a new DataSet to hold the XML content
    var dataSet = new DataSet();
    // Reading XML file from the specified directory
    dataSet.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
    
    // Loading the DataSet into a workbook
    var workbook = IronXL.WorkBook.Load(dataSet);
    // Selecting the first worksheet in the workbook
    var worksheet = workbook.WorkSheets.First();
}
```

<p class="list-decimal">

### 2.3.5. Open JSON List as Workbook

To incorporate a list of countries formatted in JSON into your Excel workbook, start by creating a JSON file with the following structure. This file should enumerate various countries along with their codes:

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

Once the JSON is prepared, define a data model in C# that matches the JSON structure, as shown below:

```cs
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

Next, add the necessary code to parse this JSON file and convert it into a format that IronXL can work with. This involves deserializing the JSON into a list of `CountryModel` and then converting that list into a `DataSet`:

```cs
static void Main(string [] args)
{
    var jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
    var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel []>(jsonFile.ReadToEnd());
    var dataSet = countryList.ToDataSet();
    var workbook = IronXL.WorkBook.Load(dataSet);
    var sheet = workbook.WorkSheets.First();
}
```

Ensure that the `Newtonsoft.Json` library is added to your project for JSON serialization and deserialization operations. This integration allows the JSON data to seamlessly populate an Excel workbook using IronXL.

<span class="list-description">Create JSON country list</span>
</p>

Here's the paraphrased section of the original article, with modified JSON example structure and resolved relative URLs:

```cs
/**
Load JSON to Initialize Workbook
anchor-start-workbook-from-json-example
**/
[
    {
        "country": "United Arab Emirates",
        "isoCode": "AE"
    },
    {
        "country": "United Kingdom",
        "isoCode": "GB"
    },
    {
        "country": "United States",
        "isoCode": "US"
    },
    {
        "country": "United States Minor Outlying Islands",
        "isoCode": "UM"
    }
]
```

<p class="list-decimal"></p>
<p class="list-decimal">2.3.6. Create a country model that will map to JSON</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Here is the class code snippet</p>

```cs
public class Nation
{
    public string Name { get; set; }
    public string Code { get; set; }
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

Here's the paraphrased section of the article with updated links and images paths resolved to `ironsoftware.com`:

```cs
/**
Transform List into DataSet
identifier-open-csv-xml-json-list-as-workbook
**/
public static class ListConversionExtension
{
    public static DataSet ConvertListToDataSet<T>(this IList<T> list)
    {
        Type typeOfElement = typeof(T);
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable();
        dataSet.Tables.Add(table);

        // Create a DataColumn in DataTable for each property in the type T
        foreach (var propertyInfo in typeOfElement.GetProperties())
        {
            Type columnType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
            table.Columns.Add(propertyInfo.Name, columnType);
        }

        // Populate the DataTable with values from the list
        foreach (T element in list)
        {
            DataRow dataRow = table.NewRow();
            
            foreach (var propertyInfo in typeOfElement.GetProperties())
            {
                dataRow[propertyInfo.Name] = propertyInfo.GetValue(element, null) ?? DBNull.Value;
            }
            
            table.Rows.Add(dataRow);
        }

        return dataSet;
    }
}
```

<p class="list-decimal"></p>
<p class="list-decimal">And finally load this dataset as a workbook</p>

Here is the paraphrased section of the article, with resolved relative URL paths and rewritten code to ensure it is both distinct and related to the original:

```cs
static void Main(string[] args)
{
    // Read JSON data from file
    using (var reader = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json"))
    {
        // Deserialize JSON file to array of CountryModel
        var listOfCountries = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(reader.ReadToEnd());

        // Convert list to a DataSet for Excel processing
        var dataSet = listOfCountries.ToDataSet();

        // Load the dataset into an IronXL workbook
        var excelWorkbook = IronXL.WorkBook.Load(dataSet);

        // Access the first worksheet in the workbook
        var initialSheet = excelWorkbook.WorkSheets.First();
    }
}
```

### Saving and Exporting Excel Files ###

In this segment, we explore various methods to export your Excel workbooks into different file formats using IronXL, including XLSX, CSV, JSON, and XML.

#### 2.4.1. Saving as XLSX ####
To preserve your workbook in the XLSX format, simply use the `SaveAs` method:

```cs
WorkBook newXLFile = WorkBook.Create(ExcelFileFormat.XLSX);
newXLFile.Metadata.Title = "New IronXL File";
WorkSheet newWorkSheet = newXLFile.CreateWorkSheet("FirstWorkSheet");
newWorkSheet["A1"].Value = "Hello World";
newWorkSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
newWorkSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

newXLFile.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
```

#### 2.4.2. Saving as CSV ####
To export your Excel data to a CSV file, you can customize the delimiter according to your requirements:

```cs
newXLFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter: "|");
```

#### 2.4.3. Exporting to JSON ####
You can serialize and save your Excel data in JSON format using the `SaveAsJson` method:

```cs
newXLFile.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```
This will result in a JSON file structured as follows:
```cs
[
    ["Hello World"],
    [""]
]
```

#### 2.4.4. Saving as XML ####
Finally, you can also save your data in XML format, which structures your data hierarchically in tags:

```cs
newXLFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
```
The resulting XML file will look like this:
```xml
<?xml version="1.0" standalone="yes"?>
<_x0031_stWorkSheet>
  <_x0031_stWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Column1>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </_x0031_stWorkSheet>
</_x0031_stWorkSheet>
```

This guide should assist you in utilizing IronXL to save and export Excel files in a variety of formats, enhancing the flexibility of your .NET applications in handling data.

<span class="list-description">We can save or export the Excel file to multiple file formats like (“.xlsx”,”.csv”,”.html”) using one of the following commands.</span>

<p class="list-decimal">

### 2.4.1 Saving to XLSX Format ###

To save an Excel workbook to an XLSX format, use the `SaveAs` method. This function allows the workbook to be saved in the familiar XLSX file format, preserving all formats, values, and styles established in your spreadsheet.

Here's how to apply this:

```cs
/**
Save to XLSX Format
anchor-save-to-xlsx
**/
static void Main(string[] args)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    workbook.Metadata.Title = "New Excel File with IronXL";
    var worksheet = workbook.CreateWorkSheet("FirstSheet");
    worksheet["A1"].Value = "Hello, Excel!";
    worksheet["A2"].Style.BottomBorder.SetColor("#ff6600");
    worksheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    // Save the workbook to a file in the current directory with an .xlsx extension
    workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\MyNewExcelFile.xlsx");
}
```

This code snippet initially creates a new workbook, adds a worksheet with some data and styling, and eventually saves it to the `.xlsx` format, ensuring it is ready for further use or distribution.

<span class="list-description">To Save to “.xlsx” use saveAs function</span>
</p>

```cs
/**
Save and Export
anchor-save-and-export
**/
static void Main(string [] args)
{
    // Create a new Excel workbook with XLSX format
    WorkBook myWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
    myWorkbook.Metadata.Title = "IronXL New File";

    // Create a worksheet called "1stWorkSheet"
    WorkSheet mySheet = myWorkbook.CreateWorkSheet("1stWorkSheet");

    // Set the value of cell A1
    mySheet ["A1"].Value = "Hello World";

    // Style cell A2 with a dashed bottom border and specific color
    mySheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    mySheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    // Save the workbook to a file in the current directory
    myWorkbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
}
```

<p class="list-decimal">

### 2.4.2. Save as CSV Format

<span class="list-description">Using the `SaveAsCsv` method, we can export the Excel document as a CSV file. This method requires two arguments: the path and filename for the saved file, and the delimiter character used in the CSV, such as ",", "|", or ":".</span>
```

</p>

```cs
// Save the Excel data as a CSV file with specified delimiter
newXLFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter:",");
```

<p class="list-decimal">

### 2.4.3. Exporting to JSON Format

<span class="list-description">To export the Excel file into a JSON format, follow the method described below:</span>

```cs
/**
Export to JSON
anchor-save-as-json
**/
static void Main(string [] args)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    workbook.Metadata.Title = "Example IronXL File";
    var worksheet = workbook.CreateWorkSheet("SampleSheet");
    worksheet["A1"].Value = "Hello World";

    workbook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\ExampleJSON.json");
}
```

<span class="list-description">The resulting JSON file will appear as follows:</span>

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

Here is the revised version of the provided Markdown snippet with resolved URL paths:

```cs
// Export the workbook to a JSON file named HelloWorldJSON.json in the current directory.
newXLFile.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
``` 

This concise comment helps to elucidate what the line of code accomplishes, enhancing clarity and maintainability.

<p class="list-decimal">
  <span class="list-description">The result file should look like this</span>
</p>

```cs
[
    ["Hello World"], // First row containing the text "Hello World"
    [""]             // Second row is empty
]
```

<p class="list-decimal">

### 2.4.4. Export to XML Format

<span class="list-description">To export an Excel file to XML, use the `SaveAsXml` method as demonstrated below:</span>

```cs
/**
Export to XML
anchor-save-to-xml
**/
static void Main(string [] args)
{
    var newXLFile = WorkBook.Create(ExcelFileFormat.XLSX);
    newXLFile.Metadata.Title = "IronXL New File";
    var newWorkSheet = newXLFile.CreateWorkSheet("1stWorkSheet");
    newWorkSheet ["A1"].Value = "Hello World";
    newWorkSheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    newWorkSheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    newXLFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.xml");
}
```

<span class="list-description">After running the command, the resultant XML file will appear as follows:</span>

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

<span class="list-description">To save to xml use SaveAsXml as follow</span>
</p>

Here's the paraphrased section with relative URL paths resolved:

```cs
// Save the current workbook as XML format
newXLFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
```

<p class="list-decimal">
  <span class="list-description">Result should be like this</span>
</p>

Here's the paraphrased section with resolved URLs:

```html
<?xml version="1.0" standalone="yes"?>
<FirstWorkSheet>
  <FirstWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Column1>
  </FirstWorkSheet>
  <FirstWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </FirstWorkSheet>
</FirstWorkSheet>
```

<hr class="separator">

## 3. Advanced Excel Functions: Sum, Average, Count, and Others ##

Explore the implementation of typical Excel functions such as SUM, AVG, and Count using practical code examples.

### 3.1. Example: Calculating the Sum ###

<p class="list-description">This example demonstrates how to compute the total of a list of numbers stored in an Excel workbook titled "Sum.xlsx".</p>
![Sum Example](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png)

```cs
/**
Calculate Total Sum
anchor-sum-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
decimal totalSum = sheet["A2:A4"].Sum();
Console.WriteLine(totalSum);
```

<p class="list-description">Let’s find the sum for this list. I created an Excel file and named it “Sum.xlsx” and added this list of numbers manually</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

Here's the paraphrased section of the C# code snippet demonstrating the SUM function:

```cs
// Calculating the sum of cell values within a range using IronXL
var loadedWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var activeSheet = loadedWorkbook.WorkSheets.First();

// The Sum() method aggregates values from range A2 to A4
decimal totalSum = activeSheet["A2:A4"].Sum();
Console.WriteLine($"The sum of the range is: {totalSum}");
```

### 3.2 Example: Calculating the Average ###

Demonstrate how to determine the average value from a data set in an Excel file using this example:

```cs
/**
Computational Function for Average (AVG)
anchor-average-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
decimal average = sheet["A2:A4"].Avg();
Console.WriteLine(average);
```

<p class="list-description">Using the same file, we can get the average:</p>

Here's the paraphrased version of the given C# code snippet, enhanced with more detailed comments and adjusted to demonstrate different file paths and usage:

```cs
// Load an Excel workbook from a specific file path
var workbook = IronXL.WorkBook.Load($@"C:\Path\To\Your\Files\SumExample.xlsx");

// Retrieve the first worksheet from the workbook
var worksheet = workbook.WorkSheets.First();

// Calculate the average value of a range of cells (A2 to A4) in the worksheet
decimal averageValue = worksheet["A2:A4"].Avg();

// Output the average value to the console
Console.WriteLine($"The average value is: {averageValue}");
```

### Example: Count Function ###

In this section, we demonstrate how to tally the elements within a specific range in an Excel spreadsheet:

```cs
// Load workbooks and retrieve the desired worksheet
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();

// Use the Count method to count elements in the defined range
decimal countResult = sheet["A2:A4"].Count();
Console.WriteLine(countResult);
```

Here, after loading the worksheet, we utilize the `Count` method on the cell range from A2 to A4 to determine the number of entries in that particular segment. This example can be adjusted for various ranges as required in different contexts.

<p class="list-description">Using the same file, we can also get the number of elements in a sequence:</p>

```cs
/* Example: Counting Elements in a Range */
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.DefaultWorkSheet;  // Access the default worksheet
decimal numberOfCells = sheet["A2:A4"].Count();  // Calculate the count of cells in the range A2:A4
Console.WriteLine(numberOfCells); // Display the count on the console
```

### 3.4. Example: Finding the Maximum Value ###

Learn how to retrieve the maximum value from a range of cells in an Excel file, using IronXL. This example demonstrates reading a predefined list of values to calculate the maximum.

Observe the following illustration for clarity on the setup:

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

Here is the code snippet to extract the maximum value from a defined range of cells:

```cs
// Load the workbook with a predefined list of values
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();

// Calculate the maximum value from cells A2 to A4
decimal maxValue = sheet["A2:A4"].Max();
Console.WriteLine(maxValue);
```

For a more complex scenario where you want to apply a transformation function during the maximum value calculation, consider this:

```cs
// Reload the workbook for advanced example
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
// Calculate the maximum value using a transformation function (e.g., ignoring formulas)
bool maxWithCondition = sheet["A1:A4"].Max(c => c.IsFormula);
Console.WriteLine(maxWithCondition);
```

In this second example, the code attempts to find the maximum value considering only cells that are formulas, which should return "false" in the console as it illustrates a condition where no formula is evaluated as maximum.

<p class="list-description">Using the same file, we can get the max value of range of cells:</p>

Here's the paraphrased version of the provided C# code snippet:

```cs
/**
Function MAX: Example Usage
anchor-max-example
**/
var excelWorkbook = IronXL.WorkBook.Load($@"{System.Environment.CurrentDirectory}\Files\Sum.xlsx");
var activeSheet = excelWorkbook.DefaultWorkSheet;
decimal maximumValue = activeSheet["A2:A4"].Max();
Console.WriteLine(maximumValue);
```

<p class="list-description">– We can apply the transform function to the result of max function:</p>

Here's the paraphrased section of the C# code, with relative URL paths resolved to "ironsoftware.com":

```cs
// Load the workbook from the specified location
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Retrieve the first worksheet from the workbook
var sheet = workbook.WorkSheets.First();

// Determine if the maximum value in the specified range is a formula
bool isFormulaMax = sheet["A1:A4"].Max(c => c.IsFormula);

// Output the boolean result to the console
Console.WriteLine(isFormulaMax);
```

<p class="list-description">This example writes “false” in the console.</p>

### 3.5. Example of Finding the Minimum Value in a Range of Cells ###

This section explores how to determine the smallest value within a specified range of cells using IronXL. We'll perform this operation on a previously created Excel file named `Sum.xlsx`.

```cs
/**
Function MIN
anchor-min-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
bool minResult = sheet["A1:A4"].Min(); // Retrieve the minimal value from the range A1 to A4
Console.WriteLine(minResult); // Output the result to the console
```

<p class="list-description">Using the same file, we can get the min value of range of cells:</p>

```cs
/**
Method for Calculating Minimum
anchor-calculate-minimum
**/
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var excelSheet = excelWorkbook.WorkSheets.First();
bool minValue = excelSheet ["A1:A4"].Min();
Console.WriteLine(minValue);
```

### 3.6. Example: Sorting Cells ###

Learn how to sort cells either in ascending or descending order using IronXL. Below demonstrates how you can achieve this with data from an Excel spreadsheet.

```cs
// Load the workbook with data
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");

// Sort cells in ascending order; use `.SortDescending()` for descending
sheet["A1:A4"].SortAscending();

// Save the workbook with the sorted data
workbook.SaveAs("SortedSheet.xlsx");
```

Here, we load a workbook, select a worksheet, specify the cells range to sort, apply the sorting function, and finally, save the modified workbook under a new name. This method provides a seamless way to organize data in your spreadsheets.

<p class="list-description">Using the same file, we can order cells by ascending or descending:</p>

Here's a rewritten version of the provided section with resolved relative links and updated descriptions:

```cs
/**
Sort Cells in Ascending or Descending Order
anchor-sort-cells-example
**/
// Loading the workbook from a specified path
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Accessing the first worksheet in the workbook
var sheet = workbook.WorkSheets.First();

// Sorting the range A1 to A4 in ascending order
// To sort in descending order, replace SortAscending with sheet["A1:A4"].SortDescending();
sheet["A1:A4"].SortAscending();

// Saving the sorted workbook under a new file name
workbook.SaveAs("SortedSheet.xlsx");
```

Here, I improved the comments to clarify each operation, enhancing understandability for developers using or learning from this code snippet.

### 3.7. Example Using IF Condition ###

Explore how to deploy conditional formulas within an Excel file:

```cs
/**
Conditional IF Example
anchor-if-condition-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
int i = 1;
foreach(var cell in sheet ["B1:B4"])
{
    // Apply an IF condition to check if value in column A is >= 20 to assign "Pass" or "Fail"
    cell.Formula = "=IF(A" + i + ">=20,\"Pass\",\"Fail\")";
    i++;
}
// Save changes to a new file
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\UpdatedExcelFile.xlsx");
```

This example above demonstrates how to set a conditional formula in cells to evaluate based on the data in another column within the Excel sheet. After setting the conditions, the workbook is saved to a new file location, ensuring that changes are stored successfully.

<p class="list-description">Using the same file, we can use the Formula property to set or get a cell’s formula:</p>

<p class="list-decimal">3.7.1. Save to XML “.xml”</p>

Here's your paraphrased section with resolved URL paths and slight modifications for clarity and uniqueness in the code:

```cs
/**
Evaluate Conditions with IF
anchor-evaluate-conditions-if-example
**/
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var currentSheet = excelWorkbook.WorkSheets.First();
int index = 1;
foreach(var currentCell in currentSheet ["B1:B4"])
{
    // Apply IF condition to check if the value in column A is 20 or more
    currentCell.Formula = $"=IF(A{index}>=20,\"Pass\",\"Fail\")";
    index++;
}
// Save the modified workbook to a new file
excelWorkbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\UpdatedExcelFile.xlsx");
```

<p class="list-decimal">7.2. Using the generated file from the previous example, we can get the Cell’s Formula:</p>

Here's the paraphrased section of the code, with relative URL paths resolved to `ironsoftware.com`:

```cs
// Load the workbook from the specified file
var workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\NewExcelFile.xlsx");
// Retrieve the first worksheet in the workbook
var sheet = workbook.WorkSheets.First();
// Iterate through a range of cells and output their formulas
foreach(var cell in sheet["B1:B4"])
{
    Console.WriteLine(cell.Formula);  // Print the formula of the cell
}
Console.ReadKey();  // Wait for a key press to close the console window
```

### 3.8. Example: Trimming Cells ###

To demonstrate how to remove extraneous spaces from cell contents, we've updated an `Sum.xlsx` spreadsheet by adding a column with extra spaces. Here’s how you can clean up the spaces using IronXL:

```cs
/**
Trim Function Example
anchor-trim-example-section
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
var sheet = workbook.WorkSheets.First();
int rowIndex = 1;

// Loop through the cells in column F and apply the TRIM formula
foreach (var cell in sheet["F1:F4"])
{
    cell.Formula = "=TRIM(D" + rowIndex + ")";
    rowIndex++;
}

// Save the workbook with the cleaned cells
workbook.SaveAs("TrimmedFile.xlsx");
```
Explore how to apply similar techniques with other Excel functions to enhance your data formatting capabilities. This streamlined approach to applying formulas ensures that your worksheets remain clutter-free and professionally maintained.

<p class="list-description">To apply trim function (to eliminate all extra spaces in cells), I added this column to the sum.xlsx file</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description">And use this code</p>

```cs
// Remove extra spaces from cells using the TRIM function in IronXL
var workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\NewExcelFile.xlsx");
var sheet = workbook.DefaultWorkSheet;  // Accessing the default worksheet
int index = 1;  // Initialize index for cell referencing

// Iterating over a specific range of cells to apply the TRIM formula
foreach (var cell in sheet["f1:f4"])
{
    cell.Formula = $"=TRIM(D{index})";  // Setting formula to trim contents in column D
    index++;  // Increment index for next cell formula
}

// Save the workbook with trimmed cell values to a new file
workbook.SaveAs("trimmedExcelFile.xlsx");
```
This code snippet demonstrates applying the TRIM function to clean up cell values from unwanted spaces in an Excel file using IronXL. Each specified cell formula in the range `f1:f4` gets updated to remove extra spaces from the corresponding cells in column D, enhancing data accuracy and appearance. The updates are then saved to a new Excel document named "trimmedExcelFile.xlsx".

<p class="list-description">Thus, you can apply formulas in the same way.</p>

<hr class="separator">

## 4. Managing Workbooks with Multiple Sheets ##

This section delves into handling Excel workbooks containing multiple sheets.

### 4.1 Accessing Data from Multiple Sheets in a Workbook ###

Learn how to seamlessly access and manipulate data from various sheets within a single Excel workbook. This section demonstrates the procedure using an example Excel file with named tabs 'Sheet1' and 'Sheet2.'

```cs
/**
Access Different Worksheets
anchor-read-data-from-multiple-sheets-in-the-same-workbook
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet specificSheet = workbook.GetWorkSheet("Sheet2");
var selectedRange = specificSheet["A2:D2"];
foreach(var cell in selectedRange)
{
    Console.WriteLine(cell.Text);
}
```
In this example, you'll notice how to specify and work with a particular worksheet using its name, facilitating interaction with different sets of data within the same file.

<p class="list-description">I created an xlsx file that contains two sheets: “Sheet1”,” Sheet2”</p>
<p class="list-description">Until now we used WorkSheets.First() to work with the first sheet. In this example we will specify the sheet name and work with it</p>

Here's the paraphrased section of your article with absolute URL paths resolved:

```cs
// Example to Access Multiple Sheets from Same Workbook
// tag-read-data-from-multiple-sheets-in-the-same-workbook
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet2"); // Access "Sheet2" specifically
Range range = sheet["A2:D2"]; // Define the range of cells

// Iterate through the selected range to print each cell's text content
foreach(var cell in range)
{
    Console.WriteLine($"Text in cell: {cell.Text}");
}
```

### 4.2. Incorporate a New Sheet into a Workbook ###

This section guides you on how to insert a new sheet into an existing workbook effectively using IronXL:

```cs
// Load an existing workbook from specified path
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create and name a new sheet within the loaded workbook
var newSheet = workbook.CreateWorkSheet("new_sheet");

// Set a value in a specific cell of the new sheet
newSheet ["A1"].Value = "Hello World";

// Save the changes to a new file
workbook.SaveAs("https://ironsoftware.com/csharp/excel/tutorials/downloads/newFile.xlsx"); 
``` 

In this example, you first open an existing workbook located at a given path. Next, you create a new worksheet named `"new_sheet"` within that workbook. Once the sheet is added, you can directly manipulate it, such as setting the value of cell `"A1"` to `"Hello World"`. Finally, the workbook is saved under a new filename, preserving all changes including the addition of the new sheet. This process demonstrates the seamless integration of new sheets into existing workbooks, enhancing flexibility and dynamic data management in your applications.

<p class="list-description">We can also add new sheet to a workbook:</p>

```cs
// Initialize 'Add New Sheet' process for an example workbook
// Tag: anchor-add-new-sheet-to-a-workbook
WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx"); // Load an existing workbook
WorkSheet sheet = workbook.CreateWorkSheet("new_sheet"); // Create a new worksheet named 'new_sheet'
sheet["A1"].Value = "Hello World"; // Set the value of cell A1 to 'Hello World'
workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx"); // Save the updated workbook to a new file
```

<hr class="separator">

## 5. Interacting with an Excel Database ##

Explore the processes of exporting and importing data between a database and Excel.

A database named "TestDb" has been set up for demonstration, which includes a table called Country. This table is structured with two columns: Id (int, primary key) and CountryName (string).

### 5.1. Populate Excel Sheet with Database Data ###

<p class="list-description">In this section, we'll demonstrate how to populate a new sheet with data extracted from the 'Country' table in our database.</p>

```cs
/**
Import Data to Sheet
anchor-fill-excel-sheet-with-data-from-database
**/
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

This code snippet illustrates how to connect to a database via a context (`TestDbEntities`), retrieve data (from `Countries`), and then systematically populate an Excel sheet with this data. Notably, it first creates headers for 'Id' and 'Country Name' before inserting the corresponding data from each `Country` record into successive rows. Finally, it saves the modified workbook to a new Excel file named `FilledFile.xlsx`.

<p class="list-description">Here we will create a new sheet and fill it with data from the Country Table</p>

Here's a revised version of the C# code snippet, which covers importing data from a database into an Excel sheet, demonstrating functionalities of IronXL's `WorkBook` and `WorkSheet` classes. Additionally, I've resolved the relative URL paths to absolute paths according to IronSoftware's domain:

```cs
/**
Populate Excel Sheet with Database Data
anchor-fill-excel-sheet-with-data-from-database
**/
// Instantiate the database context
TestDbEntities databaseContext = new TestDbEntities();

// Load an existing workbook or create a new one if the file does not exist
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create a new worksheet to store our data
WorkSheet databaseSheet = workbook.CreateWorkSheet("FromDatabase");

// Retrieve a list of countries from the database
List<Country> countries = databaseContext.Countries.ToList();

// Set the headers for the new worksheet
databaseSheet.SetCellValue(0, 0, "ID");
databaseSheet.SetCellValue(0, 1, "Country Name");

// Row index starting from 1, as 0 is used for header
int currentRow = 1;

// Populate the worksheet rows with data from the database
foreach (Country country in countries)
{
    databaseSheet.SetCellValue(currentRow, 0, country.Id);
    databaseSheet.SetCellValue(currentRow, 1, country.CountryName);
    currentRow++;
}

// Save the workbook with the new sheet
workbook.SaveAs("CompleteCountryList.xlsx");
```

### 5.2 Populate a Database from an Excel Spreadsheet

In this section, we will demonstrate how to import data into a `TestDb` database from an Excel sheet. This process involves reading the spreadsheet data into our application and then committing that data to the database.

```cs
/**
Import Data to Database
anchor-fill-database-with-data-from-excel-sheet
**/
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
System.Data.DataTable dataTable = sheet.ToDataTable(true);

// Iterate through the dataTable rows and populate the Country object
foreach (DataRow row in dataTable.Rows)
{
    Country country = new Country();
    country.CountryName = row[1].ToString();
    dbContext.Countries.Add(country);
}
dbContext.SaveChanges();
``` 

This snippet loads an Excel workbook and selects a specific worksheet named `"Sheet3"`. It then converts the worksheet directly into a `DataTable`, preserving the header row in the process. Each row in the DataTable is subsequently used to populate the properties of a `Country` object, which is added to our `DbContext`. Changes are finally committed to the database using the `SaveChanges()` method.

<p class="list-description">Insert the data to the Country table in TestDb Database</p>

```cs
/**
Load Excel Data into Database
anchor-import-data-from-excel-to-database
**/
TestDbEntities databaseContext = new TestDbEntities();
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet targetSheet = excelWorkbook.GetWorkSheet("Sheet3");
System.Data.DataTable sheetDataTable = targetSheet.ToDataTable(true);

// Iterate through each row and populate the database with the data
foreach (DataRow currentRow in sheetDataTable.Rows)
{
    Country newCountry = new Country();
    newCountry.CountryName = currentRow[1].ToString();
    databaseContext.Countries.Add(newCountry);
}

// Commit changes to the database
databaseContext.SaveChanges();
```

<hr class="separator">

### Additional Resources

For those interested in deepening their understanding of IronXL, we recommend exploring further tutorials in this series, or checking out the practical examples featured on our website, which are typically sufficient to help most developers begin their projects.

Visit our detailed [API Reference](https://ironsoftware.com/csharp/excel/object-reference/) for comprehensive information about the `WorkBook` class and other aspects of IronXL.

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

