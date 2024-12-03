# C# Tutorial on Opening and Writing Excel Files

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file_old_changed may 2021/>***


Learn through detailed examples how to open, create, and save Excel files using C#. Additionally, discover how to perform essential functions such as calculating sums, averages, counts, and more. IronXL.Excel is an independent .NET library capable of handling numerous spreadsheet formats. The use of [Microsoft Excel](https://products.office.com/en-us/excel) is not required, and it does not rely on Interop for its functionality.

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

Easily handle Excel files by opening, modifying, saving, and personalizing them with the user-friendly [IronXL C# library](https://ironsoftware.com/csharp/excel/).

You can start quickly by downloading a [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or using your own setup to proceed with the guide.

Here’s how to get started:

1. Obtain the IronXL Excel Library via [NuGet](https://www.nuget.org/packages/IronXL.Excel) or by downloading the DLL directly.
2. Load XLS, XLSX, or CSV documents using the `WorkBook.Load` method.
3. Access cell values easily with the syntax: `sheet["A11"].DecimalValue`.

Throughout this tutorial, we’ll guide you through:

- **IronXL.Excel Installation:** Detailed steps on incorporating IronXL.Excel into your existing project.
- **Basic Excel Operations:** Fundamental procedures such as creating or opening a workbook, navigating through sheets and cells, and saving your progress.
- **Advanced Sheet Manipulations:** Explore advanced functionalities like inserting headers or footers and performing mathematical operations, along with other sophisticated features.

<h4>Open an Excel File : Quick Code</h4>

```cs
using IronXL;
using System;

// Load the workbook from a given file
WorkBook workbook = WorkBook.Load("test.xlsx");
// Access the default worksheet
WorkSheet sheet = workbook.DefaultWorkSheet;

// Define a range of cells to work with
Range range = sheet["A2:A8"];

decimal sum = 0;

// Loop through each cell in the defined range
foreach (var cell in range)
{
    // Output the values in the console
    Console.WriteLine($"Cell {cell.RowIndex} has value '{cell.Value}'");

    // Sum up the numeric values to handle precision issues properly
    if (cell.IsNumeric)
    {
        sum += cell.DecimalValue;
    }
}

// Verify if the total calculated matches the expected value in the cell A11
if (sheet["A11"].DecimalValue == sum)
{
    Console.WriteLine("Validation Successful: Test Passed.");
}
```

<h4>Write and Save Changes to the Excel File : Quick Code</h4>

```cs
// Assign a numeric value to cell B1
sheet ["B1"].Value = 11.54;

// Persist modifications to the Excel file
workbook.SaveAs("test.xlsx");
```

<hr class="separator">
<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Free Installation of the IronXL C# Library

IronXL.Excel offers a robust and adaptable .NET library designed for manipulating Excel documents—enabling opening, reading, editing, and saving functionalities. This library supports integration with various .NET project types including Windows applications, ASP.NET MVC, and .NET Core applications.

<h3>Install the Excel Library to your Visual Studio Project with NuGet</h3>

The initial step involves setting up IronXL.Excel in your project. There are two primary methods for this: through the NuGet Package Manager or the NuGet Package Manager Console.

Here's how to integrate IronXL.Excel using the NuGet Package Manager, which offers a graphical user interface for ease:

1. Navigate to your project in the solution explorer, right-click, and choose 'Manage NuGet Packages'.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <p><img src="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

2. Navigate to the 'browse' tab, type in `IronXL.Excel` in the search bar, and initiate the installation.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" target="_blank">

![Search for IronXL Library](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg)

</a>

<p class="list-decimal">3. Installation complete</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
  <p><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>



<h3>Install Using NuGet Package Manager Console</h3>

1. Navigate to `Tools`, select `NuGet Package Manager`, and then click on `Package Manager Console`.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

To install the IronXL.Excel library using the Package Manager Console in Visual Studio, enter the following command:

```
PM > Install-Package IronXL.Excel -Version 2019.5.2
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<h3>Manually Install with the DLL</h3>

You also have the option to manually integrate the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) into your project or the global assembly cache.

```
PM> Install-Package IronXL.Excel
```

```
# C# Tutorial: Manipulating Excel Documents

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file_old_changed may 2021/>***


Learn how to build, open, and save Excel files using C#, as well as execute primary operations such as totaling, averaging, counting, among others. IronXL.Excel is a dedicated .NET library that can handle numerous spreadsheet formats without the need for Microsoft Excel installation or relying on Interop.

<div class="learn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Steps to Manage Excel Files using C# in .NET:</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-install-the-ironxl-c-library-free">Install IronXL C# Library</a></li>
        <li><a href="#anchor-2-2-create-a-new-excel-file">Construct a New Excel File from CSV, XML, or JSON</a></li>
        <li><a href="#anchor-4-working-with-workbooks-with-sheets">Manipulate Multisheet Workbooks</a></li>
        <li><a href="#anchor-3-advanced-operations-sum-avg-count-etc">Utilize Functions like SUM, AVG, Count, and More</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <a href="https://ironsoftware.com/downloads/assets/excel/tutorials/csharp-open-write-excel-file/tutorial-open-and-write-excel.pdf" target="_blank">
          <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write.svg" data-hover-src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write-hover.svg" class="img-responsive learn-how-to-img replaceable-img">
        </a>
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h4 class="tutorial-segment-title">Introduction</h4>
<h2>Work with Excel Files Using IronXL</h2>

Open, modify, save, and tailor Excel documents effortlessly using the <a href="https://ironsoftware.com/csharp/excel/" target="_blank">IronXL C# library</a>.

Download a <a href="https://github.com/magedo93/IronSoftware.git" target="_blank">GitHub sample project</a> or bring your own and progress through this guide. 

1. Add IronXL Excel Library from <a href="https://www.nuget.org/packages/IronXL.Excel" target="_blank">NuGet</a> or by direct DLL download
2. Employ `WorkBook.Load` method to open any XLS, XLSX, or CSV file.
3. Access cell values using straightforward syntax: `sheet["A11"].DecimalValue`

Throughout this tutorial, you’ll discover how to:

- Install IronXL.Excel: integrate IronXL.Excel into an existing solution.
- Basic Handling Techniques: steps to create or open workbooks, select sheets or cells, and save your work
- Advanced Sheet Functions: explore advanced functionality like adding headers, applying mathematical functions, and other operations.
```

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Core Functions: Creating, Opening, and Saving Excel Documents ##

### 2.1. Starting with a Hello World Console Application ###

<p class="list-description">Initiate a simple Hello World Project</p>

<p class="list-decimal">2.1.1. Begin by launching Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Opt for Create New Project</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Select Console App (.NET framework) as your project type</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Name your project “HelloWorld” and proceed to create it</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. Your Console Application is now ready</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Integrate IronXL.Excel into your project and begin installation</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Insert initial code to access and print the first cell in the first sheet of an Excel file</p>

```csharp
static void Main(string [] args)
{
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet ["A1"].StringValue;
    Console.WriteLine(cell);
}
```

### 2.2. Generating a Brand New Excel File ###

<p class="list-description">Initiate a new Excel document using IronXL</p>

```csharp
/**
Create Excel Document
anchor-create-a-new-excel-file
**/
static void Main(string [] args)
{
    var newExcelDoc = WorkBook.Create(ExcelFileFormat.XLSX);
    newExcelDoc.Metadata.Title = "New IronXL Document";
    var firstSheet = newExcelDoc.CreateWorkSheet("FirstSheet");
    firstSheet ["A1"].Value = "Hello World";
    firstSheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    firstSheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

### 2.3. Opening Different File Formats as Workbooks ###

<p class="list-decimal">2.3.1. Begin by opening a CSV file</p>

<p class="list-decimal">2.3.2. Create a plain text file, populate it with a list of names and ages and save as CSV format</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Example code snippet for opening a CSV as a workbook</p>

```csharp
/**
Open CSV as Workbook
anchor-open-csv-xml-json-list-as-workbook
**/
static void Main(string [] args)
{
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\List.csv");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet ["A1"].StringValue;	
    Console.WriteLine(cell);
}
```

### 2.1 Sample Project: HelloWorld Console Application ###

Discover how to get started with a simple "HelloWorld" project using IronXL.

#### 2.1.1 Start Visual Studio
Begin by launching Visual Studio.

![Open Visual Studio](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png)

#### 2.1.2 Create a New Project
Select the option to create a new project.

![Choose Create New Project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png)

#### 2.1.3 Select Console App
Choose the Console App option suitable for .NET framework.

![Select Console App](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg)

#### 2.1.4 Name Your Project
Name your project "HelloWorld" and click on the create button.

![Name Your Project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg)

#### 2.1.5 Project Creation
You'll see your new console application has been created.

![Project Created](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg)

#### 2.1.6 Install IronXL
Next, add the IronXL.Excel library. Simply initiate the installation with a click.

![Install IronXL](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg)

#### 2.1.7 Start Coding
Finally, add your first lines of code to read the first cell of the first sheet in an Excel file and display it.

```cs
static void Main(string[] args)
{
    var workbook = IronXL.WorkBook.Load($@"{System.IO.Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet["A1"].StringValue;
    Console.WriteLine(cell);
}
```

This simple exercise introduces you to opening and reading data from an Excel file using the IronXL library, setting a solid foundation for more complex data manipulation.

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

```cs
static void Main(string [] args)
{
    // Load the workbook from the HelloWorld excel file
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
    
    // Get the first worksheet from the workbook
    var sheet = workbook.WorkSheets.First();
    
    // Access and retrieve the string value from cell A1
    var cell = sheet["A1"].StringValue;
    
    // Print the value of cell A1 to the console
    Console.WriteLine(cell);
}
```

### 2.2: Creating a Fresh Excel Document ###

Discover how to craft a new Excel file utilizing the IronXL library with ease.

```cs
/**
* Initial setup for creating an Excel file
* anchor-create a fresh Excel document
**/
static void Main(string [] args)
{
    // Instantiate a new Excel file
    var excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "New Excel Document with IronXL";
    
    // Add a worksheet to the Excel file
    var sheet = excelFile.CreateWorkSheet("FirstSheet");
    sheet ["A1"].Value = "Hello, IronXL!";
    
    // Customize a cell with a bottom border
    sheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    sheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

This straightforward approach illustrates the creation of a new Excel workbook, addition of a worksheet, and customization of cells within it, using IronXL’s robust features.

<p class="list-description">Create a new Excel file using IronXL</p>

```cs
/**
Generate a New Excel Document
tag-create-new-excel-doc
**/
static void Main(string[] args)
{
    var excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "New IronXL Document";
    var sheet = excelFile.CreateWorkSheet("InitialSheet");
    sheet["A1"].Value = "Hello World";
    sheet["A2"].Style.BottomBorder.SetColor("#ff6600"); // Set the color of the bottom border
    sheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed; // Use a dashed line for the border
}
```

### 2.3. Import CSV, XML, and JSON Data as Workbooks ###

Discover how to seamlessly load data from various formats into your .NET applications using IronXL.

#### 2.3.1. Load a CSV File ####

Begin by creating a simple CSV file with a list of names and ages and save it as `CSVList.csv`. Below is what your CSV file's content might look like displayed in a code editor:

![Code Snippet](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg)

Now, to open this CSV file as a workbook, use the following snippet:

```cs
// Load CSV into Workbook
static void Main(string[] args)
{
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet["A1"].StringValue;
    Console.WriteLine(cell);
}
```

#### 2.3.2. Load an XML File ####

For XML, create a file that lists countries with attributes such as code, continent, etc. The XML structure might look like:

```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <country code="us" handle="united-states" continent="north america" iso="840">United States</country>
  ...
</countries>
```

Extract and convert this XML into a workbook using this approach:

```cs
// Load XML Data as Workbook
static void Main(string[] args)
{
    var xmldataset = new DataSet();
    xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
    var workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

#### 2.3.3. Load JSON Data ####

For JSON, consider a data structure like this to store country information:

```json
[
    {"name": "United States", "code": "US"},
    ...
]
```

To handle JSON data, first create a model class:

```cs
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

Then, use Newtonsoft to parse JSON data into a list of `CountryModel` objects. Convert this list to a `DataSet` and finally, load it into a workbook:

```cs
// JSON to Workbook
static void Main(string[] args)
{
    var jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
    var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
    var xmldataset = countryList.ToDataSet();
    var workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

By following these steps, you can easily handle and manipulate data in CSV, XML, and JSON formats within your .NET applications using the IronXL library.

<p class="list-decimal">2.3.1. Open CSV file</p>

<p class="list-decimal">2.3.2 Create a new text file and add to it a list of names and ages (see example) then save it as CSVList.csv</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Your code snippet should look like this</p>

Here's the paraphrased version of the provided C# code snippet with resolved URL paths for link references, as requested:

```cs
/**
Load CSV File as an IronXL Workbook
anchor-open-csv-xml-json-list-as-workbook
**/
static void Main(string [] args)
{
    // Load the CSV file into an IronXL workbook
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
    
    // Access the first worksheet in the workbook
    var sheet = workbook.WorkSheets.First();
    
    // Retrieve the value from cell A1 and print it
    var cellValue = sheet["A1"].StringValue;
    Console.WriteLine(cellValue);
}
```
This revised code snippet enhances clarity in comments and slightly alters variable naming for better comprehension.

<p class="list-decimal">

```html
2.3.3. Accessing XML Files

<p class="list-decimal">
    2.3.3. Load XML Data
    <span class="list-description">Create an XML document featuring a list of countries: begin with a root element "countries", include child elements "country", each with attributes defining aspects like code and continent.</span>
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

<p class="list-decimal">2.3.4. Implement the following code snippet to process the XML into a workbook:</p>

```cs
/**
Load XML into Workbook
anchor-open-csv-xml-json-list-as-workbook
**/
static void Main(string [] args)
{
    var xmldataset = new DataSet();
    xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
    var workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

<p class="list-decimal">
    2.3.5. Open and work with JSON Format
    <span class="list-description">Construct a JSON list of countries</span>
</p>

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

```cs
/**
Load XML into an Excel Workbook Example
anchor-load-xml-as-workbook
**/
static void Main(string [] args)
{
    // Create a new DataSet to hold the XML data
    var xmlDataset = new DataSet();
    // Read the XML data from a file into the DataSet
    xmlDataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");

    // Load the XML data into a new Excel workbook using IronXL
    var excelWorkbook = IronXL.WorkBook.Load(xmlDataset);
    // Access the first worksheet in the workbook
    var activeSheet = excelWorkbook.WorkSheets.First();
}
```

<p class="list-decimal">

### 2.3.5 Open JSON List as Workbook

This section covers how to open a JSON list as a workbook using IronXL:

```cs
/**
Open JSON as Workbook
anchor-open-csv-xml-json-list-as-workbook
**/
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

To handle JSON data efficiently, we need to model the JSON structure in a class. Here is an example of a class that maps to our JSON data:

```cs
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

To convert JSON to a dataset we can use in our Excel workbook, we include a Newtonsoft library and extend a custom converter:

```cs
/**
Convert JSON to DataSet
anchor-open-csv-xml-json-list-as-workbook
**/
public static class ListConvertExtension
{
    public static DataSet ToDataSet<T>(this IList<T> list)
    {
        Type elementType = typeof(T);
        DataSet ds = new DataSet();
        DataTable t = new DataTable();
        ds.Tables.Add(t);

        // Add a column to the table for each public property on T
        foreach (var propInfo in elementType.GetProperties())
        {
            Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
            t.Columns.Add(propInfo.Name, ColType);
        }

        // Populate the table with values from the list
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

Finally, load this dataset as a new workbook:

```cs
static void Main(string [] args)
{
    var jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
    var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
    var xmldataset = countryList.ToDataSet();
    var workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

<span class="list-description">Create JSON country list</span>
</p>

```cs
/**
Load JSON into Workbook
anchor-load-json-into-workbook
**/
[
    {
        "country": "United Arab Emirates",
        "countryCode": "AE"
    },
    {
        "country": "United Kingdom",
        "countryCode": "GB"
    },
    {
        "country": "United States",
        "countryCode": "US"
    },
    {
        "country": "United States Minor Outlying Islands",
        "countryCode": "UM"
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

Here is the paraphrased section with resolved relative URL paths:

```cs
/**
Transformation to DataSet
anchor-transform-list-to-dataset-for-workbooks
**/
public static class ConvertListToDataSet
{
    public static DataSet ConvertToDataSet<T>(this IList<T> list)
    {
        Type typeOfElement = typeof(T);
        DataSet dataSet = new DataSet();
        DataTable dataTable = new DataTable();
        dataSet.Tables.Add(dataTable);

        // For each public property of T, add a corresponding column to the DataTable
        foreach (var propertyInfo in typeOfElement.GetProperties())
        {
            Type columnType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
            dataTable.Columns.Add(propertyInfo.Name, columnType);
        }

        // Populate each row in the DataTable with values from the list
        foreach (T element in list)
        {
            DataRow dataRow = dataTable.NewRow();

            foreach (var propertyInfo in typeOfElement.GetProperties())
            {
                dataRow[propertyInfo.Name] = propertyInfo.GetValue(element, null) ?? DBNull.Value;
            }

            dataTable.Rows.Add(dataRow);
        }

        return dataSet;
    }
}
```

<p class="list-decimal"></p>
<p class="list-decimal">And finally load this dataset as a workbook</p>

Here's the paraphrased section of your article with annotations to explain the code and with updated relative paths resolved to `ironsoftware.com`:

```cs
static void Main(string[] args)
{
    // Load the JSON file containing country data
    var jsonFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Files", "CountriesList.json");
    using (var reader = new StreamReader(jsonFilePath))
    {
        // Deserialize JSON data into a list of CountryModel objects
        var countryData = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(reader.ReadToEnd());
        
        // Convert the list of CountryModel objects to a DataSet
        var countryDataSet = countryData.ToDataSet();
        
        // Load the dataset into an IronXL workbook
        var workbook = IronXL.WorkBook.Load(countryDataSet);
        
        // Retrieve the first worksheet from the workbook
        var sheet = workbook.WorkSheets.First();
        
        // Additional operations can be performed on 'sheet' as needed
    }
}
```
In this version, I included comments to make the code more understandable and used a `using` statement to ensure the StreamReader is properly disposed after reading the file. This ensures resources are managed properly, enhancing the code's robustness.

### 2.4. Saving and Exporting Files ###

<span class="list-description">IronXL allows for exporting the Excel file into several formats using the commands below.</span>

<p class="list-decimal">
    2.4.1. Export to XLSX
    <span class="list-description">To export a file as “.xlsx”, use the `SaveAs` method:</span>
</p>

```cs
/**
Export to XLSX Format
anchor-save-and-export-xlsx
**/
static void Main(string [] args)
{
    var excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "IronXL New File";
    var sheet = excelFile.CreateWorkSheet("FirstSheet");
    sheet ["A1"].Value = "Hello World";
    sheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    sheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    excelFile.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
}
```

<p class="list-decimal">
    2.4.2. Export to CSV
    <span class="list-description">To save the file as “.csv”, the `SaveAsCsv` method allows specifying the file name, path, and delimiter (e.g., “,”, “|”, “:”):</span>
</p>

```cs
excelFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter: "|");
```

<p class="list-decimal">
    2.4.3. Export to JSON
    <span class="list-description">For JSON format “.json”, use the `SaveAsJson` method:</span>
</p>

```cs
excelFile.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

<p class="list-decimal">
  <span class="list-description">The JSON file will appear as follows:</span>
</p>

```cs
[
    [
        "Hello World"
    ],
    [
        ""
    ]
]
```

<p class="list-decimal">
    2.4.4. Export to XML
    <span class="list-description">To save as XML, the `SaveAsXml` method can be used:</span>
</p>

```cs
excelFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
```

<p class="list-description">
  <span>The XML output will be structured like this:</span>
</p>

```html
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

<span class="list-description">We can save or export the Excel file to multiple file formats like (“.xlsx”,”.csv”,”.html”) using one of the following commands.</span>

<p class="list-decimal">

### 2.4.1. Save as XLSX File

<span class="list-description">Here's how you can save your document as an XLSX file using IronXL:</span>

```cs
/**
Save and Export
anchor-save-and-export
**/
static void Main(string [] args)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    workbook.Metadata.Title = "IronXL New File";
    var sheet = workbook.CreateWorkSheet("FirstSheet");
    sheet ["A1"].Value = "Hello World";
    sheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    sheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
}
```

<span class="list-description">To Save to “.xlsx” use saveAs function</span>
</p>

```cs
/**
Save and Export Operations
anchor-save-and-export-operations
**/
static void Main(string [] args)
{
    var excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "IronXL New File Example";
    var sheet = excelFile.CreateWorkSheet("FirstSheet");
    sheet["A1"].Value = "Hello World";
    sheet["A2"].Style.BottomBorder.SetColor("#ff6600");  // Set bottom border color to orange
    sheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;  // Set border type to dashed

    // Save the Excel file to the current directory
    excelFile.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
}
```

<p class="list-decimal">

2.4.2. Export to CSV

<span class="list-description">To convert and save the Excel file as a CSV, utilize the `SaveAsCsv` method which requires two arguments: the path and filename for the CSV, and the delimiter character, which could be a comma (","), pipe ("|"), or colon (":").</span>

</p>

Here's your paraphrased section with the relative URL paths resolved:

```cs
newXLFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter:",");
```

<p class="list-decimal">

### 2.4.3 Export to JSON Format

<span class="list-description">Exporting an Excel workbook to a JSON file can be efficiently done using the following method:</span>

```cs
/**
Export to JSON
anchor-export-to-json
**/
static void Main(string [] args)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    workbook.Metadata.Title = "New IronXL File";
    var worksheet = workbook.CreateWorkSheet("Sheet1");
    worksheet ["A1"].Value = "Sample Data";
    worksheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    worksheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    // Export the workbook to a JSON file
    workbook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\SampleData.json");
}
```

<p class="list-decimal">
  <span class="list-description">Here is how the JSON output appears:</span>
</p>

```cs
[
    [
        "Sample Data"
    ],
    [
        ""
    ]
]
```

This process highlights the simplicity of converting Excel data into JSON format, allowing for easy data exchange and integration with web applications and services.

<span class="list-description">To save to Json “.json” use SaveAsJson as follow</span>
</p>

```cs
// Saving the Excel file as JSON format
newXLFile.SaveAsJson(Path.Combine(Directory.GetCurrentDirectory(), "Files", "HelloWorldJSON.json"));
```

<p class="list-decimal">
  <span class="list-description">The result file should look like this</span>
</p>

Here is the paraphrased section of the JSON export example from the article, with the URL paths resolved to `ironsoftware.com`:

___

```cs
[
    [
        "Hello World"
    ],
    [
        ""
    ]
]
```

<p class="list-decimal">

### 2.4.4 Save as XML Format

By choosing to save your Excel document as an XML file, you can maintain the structure and access it easily across different platforms that support XML standards.

```cs
/**
Export to XML Format
**/
static void Main(string[] args) 
{
    var excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "IronXL New File";
    var workSheet = excelFile.CreateWorkSheet("1stWorkSheet");

    // Data input in the worksheet
    workSheet["A1"].Value = "Hello World";
    workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
    workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    // Saving as XML
    excelFile.SaveAsXml(@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.xml");
}
```

The output XML will look like this:

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

This allows you to transform your worksheet into a universally recognizable format, making it suitable for data interchange between applications that utilize XML.

<span class="list-description">To save to xml use SaveAsXml as follow</span>
</p>

```cs
// Save the Workbook as an XML file to the specified directory
newXLFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
```

<p class="list-decimal">
  <span class="list-description">Result should be like this</span>
</p>

```html
<?xml version="1.0" encoding="utf-8"?>
<FirstWorkSheet>
  <Row>
    <Cell xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Cell>
  </Row>
  <Row>
    <Cell xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></Cell>
  </Row>
</FirstWorkSheet>
```

<hr class="separator">

## 3. Enhanced Calculations: Sum, Average, Count, and More ##

Explore standard Excel functions such as SUM, AVG, and COUNT with detailed coding examples.

### 3.1. Example: Calculating the Sum ###

This example demonstrates how to calculate the total of a list of values in an Excel file. The file, named `Sum.xlsx`, contains a column of numbers, specifically in the range from cell `A2` to `A4`.

```cs
// Load the Excel workbook and access the first worksheet
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();

// Calculate the sum of values in the specified range
decimal sum = sheet["A2:A4"].Sum();

// Output the sum to the console
Console.WriteLine(sum);
```

This code snippet efficiently computes the sum of the numbers located in the cells from `A2` to `A4`. By loading the workbook and accessing the first worksheet, it leverages IronXL's ability to perform straightforward arithmetic operations directly on cell ranges. The result is then printed, providing immediate feedback on the operation's outcome.

<p class="list-description">Let’s find the sum for this list. I created an Excel file and named it “Sum.xlsx” and added this list of numbers manually</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

```cs
// Calculate the total sum of a specific range in an Excel sheet using IronXL
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var excelSheet = excelWorkbook.WorkSheets.First();
decimal totalSum = excelSheet["A2:A4"].Sum(); // Sum values between A2 and A4
Console.WriteLine(totalSum); // Display the result
```

### 3.2. Example: Calculating the Average ###

Discover how to compute the average value from an Excel spreadsheet using the following example:

```cs
// Load the workbook from the local directory
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet
var sheet = workbook.WorkSheets.First();

// Calculate the average of values in cells from A2 to A4
decimal average = sheet["A2:A4"].Avg();

// Output the average
Console.WriteLine(average);
```

<p class="list-description">Using the same file, we can get the average:</p>

Here's your paraphrased section with updates to relative URL paths and additional context in comments for clarity within the code snippet:

```cs
// Calculate the Average Value of a Specific Range in Excel
// Tag: anchor-avg-example
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First(); // Access the first worksheet
decimal averageValue = sheet["A2:A4"].Avg(); // Compute the average for cells from A2 to A4
Console.WriteLine(averageValue); // Output the computed average to the console
```

### 3.3. Example: Counting Cells ###

Here's how you can determine the total number of elements in a sequence from a spreadsheet:

```cs
/**
Count Elements in Range
anchor-count-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
decimal count = sheet["A2:A4"].Count();
Console.WriteLine(count);
```

<p class="list-description">Using the same file, we can also get the number of elements in a sequence:</p>

Here's the paraphrased section of the article you provided, with updated relative URL paths as specified:

```cs
/**
Method to Count Cells in Range
anchor-exemplify-count
**/
// Load the workbook
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet
var sheet = workbook.WorkSheets.First();

// Perform a count on the cell range from A2 to A4
decimal cellCount = sheet["A2:A4"].Count();

// Output the count to the console
Console.WriteLine(cellCount);
```

### 3.4. Example of Finding the Maximum Value ###

Discover how to identify the highest value within a range of cells in an Excel file using the IronXL library. This can be a useful function for data analysis and quick summaries.

```cs
// Load an existing workbook with numerical data
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();

// Calculate the maximum value from a specific range
decimal maximumValue = sheet ["A2:A4"].Max();
Console.WriteLine(maximumValue);
```

In this block of code, we:
1. Open an Excel file named `Sum.xlsx`.
2. Access the first worksheet.
3. Compute the maximum decimal value from the cells A2 to A4.
4. Output the highest number to the console.

Additionally, you can transform the results using the `Max` method with a lambda expression to evaluate conditions or formulas within the cells.

```cs
// Load the workbook
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();

// Find if the maximum value in the range A1 to A4 is defined by a formula
bool isFormulaMax = sheet ["A1:A4"].Max(c => c.IsFormula);
Console.WriteLine(isFormulaMax);
```

This piece demonstrates:
- Loading the same workbook and accessing the initial sheet.
- Evaluating whether the maximum value from cells A1 to A4 is derived from a formula, and then printing out `true` or `false` based on the assessment.

Using IronXL's functionality, interactions like these enhance the flexibility in processing spreadsheet data programmatically.

<p class="list-description">Using the same file, we can get the max value of range of cells:</p>

```cs
// Example: Finding the Maximum Value in a Range
// Tag: example-max-value
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.GetSheet("Sheet1");
decimal maximumValue = sheet["A2:A4"].Max();
Console.WriteLine("The maximum value is: {0}", maximumValue);
```

<p class="list-description">– We can apply the transform function to the result of max function:</p>

```cs
// Load the workbook from the current directory
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
// Access the first worksheet in the workbook
var sheet = workbook.WorkSheets.First();
// Determine the maximum value in the range A1 to A4, specifically checking for formulas
bool maxIsFormula = sheet["A1:A4"].Max(cell => cell.IsFormula);
// Print the result
Console.WriteLine(maxIsFormula);
```

<p class="list-description">This example writes “false” in the console.</p>

### 3.5. Example: Finding the Minimum Value ###

This portion of the tutorial demonstrates how to determine the smallest value within a specific range from an existing Excel document titled "Sum.xlsx". It specifically focuses on cells in the range from A2 to A4.

```cs
/**
Function MIN
anchor-min-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
decimal min = sheet["A2:A4"].Min();
Console.WriteLine(min);
```
This simple example uses the `Min` method to calculate the minimum value, providing a quick and efficient way to manage and analyze numerical data stored in Excel files using IronXL.

<p class="list-description">Using the same file, we can get the min value of range of cells:</p>

```cs
/**
Method to Calculate Minimum
anchor-min-value-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
bool minVal = sheet["A1:A4"].Min();
Console.WriteLine(minVal);
```

### 3.6. Example of Cell Ordering ###

Discover how to organize the cell entries in ascending or descending order within the same worksheet.

```cs
/**
Function Order Cells
anchor-order-cells-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
// Sorting the range in ascending order
sheet ["A1:A4"].SortAscending(); // Alternatively, use sheet ["A1:A4"].SortDescending(); for descending order
workbook.SaveAs("SortedSheet.xlsx");
```

<p class="list-description">Using the same file, we can order cells by ascending or descending:</p>

Here is the paraphrased section with resolved URL paths:

```cs
// Example of ordering cell values within a range in ascending or descending order
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx"); // Load the workbook
var sheet = workbook.WorkSheets.First(); // Access the first worksheet

// Sort data in range "A1:A4". Replace `SortAscending` with `SortDescending` if you need descending order.
sheet["A1:A4"].SortAscending(); // Uncomment the next line to sort in descending order
// sheet["A1:A4"].SortDescending();

workbook.SaveAs("SortedSheet.xlsx"); // Save the changes to a new file
```

This code snippet demonstrates how to sort a specific range of cells in an Excel worksheet either in ascending or descending order and then save the changes to a new Excel file using the IronXL library.

### 3.7. Usage Example: Conditional Statements ###

Let’s delve into how we can apply conditional logic directly within an Excel spreadsheet.

For instance, consider an Excel file named `Sum.xlsx`. Based on a certain condition, we want to evaluate whether scores meet a pass/fail criterion:

```cs
/**
Conditional Formulas in Cells
anchor-if-condition-tutorial
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
int rowNum = 1;
foreach(var cell in sheet ["B1:B4"])
{
    // Set formula to check if the corresponding score in column A is 20 or above
    cell.Formula = $"=IF(A{rowNum}>=20,\" Pass\", \" Fail\")";
    rowNum++;
}
// Save the workbook with the conditional evaluations
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\EvaluatedResults.xlsx");
```

After processing the conditions, you can also fetch and display the formula set for each cell:

```cs
var evaluatedWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\EvaluatedResults.xlsx");
var evaluatedSheet = evaluatedWorkbook.WorkSheets.First();
foreach(var cell in evaluatedSheet ["B1:B4"])
{
    Console.WriteLine($"Cell formula at {cell.Address}: {cell.Formula}"); 
}
Console.ReadKey();
```

In this example, cell formulas are set to evaluate whether the scores in column A meet the pass mark of 20, and the results are saved. Upon opening the `EvaluatedResults.xlsx`, you can see the formulas applied in the 'B' column showing "Pass" or "Fail".

<p class="list-description">Using the same file, we can use the Formula property to set or get a cell’s formula:</p>

<p class="list-decimal">3.7.1. Save to XML “.xml”</p>

```cs
// Setting cell conditions with IF statements
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.FirstSheet();
int index = 1;

foreach(var cell in sheet ["B1:B4"])
{
    // Applying conditional formulas to the cells
    cell.Formula = $"=IF(A{index}>=20, \"Pass\", \"Fail\")";
    index++;
}
// Saving the workbook after modification
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
```

<p class="list-decimal">7.2. Using the generated file from the previous example, we can get the Cell’s Formula:</p>

```cs
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
WorkSheet sheet = workbook.WorkSheets.First();
// Loop through each cell in the specified range
foreach (var cell in sheet ["B1:B4"])
{
    // Output the formula in each cell to the console
    Console.WriteLine(cell.Formula);
}

// Wait for a key press before closing the application
Console.ReadKey();
```

### 3.8. Example of Trimming Cells ###

In this example, we apply the trim function to remove all unnecessary spaces from the cells. For demonstration, I have added a column to the `Sum.xlsx` file as shown below:
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

To utilize the trim function, use the following C# code snippet:

```cs
// Code to apply trim function to cells
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
var sheet = workbook.WorkSheets.First();
int index = 1;
foreach (var cell in sheet["f1:f4"])
{
    cell.Formula = "=trim(D" + index + ")";
    index++;
}
workbook.SaveAs("editedFile.xlsx");
```

This approach demonstrates how you can efficiently use Excel formulas within IronXL to manipulate and clean data directly.

<p class="list-description">To apply trim function (to eliminate all extra spaces in cells), I added this column to the sum.xlsx file</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description">And use this code</p>

Here's your paraphrased section with the relative URL paths resolved:

```cs
// Example on how to apply a trim function to remove extra spaces from cells
var excelBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
var primarySheet = excelBook.WorkSheets.First();
int index = 1;

// Applying the TRIM formula to each cell in the specified range
foreach (var singleCell in primarySheet["f1:f4"])
{
    singleCell.Formula = $"=TRIM(D{index})";
    index++;
}

// Save the modified workbook as a new file
excelBook.SaveAs("editedFile.xlsx");
``` 

I've updated the code comments to make it clearer and modified some variable names for better understanding. The relative paths and functionalities remain consistent with the original content.

<p class="list-description">Thus, you can apply formulas in the same way.</p>

<hr class="separator">

## Multi-Sheet Workbook Management in C# ##

This section explores the handling of Excel workbooks composed of multiple sheets. We'll demonstrate techniques for interacting with workbooks containing several sheets.

### Access and Modify Data Across Multiple Sheets ###

When dealing with an Excel file containing multiple sheets, such as "Sheet1" and "Sheet2," you can direct your operations subtly to manipulate specific sheets within a single workbook.

```cs
/**
Selecting and Working on Different Sheets
anchor-read-data-from-multiple-sheets-in-the-same-workbook
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet2");
var range = sheet ["A2:D2"];
foreach(var cell in range)
{
    Console.WriteLine(cell.Text);
}
```

### Introducing New Sheets to an Existing Workbook ###

It's also straightforward to enhance a workbook by adding new sheets, thus expanding its data structure versatility:

```cs
/**
Introduce New Sheets to Workbook
anchor-add-new-sheet-to-a-workbook
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
var newSheet = workbook.CreateWorkSheet("new_sheet");
newSheet ["A1"].Value = "Hello World";
workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx"); 
```

Here's the paraphrased section with the updated URL paths:

### 4.1. Accessing Data Across Various Sheets in a Workbook ###

In this section, we explore how to retrieve information from multiple sheets within a single Excel workbook using IronXL. For demonstration purposes, I've prepared an `.xlsx` file that includes two distinct sheets named "Sheet1" and "Sheet2."

Previously, we used the method `WorkSheets.First()` to interact with the initial sheet. However, here's how you can explicitly select a sheet by its name and extract data from specified cells:

```cs
/**
Access Specific Sheet by Name
anchor-read-data-from-multiple-sheets-in-the-same-workbook
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet2");
var range = sheet["A2:D2"];
foreach(var cell in range)
{
    Console.WriteLine(cell.Text);
}
```

<p class="list-description">I created an xlsx file that contains two sheets: “Sheet1”,” Sheet2”</p>
<p class="list-description">Until now we used WorkSheets.First() to work with the first sheet. In this example we will specify the sheet name and work with it</p>

Here is the paraphrased section of the article with updated relative URL paths resolved to `ironsoftware.com`:

```cs
// Code to demonstrate accessing multiple worksheets within a workbook
// Topic: Reading data from multiple sheets within the same workbook
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet selectedSheet = workbook.GetWorkSheet("Sheet2"); // Access the second worksheet
Range selectedRange = selectedSheet["A2:D2"]; // Define the cell range

// Loop through each cell in the specified range
foreach(var cell in selectedRange)
{
    Console.WriteLine(cell.Text); // Print the text content of each cell
}
```

### 4.2. Incorporating a New Sheet into Your Workbook ###

Adding a new worksheet to your workbook is a straightforward process with IronXL. Here’s how you can do it:

```cs
// Load an existing workbook
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create and add a new worksheet named 'new_sheet'
var newSheet = workbook.CreateWorkSheet("new_sheet");

// Set the value of cell A1 to "Hello World"
newSheet ["A1"].Value = "Hello World";

// Save the workbook with the new sheet included
workbook.SaveAs("https://ironsoftware.com/F/MY WORK/IronPackage/Xl tutorial/newFile.xlsx");
```

In this example, we first open an existing workbook. We then generate a new worksheet within that workbook, name it `new_sheet`, and input data into it. Finally, the workbook with the newly added sheet is saved, preserving all modifications.

<p class="list-description">We can also add new sheet to a workbook:</p>

Below is the paraphrased section:

```cs
// Add a new worksheet
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
var sheet = workbook.CreateWorkSheet("new_sheet");
sheet ["A1"].Value = "Hello World";
// Save the workbook with the new sheet
workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx"); 
```

<hr class="separator">

## 5. Integrating with Excel Database ##

In this section, we explore how to import data from and export data to a database using IronXL.

For demonstration purposes, I've established a database named "TestDb" that includes a table called `Country`. This table is structured with two columns: `Id` (an integer with identity specification) and `CountryName` (a string).

This setup allows us to efficiently handle data integration tasks between Excel files and our database, leveraging the capabilities of IronXL to manipulate Excel data programmatically.

### 5.1. Populating an Excel Sheet from Database Content ###

In this segment, we explore the process of transferring data from a database into a new Excel sheet. The procedure involves creating a new sheet in an Excel workbook and filling it with data extracted from the `Country` table in the database named "TestDb," which includes columns for Id (int, identity) and CountryName (string).

```cs
/**
Import Data to Sheet
anchor-fill-exel-sheet-from-db-data
**/
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.CreateWorkSheet("FromDb");
List<Country> countryList = dbContext.Countries.ToList();
sheet.SetCellValue(0, 0, "Id");
sheet.SetCellValue(0, 1, "CountryName");
int row = 1;
foreach (var item in countryList)
{
    sheet.SetCellValue(row, 0, item.id);
    sheet.SetCellValue(row, 1, item.CountryName);
    row++;
}
workbook.SaveAs("FilledFile.xlsx");
```

This method initializes a connection to the `TestDb` database, locates the specific Excel file, and creates a new worksheet within it. It then retrieves a list of countries stored in the database and iteratively populates the newly created worksheet. Each row of the Excel sheet is filled with the `id` and `CountryName` from the database's `Country` table. Finally, the file is saved with all the new data neatly organized in a worksheet labeled "FromDb".

<p class="list-description">Here we will create a new sheet and fill it with data from the Country Table</p>

Here's the paraphrased section with resolved URL paths:

```cs
// Initialize database connection
TestDbEntities dbContext = new TestDbEntities();

// Load an existing workbook and create a new worksheet for database data
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet worksheet = workbook.CreateWorkSheet("DatabaseData");

// Retrieve country data from the database
List<Country> countries = dbContext.Countries.ToList();

// Set headers for the columns in the new worksheet
worksheet.SetCellValue(0, 0, "Id");
worksheet.SetCellValue(0, 1, "CountryName");

// Populate the worksheet with data from the database
int currentRow = 1;
foreach (Country country in countries)
{
    worksheet.SetCellValue(currentRow, 0, country.id);
    worksheet.SetCellValue(currentRow, 1, country.CountryName);
    currentRow++;
}

// Save the modified workbook with data from the database
workbook.SaveAs("UpdatedExcelFile.xlsx");
```

This code initializes a database context and loads an Excel workbook from the specified directory. It then creates a new worksheet, retrieves a list of countries from the database, and populates the worksheet with these details. Finally, it saves the workbook to a new file.

### 5.2 Transfer Data from Excel Sheet to Database ###

<p class="list-description">Learn how to populate your database using data extracted from an Excel worksheet</p>

```cs
/**
Data Import to Database
anchor-import-data-from-excel-to-database
**/
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
DataTable dataTable = sheet.ToDataTable(true);
foreach (DataRow row in dataTable.Rows)
{
    Country c = new Country();
    c.CountryName = row [1].ToString();
    dbContext.Countries.Add(c);
}
dbContext.SaveChanges();
```

In this section, an Excel file is utilized to populate a database table. Here, we first establish a connection to our database context. Next, we load an Excel workbook from a specific directory and specifically target a worksheet named 'Sheet3'. We then convert this sheet into a DataTable format, facilitating the extraction of rows. Each row's data is used to create a new instance of the `Country` entity, which is subsequently added to the database. The changes are finalized and committed using the `SaveChanges` method, effectively updating the database with the new data extracted from the Excel sheet.

<p class="list-description">Insert the data to the Country table in TestDb Database</p>

Here's a paraphrased version of the provided C# code snippet for importing data from an Excel sheet to a database:

```cs
// Importing data from an Excel file into a database
TestDbEntities dataContext = new TestDbEntities();
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet targetSheet = excelWorkbook.GetWorkSheet("Sheet3");
System.Data.DataTable recordsTable = targetSheet.ToDataTable(true);

// Iterating through each data row and adding to the database
foreach (DataRow record in recordsTable.Rows)
{
    Country newCountry = new Country();
    newCountry.CountryName = record[1].ToString();  // Assuming column 1 holds the country name
    dataContext.Countries.Add(newCountry);
}
dataContext.SaveChanges();  // Commit the changes to the database
```

<hr class="separator">

### Additional Resources

Explore further into IronXL by reviewing the various tutorials available in this section and by examining the examples featured on our homepage, which are detailed and comprehensive enough for most developers to begin their projects.

For detailed documentation on the `WorkBook` class, refer to our [API Reference](https://ironsoftware.com/csharp/excel/object-reference/).

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

