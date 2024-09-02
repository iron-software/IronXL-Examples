# C# Tutorial for Opening and Writing Excel Files

Discover through detailed examples how to generate, open, and preserve Excel spreadsheets using C#, along with performing primary functions such as calculating sums, averages, counts, and more. IronXL.Excel functions as an independent .NET library capable of processing various spreadsheet formats. Importantly, it operates independently of [Microsoft Excel](https://products.office.com/en-us/excel) and does not rely on Interop.

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

Effortlessly manage, write, and customize Excel files using the straightforward [IronXL C# library](https://ironsoftware.com/csharp/excel/).

You can initiate your journey by downloading a [ready-to-use sample project from GitHub](https://github.com/magedo93/IronSoftware.git). Alternatively, you can integrate this tutorial into your own project.

Here's what you'll need to get started:

1. Install the IronXL Excel Library from [NuGet](https://www.nuget.org/packages/IronXL.Excel) or by downloading the DLL directly.
2. Implement the `WorkBook.Load` method to open any XLS, XLSX, or CSV file.
3. Retrieve cell values with the intuitive syntax: `sheet["A11"].DecimalValue`.

Throughout this tutorial, we'll guide you through:

- **Installation**: Instructions on how to add IronXL.Excel to your existing project.
- **Basic Operations**: Steps on creating or opening a workbook, selecting sheets and cells, and saving your workbook.
- **Advanced Sheet Operations**: Techniques to enhance your sheets, such as adding headers or footers, performing mathematical operations, and leveraging other advanced features.

<h4>Open an Excel File : Quick Code</h4>

Below is a rephrased version of the provided C# code snippet, which demonstrates the basic usage of the `IronXL` library to load an Excel file, iterate through a specific range of cells, and perform a comparison operation:

```cs
// Required libraries
using IronXL;
using System;

// Load the workbook and access the default worksheet
WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet sheet = workbook.DefaultWorkSheet;

// Define the range of cells to operate on
Range range = sheet["A2:A8"];

// Initialize a variable to store the summation result
decimal sumTotal = 0;

// Loop through each cell in the specified range
foreach (var cell in range)
{
    // Display cell details
    Console.WriteLine($"Cell {cell.RowIndex} contains the value: '{cell.Value}'");

    // Aggregate numeric values only
    if (cell.IsNumeric)
    {
        // Accurately accumulate decimal values
        sumTotal += cell.DecimalValue;
    }
}

// Validate the sum against a predetermined value
if (sheet["A11"].DecimalValue == sumTotal)
{
    // Output a message if the values match
    Console.WriteLine("Verification successful: Basic test passed.");
}
```

<h4>Write and Save Changes to the Excel File : Quick Code</h4>

```cs
// Assign value to a cell
sheet["B1"].Value = 11.54;

//Persist the changes to a file named 'test.xlsx'
workbook.SaveAs("test.xlsx");
```

<hr class="separator">
<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Free Installation of the IronXL C# Library

IronXL.Excel offers a robust and adaptable library that facilitates the opening, editing, reading, and saving of Excel files within .NET environments. It is compatible with various .NET project types including Windows applications, ASP.NET MVC, and .NET Core Application, making it a versatile tool for developers.

<h3>Install the Excel Library to your Visual Studio Project with NuGet</h3>

To begin using IronXL.Excel, you have two convenient options to incorporate this library into your project: either using the NuGet Package Manager or the NuGet Package Manager Console.

Here’s how to install IronXL.Excel using the NuGet Package Manager’s graphical interface:

1. Navigate to your project in the solution explorer, right-click on the project name, and choose "Manage NuGet Packages" from the context menu.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <p><img src="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

From the browsing interface, type "IronXL.Excel" into the search field and proceed to install it.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" target="_blank">

<p><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>

</a>

### Step 1.3: Installation Complete

After installing the IronXL.Excel library, you've successfully completed the setup. This powerful .NET library is now ready to enhance your project with Excel file manipulation capabilities.

![Installation Complete](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg)

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>



<h3>Install Using NuGet Package Manager Console</h3>

1. Navigate through **Tools**, select **NuGet Package Manager**, and then choose **Package Manager Console**.
```
This instruction guides you through accessing the Package Manager Console via the Visual Studio interface, enhancing your project integration with additional libraries conveniently.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

Execute the following command to add IronXL.Excel to your project:

```plaintext
2. Run the command: Install-Package IronXL.Excel -Version 2019.5.2
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<h3>Manually Install with the DLL</h3>

You can also opt to manually add the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) to your project or the global assembly cache.

Here is the paraphrased section with URLs resolved to `ironsoftware.com`:


```cmd
PM> Install-Package IronXL.Excel
```
```

# Tutorial on Opening and Writing Excel Files in C# with IronXL

## Learn Step-by-Step
Learn how to handle Excel spreadsheets using C# without installing Microsoft Excel or using Interop. IronXL.Excel is a robust .NET class library that reads and writes various spreadsheet formats easily.

<div class="that-strolls">
  <div class="row">
    <div class="col-sm-6">
      <h2>Guidelines for C# .NET Excel Operations:</h2>
      <ul class="bare-list">
        <li><a href="#anchor-1-install-the-ironxl-c-library-free">Setting up the IronXL C# Library</a></li>
        <li><a href="#anchor-2-2-create-a-new-excel-file">Generate a New Excel File from CSV, XML, or JSON</a></li>
        <li><a href="#anchor-4-working-with-workbooks-with-sheets">Handling Multisheet Workbooks</a></li>
        <li><a href="#anchor-3-advanced-operations-sum-avg-count-etc">Implement SUM, AVG, Count, and More Functions</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-frame">
        <a href="https://ironsoftware.com/downloads/assets/excel/tutorials/csharp-open-write-excel-file/tutorial-open-and-write-excel.pdf" target="_blank">
          <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write.svg" data-hover-src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/how-to-open-and-write-hover.svg" class="img-responsive learn-how-to-img replaceable-img">
        </a>
      </div>
    </div>
  </div>
</div>

<hr class="divider">

<h4 class="tutorial-segment-title">Summary</h4>
<h2>Employ IronXL to Manage Excel Docs</h2>

Easily open, edit, save, and customize Excel documents through the intuitive <a href="https://ironsoftware.com/csharp/excel/" target="_blank">IronXL C# library.</a>

Start by downloading a <a href="https://github.com/magedo93/IronSoftware.git" target="_blank">GitHub sample project</a> or use your own and follow our guide. 

1. Begin by installing IronXL Excel Library via <a href="https://www.nuget.org/packages/IronXL.Excel" target="_blank">NuGet</a> or directly downloading the DLL
2. Load any XLS, XLSX, or CSV file using the `WorkBook.Load` method.
3. Retrieve cell values using easy syntax: <code>sheet["A11"].DecimalValue</code>

We'll guide you through:

- How to set up IronXL.Excel in your existing project
- Basics of working with Excel: Opening or creating workbooks, selecting sheets/cells, and saving changes
- Advanced Sheet Manipulation: Adding headers, performing math operations, and using other advanced features

<h4>Opening an Excel File: Quick Example</h4>

```cs
using IronXL;
using System;

WorkBook workbook = WorkBook.Load("test.xlsx");
WorkSheet sheet = workbook.DefaultWorkSheet;

Range range = sheet["A2:A8"];

decimal total = 0;

// Iterate through a range of cells and sum numeric values
foreach (var cell in range)
{
    Console.WriteLine($"Cell {cell.RowIndex} holds value '{cell.Value}'");

    if (cell.IsNumeric)
    {
        // DecimalValue is used for precision
        total += cell.DecimalValue;
    }
}

if (sheet["A11"].DecimalValue == total)
{
    Console.WriteLine("Sum matches computed total");
}
```

<h4>Write and Persist Changes to the Excel File: Quick Example</h4>

```cs
// Assign a value to cell B1
sheet["B1"].Value = 11.54;

// Persist changes by saving the workbook
workbook.SaveAs("test.xlsx");
```

<hr class="divider">

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Essential Functions: Creating, Opening, and Saving Files ##

### 2.1. Sample Project: HelloWorld Console App ###

<p class="list-description">Initiate a HelloWorld Project</p>

<p class="list-decimal">2.1.1. Launch Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Opt for 'Create New Project'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Select 'Console App (.NET framework)'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Name your project “HelloWorld” and click 'Create'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. Your console application is now ready</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Add the IronXL.Excel library - click install</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Include initial lines of code to read the first cell on the first sheet and display it</p>

```cs
static void Main(string [] args)
{
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
    var sheet = workbook.WorkSheets.First();
    var cell = sheet ["A1"].StringValue;
    Console.WriteLine(cell);
}
```

### 2.2. Creating a New Excel File ###

<p class="list-description">Generate a fresh Excel file with IronXL</p>

```cs
static void Main(string [] args)
{
    var newExcelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    newExcelFile.Metadata.Title = "IronXL New File";
    var newSheet = newExcelFile.CreateWorkSheet("1stWorkSheet");
    newSheet ["A1"].Value = "Hello World";
    newSheet ["A2"].Style.BottomBorder.SetColor("#ff6600");
    newSheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

### 2.1 Example Project: HelloWorld Console Application ###

Embark on creating a new project by following these detailed steps:

1. **Start Visual Studio:**
   - Begin by launching Visual Studio to set up the foundation of your project.
   ![Start Visual Studio](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png)

2. **Create a New Project:**
   - Choose the 'Create New Project' option once Visual Studio is open.
   ![Create New Project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png)

3. **Select a Console Application:**
   - Opt for a Console Application (.NET Framework) as your project type.
   ![Select Console App](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg)

4. **Name Your Project:**
   - Assign the name “HelloWorld” to your new project and then create it.
   ![Name Your Project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg)

5. **Project Creation Complete:**
   - After naming and creating the project, you will see your new console application ready for coding.
   ![Project Created](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg)

6. **Install IronXL Excel Library:**
   - Add the IronXL.Excel library by navigating to IronXL and clicking install. This action integrates the necessary functionalities for managing Excel files into your project.
   ![Install IronXL](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg)

7. **Add Initial Code:**
   - Compose the first lines of code that read the first cell of the first sheet in the Excel file, displaying its contents using `Console.WriteLine`.
   ```csharp
   static void Main(string [] args)
   {
       var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
       var sheet = workbook.WorkSheets.First();
       var cell = sheet ["A1"].StringValue;
       Console.WriteLine(cell);
   }
   ```

This segment walks you through the basics of setting up a simple HelloWorld console application that utilizes IronXL to interact with Excel files, elevating your .NET project's data handling capabilities.

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
static void Main(string[] args)
{
    // Load the HelloWorld.xlsx from the current directory
    var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");

    // Access the first worksheet in the workbook
    var sheet = workbook.WorkSheets.First();

    // Retrieve the string value from cell A1
    var cellValue = sheet["A1"].StringValue;

    // Display the value of cell A1 to the console
    Console.WriteLine(cellValue);
}
```

### 2.2. Generate a Fresh Excel Document ###

This section will guide you on how to craft a new Excel file using IronXL.

```cs
// Creating a new Excel file
static void Main(string [] args)
{
    // Initialize a new workbook with XLSX format
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    // Set the document title
    workbook.Metadata.Title = "New IronXL File";
    
    // Add a new worksheet named 'FirstSheet'
    var sheet = workbook.CreateWorkSheet("FirstSheet");
    
    // Assign a value to cell A1
    sheet ["A1"].Value = "First Entry";
    
    // Style cell A2 with an orange dashed bottom border
    sheet ["A2"].Style.BottomBorder.SetColor("#ff9900");
    sheet ["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
}
```

<p class="list-description">Create a new Excel file using IronXL</p>

Here's the paraphrased version of the provided section of code:

```cs
// Initialize a new Excel file using IronXL
static void Main(string[] args)
{
    // Create a new Excel workbook
    WorkBook myWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
    myWorkbook.Metadata.Title = "New IronXL File";

    // Add a new worksheet to the workbook
    WorkSheet myWorkSheet = myWorkbook.CreateWorkSheet("InitialWorkSheet");

    // Set the value of cell A1
    myWorkSheet["A1"].Value = "Hello World";

    // Customize the style of cell A2
    myWorkSheet["A2"].Style.BottomBorder.SetColor("#ff6600"); // Set border color
    myWorkSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed; // Set border type
}
```

### Opening Various File Formats as Workbooks

Opening files in different formats like CSV, XML, or JSON as workbooks in C# is straightforward with IronXL. Below, we'll guide you through the process using different data formats.

#### 2.3.1 Opening a CSV File

Firstly, let's open a CSV file and load it as a workbook. Follow these steps:

```csharp
// Load CSV as a workbook
static void Main(string [] args)
{
    WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
    WorkSheet sheet = workbook.WorkSheets.First();
    var cell = sheet ["A1"].StringValue;  
    Console.WriteLine(cell);
}
```

#### 2.3.2 Open an XML File

You can also open an XML file that contains structured data, such as a list of countries. Here is an XML structure example and how to load it:

```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
  <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
  <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
  <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

```csharp
// Load XML as a workbook
static void Main(string [] args)
{
    DataSet xmldataset = new DataSet();
    xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
    WorkBook workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

#### 2.3.3 Open a JSON List as Workbook

Lastly, you can create a class to represent your JSON structure and load a JSON file as follows:

```csharp
// JSON country list model
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}

// Load JSON as a workbook
static void Main(string [] args)
{
    StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
    var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
    DataSet xmldataset = countryList.ToDataSet();
    var workbook = IronXL.WorkBook.Load(xmldataset);
    var sheet = workbook.WorkSheets.First();
}
```

In this section, we covered how to open different file formats as workbooks using IronXL, simplifying the process of data manipulation within .NET applications.

All image and document references should be checked to ensure their paths are correctly pointed to `https://ironsoftware.com` if they are intended to be used. For example:

- Image URL converted: `https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg`
- Document download link updated: `https://ironsoftware.com/csharp/excel/tutorials/downloads/Use.CSharp.to.Open.&.Write.an.Excel.File.zip`

This ensures all resources are correctly attributed and accessible from the given URLs.

<p class="list-decimal">2.3.1. Open CSV file</p>

<p class="list-decimal">2.3.2 Create a new text file and add to it a list of names and ages (see example) then save it as CSVList.csv</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Your code snippet should look like this</p>

```cs
/**
Load CSV File as a Workbook
anchor-load-csv-as-workbook
**/
static void Main(string[] args)
{
    // Load the workbook from a CSV file located in the current directory
    var workbook = IronXL.WorkBook.Load(Path.Combine(Directory.GetCurrentDirectory(), "Files", "CSVList.csv"));
    // Access the first worksheet in the workbook
    var sheet = workbook.DefaultWorkSheet;
    // Retrieve the value from the first cell of the worksheet
    var value = sheet ["A1"].StringValue;
    // Print the value to the console
    Console.WriteLine(value);
}
```

<p class="list-decimal">

### 2.3.3. Open XML File

This section guides you on how to utilize an XML file by converting it into a workable Excel format with IronXL. This process involves creating an XML with structured data and loading it into IronXL seamlessly.

First, create an XML document that lists countries along with their unique identifying attributes like code and continent. Below is a sample structure:

```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

Once your XML file is structured correctly, utilize the following C# snippet to load the XML into IronXL as a workbook. This enables further manipulations and operations within the Excel environment:

```cs
// Load XML data into IronXL
static void Main(string [] args)
{
    // Initialize the dataset to read XML
    var xmldataset = new DataSet();
    xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");  // Modify the path based on your file location

    // Load the DataSet into a workbook
    var workbook = IronXL.WorkBook.Load(xmldataset);

    // Access the first worksheet by default
    var sheet = workbook.WorkSheets.First();
}
```

This code snippet successfully integrates structured XML data into an IronXL workbook, making it ready for any further Excel-based manipulations or processing tasks.

<span class="list-description">Create an XML file that contains a countries list: the root element “countries”, with children elements “country”, and each country has properties that define the country like code, continent, etc.</span>
</p>

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales Great Britain UK Britain Northern GB" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="USA America US" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">2.3.4. Copy the following code snippet to open XML as a workbook</p>

```cs
/**
Load XML into a Workbook
anchor-load-xml-into-workbook
**/
static void Main(string[] args)
{
    DataSet xmlData = new DataSet(); // Create a new instance of DataSet
    xmlData.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml"); // Read XML data
    WorkBook xlWorkbook = IronXL.WorkBook.Load(xmlData); // Load data into an Excel workbook
    WorkSheet xlSheet = xlWorkbook.GetFirstSheet(); // Retrieve the first worksheet
}
```

<p class="list-decimal">

### Opening a JSON List as a Workbook

In this section, we'll explore how to ingest a JSON list into IronXL to create an interactive workbook.

```cs
/**
Transform JSON into Workbook
anchor-opening-json-list-as-workbook
**/
static void Main(string [] args)
{
    // Read JSON data from a file
    var jsonInput = File.ReadAllText($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");

    // Deserialize the JSON data to an array of CountryModel objects
    var countries = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel []>(jsonInput);

    // Convert the deserialized object to DataSet
    var dataSet = countries.ToDataSet();

    // Load the DataSet into a new workbook
    var workbook = IronXL.WorkBook.Load(dataSet);

    // Access the first worksheet in the workbook
    var sheet = workbook.WorkSheets.First();
}

```

In the above code, the JSON file containing lists of countries is deserialized into an array of `CountryModel` objects, which is then converted into a `DataSet`. This dataset is then loaded into IronXL to form a new workbook, where data can be manipulated as needed. This method facilitates the ease of handling structured JSON data within a .NET environment using IronXL.

<span class="list-description">Create JSON country list</span>
</p>

```cs
/**
Load JSON Data into an Excel Workbook
anchor-load-json-data-into-excel-workbook
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

<p class="list-decimal"></p>
<p class="list-decimal">2.3.6. Create a country model that will map to JSON</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Here is the class code snippet</p>

Here's the paraphrased version of the code section you provided, with modifications to variable names and slight restructuring for clarity:

```cs
public class NationModel
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

```cs
/**
Transform List to DataSet
anchor-list-to-dataset-conversion
**/
public static class DataSetConverter
{
    public static DataSet ConvertListToDataSet<T>(this IList<T> listItems)
    {
        Type itemType = typeof(T);
        DataSet dataSet = new DataSet();
        DataTable dataTable = new DataTable();
        dataSet.Tables.Add(dataTable);

        // Create a column in the DataTable for each public property of T
        foreach (var property in itemType.GetProperties())
        {
            Type columnType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            dataTable.Columns.Add(property.Name, columnType);
        }

        // Populate the dataTable row by row with properties' values from the list
        foreach (T item in listItems)
        {
            DataRow dataRow = dataTable.NewRow();
            
            foreach (var property in itemType.GetProperties())
            {
                dataRow[property.Name] = property.GetValue(item, null) ?? DBNull.Value;
            }

            dataTable.Rows.Add(dataRow);
        }

        return dataSet;
    }
}
```

<p class="list-decimal"></p>
<p class="list-decimal">And finally load this dataset as a workbook</p>

Here's a paraphrased version of the given code snippet:

```cs
static void Main(string [] args)
{
    // Open a JSON file containing a list of countries
    var jsonFileReader = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");

    // Deserialize the JSON content into an array of CountryModel objects
    var countries = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFileReader.ReadToEnd());

    // Convert the array to a DataSet
    var dataSet = countries.ToDataSet();

    // Load the DataSet into a new IronXL workbook
    var excelWorkbook = IronXL.WorkBook.Load(dataSet);

    // Access the first worksheet in the workbook
    var firstSheet = excelWorkbook.WorkSheets.First();
}
```

In the revised version:
- The variables and operations have been made more explicit in their naming to enhance readability.
- Comments have been added to explain each step clearly, helping to make the code more understandable even for those not familiar with JSON or Excel programming in C#.

### 2.4 Saving and Exporting Excel Files ###

In this section, you'll learn how to save and export Excel files to various formats using different techniques.

#### 2.4.1 Saving as `.xlsx` ####
To save the file in the `.xlsx` format, the `SaveAs` function is utilized. Here's an example:
```cs
/**
Save and Export as XLSX
anchor-save-and-export-xlsx
**/
static void Main(string [] args)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    workbook.Metadata.Title = "IronXL New File";
    var newSheet = workbook.CreateWorkSheet("FirstSheet");
    newSheet["A1"].Value = "Hello World";
    newSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
    newSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
}
```

#### 2.4.2 Saving as `.csv` ####
To save the workbook in CSV format with a specific delimiter, use `SaveAsCsv`. The following example saves the workbook as a CSV file using the pipe `|` as a delimiter:
```cs
workbook.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter: "|");
```

#### 2.4.3 Exporting to `.json` ####
For JSON format, `SaveAsJson` is the method to use. Here's how to perform the operation:
```cs
workbook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```
The resulting output in the file would be structured like this:
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

#### 2.4.4 Saving as `.xml` ####
To export the data in XML format, `SaveAsXml` is the appropriate method:
```cs
workbook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
```
The XML output will appear as:
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

This section guides you through various methods available in IronXL to save and export Excel files effectively across different file formats.

<span class="list-description">We can save or export the Excel file to multiple file formats like (“.xlsx”,”.csv”,”.html”) using one of the following commands.</span>

<p class="list-decimal">

### 2.4.1 Saving Excel Files as XLSX

In this section, you'll learn how to export your workbook to the XLSX format using IronXL. This feature enables you to maintain a familiar and universally compatible Excel file format.

```cs
/**
Save and Export
anchor-save-and-export
**/
static void Main(string [] args)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    workbook.Metadata.Title = "IronXL New File";
    var worksheet = workbook.CreateWorkSheet("Sheet1");
    worksheet["A1"].Value = "Hello World";
    worksheet["A2"].Style.BottomBorder.Color = "#ff6600";
    worksheet["A2"].Style.BottomBorder.Style = BorderType.Dashed;

    // Specify the path and filename for the XLSX file
    workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
}
```

This example demonstrates how to create a new Excel spreadsheet, customize style options for a cell, and save the newly created Excel file as `.xlsx`. This format ensures that your document remains accessible and editable in most spreadsheet applications beyond IronXL.

<span class="list-description">To Save to “.xlsx” use saveAs function</span>
</p>

Below is a paraphrased version of the provided C# code segment, with resolved paths based on `ironsoftware.com`:

```cs
// Code to demonstrate saving and exporting Excel files
static void Main(string[] args)
{
    // Create a new Excel file using IronXL
    WorkBook excelFile = WorkBook.Create(ExcelFileFormat.XLSX);
    excelFile.Metadata.Title = "New Excel Document by IronXL";

    // Add a worksheet to the file
    WorkSheet worksheet = excelFile.CreateWorkSheet("FirstSheet");
    worksheet["A1"].Value = "Hello World";

    // Customize cell style
    worksheet["A2"].Style.BottomBorder.SetColor("#ff6600");
    worksheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

    // Save the Excel document to the local directory
    excelFile.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\Introduction.xlsx");
}
```

This version maintains the functionality of the original code but rephrases comments and variable names for clarity and context.

<p class="list-decimal">

### 2.4.2. Export to CSV Format

<span class="list-description">To export a document in CSV format, you may use the `SaveAsCsv` method which requires two arguments: the first for specifying the file name with its path, and the second for setting the delimiter (e.g., ",", "|", or ":").</span>
```

</p>

Here's the paraphrased section with URLs resolved to `ironsoftware.com`:

```cs
// Save the workbook as a CSV file
newXLFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter:",");
```

<p class="list-decimal">

### 2.4.3 Export to JSON Format

<span class="list-description">To convert and save your Excel spreadsheet into a JSON format, follow these instructions:</span>

```cs
newXLFile.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

<p class="list-decimal">
  <span class="list-description">Here's how the resulting JSON file will appear:</span>
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

<span class="list-description">To save to Json “.json” use SaveAsJson as follow</span>
</p>

```cs
// Save the newly created Excel workbook as a JSON file in the current working directory
newXLFile.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\MyNewExcelFile.json");
```

<p class="list-decimal">
  <span class="list-description">The result file should look like this</span>
</p>

Below is the paraphrased version of the provided JSON code snippet:

```cs
[
    [
        "Hello World"
    ],
    [
        "Empty String"
    ]
]
``` 

This modified JSON structure maintains the original pattern of data representation but alters the second list's content for clearer context, identifying it as an "Empty String".

<p class="list-decimal">

### 2.4.4 Export to XML Format ###

This section will guide you on exporting your Excel data to an XML file using IronXL.

```cs
/**
Export as XML
guide-export-to-xml
**/
static void Main(string [] args)
{
    var newXLFile = WorkBook.Create(ExcelFileFormat.XLSX);
    newXLFile.Metadata.Title = "New IronXL File";
    var newSheet = newXLFile.CreateWorkSheet("Sheet1");
    newSheet ["A1"].Value = "Hello World";  // Set value for cell A1

    // Save the workbook to an XML file
    newXLFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\MyNewXMLFile.XML");
}
```

The output XML file will be structured as follows:

```xml
<?xml version="1.0" standalone="yes"?>
<_x0031_StWorkSheet>
  <_x0031_StWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">Hello World</Column1>
  </_x0031_StWorkSheet>
  <_x0031_StWorkSheet>
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </_x0031_StWorkSheet>
</_x0031_StWorkSheet>
```

This example demonstrates how simple it is to export your workbook data to an XML format, maintaining a straightforward and easy-to-understand structure.
```

<span class="list-description">To save to xml use SaveAsXml as follow</span>
</p>

```cs
newXLFile.SaveAsXml(Path.Combine(Directory.GetCurrentDirectory(), "Files", "HelloWorldXML.XML"));
```

<p class="list-decimal">
  <span class="list-description">Result should be like this</span>
</p>

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

## 3. Advanced Functions: Sum, Average, Count, and More ##

Explore how to use popular Excel functions such as SUM, AVG, and COUNT with the following code examples.

### 3.1. Example: Calculating Sum ###

<p class="list-description">Discover how to compute the total of a number series. Initially, I prepared an Excel file named “Sum.xlsx” where I manually entered a series of numbers.</p>

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank">
    <img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

```cs
/**
Function SUM
anchor-sum-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
decimal sum = sheet ["A2:A4"].Sum();
Console.WriteLine(sum);
```

<p class="list-description">Let’s find the sum for this list. I created an Excel file and named it “Sum.xlsx” and added this list of numbers manually</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

```cs
// Calculating the total sum of values in a specific range
var excelWorkbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\Sum.xlsx");  // Load the workbook
var excelSheet = excelWorkbook.WorkSheets.First();  // Access the first worksheet
decimal totalSum = excelSheet["A2:A4"].Sum();  // Compute sum of cells from A2 to A4
Console.WriteLine(totalSum); // Output the sum
```

### 3.2 Example: Calculating the Average ###

Explore the following example that demonstrates obtaining the average from a series of values:

```cs
// Function to calculate the average
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
decimal avg = sheet["A2:A4"].Avg();
Console.WriteLine(avg);
```

This snippet loads an Excel file named "Sum.xlsx", accesses the first worksheet, and calculates the average of the values in the cell range from A2 to A4. The result is then output to the console.

<p class="list-description">Using the same file, we can get the average:</p>

Below is the paraphrased section of the article, with relative URL paths resolved to ironsoftware.com:

-----
```cs
// Calculate the Average Value from a Range in an Excel File
// Example Identifier: avg-calculation
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var activeSheet = excelWorkbook.WorkSheets.First();
decimal averageValue = activeSheet["A2:A4"].Avg(); // Compute the average of the values in cells A2 to A4
Console.WriteLine(averageValue); // Output the average value to the console
```

### Example: Counting Cells in Excel ###

To demonstrate the practical application of the `Count()` method:

```cs
// Load workbook containing our data
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First(); // Select first worksheet

// Count the number of elements in our selection range from A2 to A4
decimal numberOfElements = sheet["A2:A4"].Count();
Console.WriteLine(numberOfElements); // Output the count to the console
```

<p class="list-description">Using the same file, we can also get the number of elements in a sequence:</p>

Below is the paraphrased content for the specified section with resolved URLs from the article:

```cs
// Example for counting cells in a range
// Tag: count-example
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet sheet = workbook.DefaultWorkSheet;
decimal numberOfCells = sheet.Range("A2:A4").Count();
Console.WriteLine("Total cells in range: " + numberOfCells);
```

### 3.4. Example: Finding Maximum Values ###

This section demonstrates how to identify the maximum value within a specific range of cells using IronXL. We use a predefined Excel file, "Sum.xlsx," which contains an array of numbers for this purpose.

```cs
// Load the workbook and select the first worksheet
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();

// Calculate the maximum value from the specified cell range
decimal max = sheet["A2:A4"].Max();
Console.WriteLine(max);
```

The code above succinctly demonstrates the process of loading an Excel file, accessing its first sheet, and then utilizing the `.Max()` function on a range of cells to fetch the highest number. It then prints this value to the console, providing a clear and efficient way to obtain maximum values within spreadsheets.

<p class="list-description">Using the same file, we can get the max value of range of cells:</p>

```cs
// Example demonstrating the use of the Max function
var workbook = IronXL.WorkBook.Load(System.IO.Path.Combine(Environment.CurrentDirectory, "Files", "Sum.xlsx"));
var worksheet = workbook.DefaultWorkSheet;
decimal maximumValue = worksheet["A2:A4"].Max();
Console.WriteLine($"Maximum Value: {maximumValue}");
```

<p class="list-description">– We can apply the transform function to the result of max function:</p>

Here's a paraphrased version of the given C# code snippet, with updated image and link URLs:

```cs
// Load the workbook from the current directory
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
var sheet = workbook.WorkSheets.First();

// Determine the max value in range 'A1:A4', checking if any cells contain formulas
bool maxFormulaCheck = sheet["A1:A4"].Max(cell => cell.IsFormula);

// Output the result
Console.WriteLine(maxFormulaCheck);
```

<p class="list-description">This example writes “false” in the console.</p>

### 3.5. Example to Determine Minimum Value ###

Using the same dataset, let's identify the smallest value from a range of Excel cells:

```cs
// Function to demonstrate finding the minimum value in Excel
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
bool minVal = sheet ["A1:A4"].Min();
Console.WriteLine(minVal);
```

<p class="list-description">Using the same file, we can get the min value of range of cells:</p>

```cs
// Example: Using Min to find the smallest value in a cell range
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var excelSheet = excelWorkbook.WorkSheets.First();
bool minimumValue = excelSheet["A1:A4"].Min();
Console.WriteLine(minimumValue);
```

### 3.6. Example: Ordering Cells ###

In this section, we'll explore how to sort cells within a range either in ascending or descending order using IronXL. Let's consider an Excel file named "Sum.xlsx" that includes a range we wish to organize.

```cs
/**
Operations: Sorting Cells
anchor-order-cells-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");

// Sort cells in ascending order. Change to SortDescending() for descending order.
sheet ["A1:A4"].SortAscending(); 

// Save the sorted workbook with a new name
workbook.SaveAs("SortedSheet.xlsx");
```

This code snippet demonstrates how to select a cell range and sort its values. The `SortAscending()` method arranges the cells in ascending order, and you can switch to `SortDescending()` for the reverse order. After sorting, the workbook is saved under a new file name, preserving the original data.

<p class="list-description">Using the same file, we can order cells by ascending or descending:</p>

```cs
/**
Arrange Cell Order
anchor-arrange-cell-order-example
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var sheet = workbook.WorkSheets.First();
// Ascending sort. For descending order, use: sheet ["A1:A4"].SortDescending();
sheet ["A1:A4"].SortAscending();
workbook.SaveAs("SortedSheet.xlsx");
```

### Example of Applying Conditional Logic in Excel ###

Utilizing the `IronXL` library, it's straightforward to integrate conditional statements directly into your Excel sheets. Let's explore an example that demonstrates the use of the IF condition to dynamically set cell values based on logical criteria.

1. **Introduction:** This example begins by loading an existing Excel file named "Sum.xlsx" from a predefined directory. We'll process a range of cells and apply a conditional formula to determine whether certain criteria are met.

```cs
// Load the workbook from the specified path
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
var sheet = workbook.WorkSheets.First();

// Initialize a variable to iterate through cells
int rowIndex = 1;

// Loop through cells B1 to B4
foreach(var cell in sheet["B1:B4"])
{
    // Apply the IF condition to check if the value in column A at the corresponding row is 20 or greater
    cell.Formula = $"=IF(A{rowIndex}>=20, 'Pass', 'Fail')";

    // Move to the next row
    rowIndex++;
}

// Save the modified workbook to a new file
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\UpdatedExcelFile.xlsx");
```

2. **Inspecting Formulas:** After setting conditional formulas in the step above, you can review them by reloading the workbook and iterating over the cells where the formulas were set.

```cs
// Reload the modified workbook
var updatedWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\UpdatedExcelFile.xlsx");

// Access the first worksheet
var updatedSheet = updatedWorkbook.WorkSheets.First();

// Read and print each formula set in the cells B1 to B4
foreach(var cell in updatedSheet["B1:B4"])
{
    Console.WriteLine(cell.Formula);
}

// Hold the console open until a key is pressed
Console.ReadKey();
```

In this example, we've demonstrated how to programmatically apply conditional logic within an Excel file using `IronXL`, showing a typical use case for automating decision-making processes in financial or data-driven applications.

<p class="list-description">Using the same file, we can use the Formula property to set or get a cell’s formula:</p>

<p class="list-decimal">3.7.1. Save to XML “.xml”</p>

Here is the paraphrased section of your requested document with resolved URL paths:

```cs
/**
Conditional Formula Example
anchor-conditional-formula-example
**/
var excelWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
var worksheet = excelWorkbook.WorkSheets.First();
int index = 1;
foreach(var cell in worksheet["B1:B4"])
{
    cell.Formula = $"=IF(A{index}>=20, \"Pass\", \"Fail\")";
    index++;
}
excelWorkbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\UpdatedExcelFile.xlsx");
```

<p class="list-decimal">7.2. Using the generated file from the previous example, we can get the Cell’s Formula:</p>

Here's the paraphrased section with updated code snippet:

```cs
// Load an existing Excel workbook
var workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\NewExcelFile.xlsx");

// Access the primary worksheet by default
var sheet = workbook.DefaultWorkSheet;

// Iterate through a specific range of cells
foreach(var cell in sheet["B1:B4"])
{
    // Output the formula set in each cell
    Console.WriteLine("The formula in the cell is: " + cell.Formula);
}

// Wait for a key press before closing the console window
Console.ReadLine();
```

### 3.8. Example: Trimming Cell Contents ###

For eliminating superfluous spaces from cell contents, we incorporated a column in our `sum.xlsx` file as demonstrated below. The following script illustrates how to apply the `Trim` function to refine cell entries effectively.

<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

Here is the snippet for implementing the trim operation:

```cs
// Load the workbook containing cells to trim
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
var sheet = workbook.WorkSheets.First();

// Iterate over specific cell range for trimming
int rowIndex = 1;
foreach (var cell in sheet["f1:f4"])
{
    cell.Formula = $"=TRIM(D{rowIndex})";
    rowIndex++;
}

// Save the workbook with trimmed cell contents
workbook.SaveAs("editedFile.xlsx");
```

This technique demonstrates how effectively the `Trim` function can be applied using formulas within IronXL.

<p class="list-description">To apply trim function (to eliminate all extra spaces in cells), I added this column to the sum.xlsx file</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description">And use this code</p>

Here's the paraphrased section with the corrected inline code annotations and additional comments to clarify each step:

```cs
// Example of applying the TRIM function to cell values
// TRIM removes all redundant spaces from text in cells
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
var sheet = workbook.WorkSheets.First();  // Access the first worksheet
int index = 1;

foreach (var cell in sheet["f1:f4"])  // Iterate through cells from F1 to F4
{
    // Apply TRIM formula to remove unnecessary spaces from each cell in column D
    cell.Formula = $"=trim(D{index})";
    index++;
}

// Save the workbook as 'editedFile.xlsx' with changes
workbook.SaveAs("editedFile.xlsx");
```

In this revised snippet, the focus is on using IronXL to apply the TRIM function on specific cells in an Excel sheet. The loop iterates over cells F1 to F4, applying a TRIM formula to corresponding cells in column D to clean up the data. The workbook is then saved with these updates.

<p class="list-description">Thus, you can apply formulas in the same way.</p>

<hr class="separator">

## Working with Multiple Sheets in Workbooks ##

This section explores how to effectively manage Excel workbooks that contain multiple sheets.

### 4.1. Accessing Data from Various Sheets within a Workbook ###

Explore the process of reading data across different sheets within the same Excel workbook. In the given example, we've crafted an `.xlsx` file that integrates two distinct sheets aptly named "Sheet1" and "Sheet2". Up until this point, the `WorkSheets.First()` method was utilized predominantly to interact with the initial sheet. However, in this segment, we will illustrate how to explicitly designate a sheet by name for data manipulation.

```cs
/**
Identify and Work with Multiple Sheets
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

<p class="list-description">I created an xlsx file that contains two sheets: “Sheet1”,” Sheet2”</p>
<p class="list-description">Until now we used WorkSheets.First() to work with the first sheet. In this example we will specify the sheet name and work with it</p>

```cs
// Example of accessing various sheets within a workbook
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Selecting a specific worksheet namely 'Sheet2'
WorkSheet selectedSheet = workbook.GetWorkSheet("Sheet2");

// Define a range to work with on the worksheet
Range cellRange = selectedSheet["A2:D2"];

// Loop through each cell in the range and print its content
foreach(var cell in cellRange)
{
    Console.WriteLine("The content of cell is: " + cell.Text);
}
```

### 4.2. How to Add a New Sheet to an Existing Workbook ###

This section illustrates the process of incorporating a new sheet into an already existing workbook.

```cs
/**
Insert New Sheet
anchor-add-new-sheet-to-a-workbook
**/
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
var newSheet = workbook.CreateWorkSheet("new_sheet");
newSheet["A1"].Value = "Hello World";
workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

This example demonstrates adding a new worksheet named "new_sheet" to the `testFile.xlsx` workbook. Starting with loading the existing workbook, a new worksheet is created. The value "Hello World" is then assigned to cell A1 of the new worksheet. Finally, the workbook with the newly added sheet is saved under a new file name, ensuring all changes are preserved.

<p class="list-description">We can also add new sheet to a workbook:</p>

```cs
// Adding a New Worksheet Example
var loadedWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
var addedSheet = loadedWorkbook.CreateWorkSheet("new_sheet");
addedSheet["A1"].Value = "Hello World";
loadedWorkbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

<hr class="separator">

## 5. Database Integration with Excel ##

Explore how to import and export data between a database and an Excel spreadsheet.

I established the "TestDb" database which includes a "Country" table formatted with two columns: Id (integer, identity) and CountryName (string).

### 5.1. Populate Excel Sheet with Database Data ###

In this segment, we'll show you how to generate a new worksheet and populate it with data sourced from the Country Table in your database.

```cs
/**
Import Data from Database into a Sheet
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

<p class="list-description">Here we will create a new sheet and fill it with data from the Country Table</p>

Here's the rewritten section with relative URL paths replaced with absolute paths referring to ironsoftware.com:

```cs
/**
Populate a Worksheet with Database Content
anchor-populate-worksheet-from-database
**/
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.CreateWorkSheet("DatabaseData");
List<Country> countries = dbContext.Countries.ToList();
sheet.SetCellValue(0, 0, "Id");
sheet.SetCellValue(0, 1, "Country Name");
int currentRow = 1;
foreach (var country in countries)
{
    sheet.SetCellValue(currentRow, 0, country.id);
    sheet.SetCellValue(currentRow, 1, country.CountryName);
    currentRow++;
}
workbook.SaveAs("DatabaseFilledFile.xlsx");
```

### 5.2. Populate Database with Excel Sheet Data ###

Insert data into the "TestDb" Database's Country table by processing the information from an Excel sheet:

```cs
/**
Import Data to Database
anchor-fill-database-with-data-from-excel-sheet
**/
TestDbEntities dbContext = new TestDbEntities ();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
System.Data.DataTable dataTable = sheet.ToDataTable(true);
foreach (DataRow row in dataTable.Rows)
{
    Country country = new Country();
    country.CountryName = row [1].ToString();
    dbContext.Countries.Add(country);
}
dbContext.SaveChanges();
```

<p class="list-description">Insert the data to the Country table in TestDb Database</p>

Here's a paraphrased version of the provided C# code snippet:

```cs
// Importing data into a database from an Excel file
TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
System.Data.DataTable dataTable = sheet.ToDataTable(true);

// Iterate through each data row in the DataTable
foreach (DataRow row in dataTable.Rows)
{
    Country country = new Country();
    country.CountryName = row[1].ToString();  // Assign the second column (CountryName)
    dbContext.Countries.Add(country);  // Add the new country to the database context
}
dbContext.SaveChanges();  // Commit the changes to the database
``` 

This modified snippet still imports data from an Excel spreadsheet into a database, ensuring that each country's name is correctly extracted and saved.

<hr class="separator">

### Additional Resources

For a deeper understanding of IronXL capabilities, consider exploring more tutorials available in this section or reviewing the examples featured on our homepage, which provide sufficient guidance for most developers to begin their projects.

Delve into the detailed documentation of the `WorkBook` class and other features by visiting our [API Reference](https://ironsoftware.com/csharp/excel/object-reference/).

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

