# C# Excel File Manipulation [Interop-Free] - Code Example Tutorial

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file/>***


Discover through practical examples how to create, open, and preserve Excel documents using C#, while performing elementary operations such as summing, averaging, counting, among others. IronXL.Excel is an independent .NET library capable of interacting with various spreadsheet formats. This library functions seamlessly without the need for [Microsoft Excel](https://products.office.com/en-us/excel) installation or reliance on Interop.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

<h2>Use IronXL to Open and Write Excel Files</h2>

------
Easily open, write, save, and modify Excel files using the highly intuitive [IronXL C# library](https://ironsoftware.com/csharp/excel/).

Acquire a [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or start with your own to follow along with this guide.

Steps to begin:

1. Add the IronXL Excel Library to your project from [NuGet](https://www.nuget.org/packages/IronXL.Excel) or by downloading the DLL.
   
2. Employ the `WorkBook.Load` method for accessing any XLS, XLSX, or CSV files.

3. Retrieve cell values effortlessly with the syntax: `sheet["A11"].DecimalValue`

Throughout this tutorial, we will guide you through:

- **Installing IronXL.Excel**: Detailed steps on integrating the IronXL.Excel library into your project.
- **Basic Operations**: Instructions on how to create or open a workbook, select sheets and cells, and save your work.
- **Advanced Sheet Operations**: Explore advanced features such as adding headers and footers, performing mathematical functions, and more.

<h4>Open an Excel File : Quick Code</h4>

Here's the paraphrased section of the article:

```cs
using IronXL;

// Load the workbook from a file
WorkBook workbook = WorkBook.Load("test.xlsx");
// Access the default worksheet
WorkSheet worksheet = workbook.DefaultWorkSheet;
// Define a range of cells to work with
IronXL.Range cellRange = worksheet["A2:A8"];
decimal accumulatedTotal = 0;

// Loop through the specified range of cells
foreach (var cell in cellRange)
{
    // Display the row index and value of the cell
    Console.WriteLine("Cell at row {0} has the value: '{1}'", cell.RowIndex, cell.Value);
    // Check if the cell contains a numeric value
    if (cell.IsNumeric)
    {
        // Sum up the decimal values to handle precision
        accumulatedTotal += cell.DecimalValue;
    }
}

// Validate by comparing calculated total with a pre-defined cell value
if (worksheet["A11"].DecimalValue == accumulatedTotal)
{
    // Output test result
    Console.WriteLine("Basic Test Passed");
}
```

<h4>Write and Save Changes to the Excel File : Quick Code</h4>

Below is the paraphrased section of the article, with resolved URL paths:

-----
```cs
// Assign a decimal value to cell B1
workSheet["B1"].Value = 11.54;

// Save the workbook with the updated changes
workBook.SaveAs("test.xlsx");
```

<hr class="separator">
<p class="main-content__segment-title">Step 1</p>

## 1. Installing the IronXL C# Library at No Cost

-------------------------------------------

IronXL.Excel offers a robust and adaptable library designed for managing Excel files within a .NET environment. This library supports a broad range of .NET project types including desktop applications, ASP.NET MVC, and .NET Core applications, facilitating tasks such as reading, writing, modifying, and saving Excel documents. Whether for desktop or web applications, IronXL ensures comprehensive Excel manipulation without the need for Microsoft Excel.

<h3>Install the Excel Library to your Visual Studio Project with NuGet</h3>

The initial step involves incorporating IronXL.Excel into your project. You have two methods to achieve this: using the NuGet Package Manager or the NuGet Package Manager Console.

For installing IronXL.Excel with the NuGet Package Manager, leverage the graphical interface as follows:

1. Navigate by clicking with your mouse. Right-click on your project's name in your solution explorer and choose "Manage NuGet Packages."

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <p><img src="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

2. Navigate to the browse tab, type `IronXL.Excel` in the search bar, and click `Install`.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" target="_blank">

![Search for IronXL](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg)

</a>

Below is your paraphrased content with URLs resolved against ironsoftware.com:

-----
3. Installation Complete
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>



<h3>Install Using NuGet Package Manager Console</h3>

1. Navigate from `Tools` to `NuGet Package Manager`, then select `Package Manager Console`.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

```
2. Execute the following command in the console: `Install-Package IronXL.Excel -Version 2019.5.2`
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<h3>Manually Install with the DLL</h3>

Additionally, you have the option to manually integrate the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) into your project or the global assembly cache if you prefer a non-NuGet installation method.

```
 PM > Install-Package IronXL.Excel
```

# C# Write to Excel [Using IronXL Library] Code Tutorial

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file/>***


This tutorial provides a series of step-by-step instructions for creating, opening, and editing Excel files using C#, exploring basic operations like summation, averaging, counting, and beyond. The IronXL.Excel library, which operates independently of .NET, enables these functions without necessitating the installation of Microsoft Excel or reliance on Interop.

<hr class="separator">

<p class="main-content__segment-title">Introduction</p>

<h2>Manipulate Excel Files with IronXL</h2>

Learn how to manage Excel documents effectively using the straightforward features of the <a href="https://ironsoftware.com/csharp/excel/" target="_blank">IronXL C# Library.</a>

Download a <a href="https://github.com/magedo93/IronSoftware.git" target="_blank">GitHub sample project</a> or initiate one on your own to proceed with the tutorial.

1. Acquire the IronXL Excel Library via <a href="https://www.nuget.org/packages/IronXL.Excel" target="_blank">NuGet</a> or through direct DLL download.
2. Load any XLS, XLSX, or CSV document utilizing the `WorkBook.Load` method.
3. Extract Cell values using an intuitive command: `sheet["A11"].DecimalValue`.

Throughout this tutorial, you'll learn:

- How to integrate IronXL.Excel into your project.
- The foundational steps for working with Excel files, including creation, opening, cell selection, and saving.
- Enhanced techniques for augmenting Excel sheets such as adding headers, performing calculations, and injecting additional functionalities.

<h4>Opening an Excel File: Quick Example</h4>

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;
IronXL.Range range = workSheet["A2:A8"];
decimal total = 0;

// Loop through a range of cells
foreach (var cell in range)
{
    Console.WriteLine($"Cell {cell.RowIndex} has value '{cell.Value}'");
    if (cell.IsNumeric)
    {
        // Secure accurate decimal value handling
        total += cell.DecimalValue;
    }
}

// Validate formula outcome
if (workSheet["A11"].DecimalValue == total)
{
    Console.WriteLine("Validation Successful");
}
```

<h4>Editing and Saving Changes in an Excel File: Quick Example</h4>

```cs
workSheet["B1"].Value = 11.54;

// Commit Changes
workBook.SaveAs("updated.xlsx");
```

<hr class="separator">

Throughout these instructions, the aim is to provide practical, easy-to-follow guidance to harness the robust features of the IronXL library efficiently. Whether you're constructing a complex financial model or managing simple data lists, this tutorial is designed to get you up and running with minimal fuss.

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Fundamental Tasks: Creating, Opening, and Saving ##

### 2.1. Example Project: HelloWorld Console Application ###

<p class="list-description">Initiate a HelloWorld Project</p>

<p class="list-decimal">2.1.1. Launch Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Opt for 'Create New Project'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Select 'Console App (.NET framework)'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Name your sample 'HelloWorld' and then hit 'Create'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. A new console application is now established</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Incorporate IronXL.Excel by choosing 'Install'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Create initial code to read the first Excel cell on the first sheet and display it:</p>

```cs
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
WorkSheet worksheet = workbook.WorkSheets.First();
string cellValue = worksheet["A1"].StringValue;
Console.WriteLine(cellValue);
```

### 2.2. Generate a New Excel Document ###

<p class="list-description">Using IronXL to craft a fresh Excel document</p>

```cs
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
workbook.Metadata.Title = "New IronXL File";
WorkSheet worksheet = workbook.CreateWorkSheet("InitialSheet");
worksheet["A1"].Value = "Hello World";
worksheet["A2"].Style.BottomBorder.SetColor("#ff6600");
worksheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
```

### 2.3. Opening Diverse File Formats as Workbooks ###

<p class="list-decimal">2.3.1. Load a CSV File</p>

<p class="list-decimal">2.3.2. Construct a text file, populate it with a list of names and ages, and save it as 'CSVList.csv'</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">Your CSV file should have the following structure</p>

```cs
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
WorkSheet worksheet = workbook.WorkSheets.First();
string csvCell = worksheet["A1"].StringValue;
Console.WriteLine(csvCell);
```

<p class="list-decimal">
    2.3.3. Load an XML File
    <span class="list-description">Generate an XML file containing a list of countries with specific properties.</span>
```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">2.3.4. Adapt the XML content into a workbook format:</p>

```cs
DataSet xmlDataSet = new DataSet();
xmlDataSet.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workbook = IronXL.WorkBook.Load(xmlDataSet);
WorkSheet worksheet = workbook.WorkSheets.First();
```

<p class="list-decimal">
    2.3.5. Load a JSON List
    <span class="list-description">Create a JSON file listing countries:</span>
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

<p class="list-decimal">2.3.6. Define a class to map to the JSON structure:</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/create-country-model.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/create-country-model.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">Class code snippet:</p>

```cs
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

<p class="list-decimal">2.3.8. Incorporate Newtonsoft library to convert JSON into a list of country models:</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-newtonsoft-library-to-convert-json.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-newtonsoft-library-to-convert-json.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.3.9 Develop a new class extension named 'ListConvertExtension' to convert the list to a dataset:</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/convert-list-to-dataset.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/convert-list-to-dataset.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">Add the following code:</p>

```cs
public static class ListConvertExtension
{
    public static DataSet ToDataSet<T>(this IList<T> list)
    {
        Type elementType = typeof(T);
        DataSet ds = new DataSet();
        DataTable t = new DataTable();
        ds.Tables.Add(t);
        // For each public property on T add a column to the table
        foreach (var propInfo in elementType.GetProperties())
        {
            Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
            t.Columns.Add(propInfo.Name, ColType);
        }
        // Iterate through each property on T and add each value to the table
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

<p class="list-decimal">Finally, load this dataset into a workbook:</p>

```cs
StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
var xmlDataset = countryList.ToDataSet();
WorkBook workbook = IronXL.WorkBook.Load(xmlDataset);
WorkSheet worksheet = workbook.WorkSheets.First();
```

### 2.1. Example Project: "HelloWorld" Console Application ###

<p class="list-description">Creating the HelloWorld Project</p>

<p class="list-decimal">2.1.1. Begin by Opening Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Select the 'Create New Project' Option</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Opt for Console App (.NET Framework)</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Assign the name “HelloWorld” to our sample and proceed to create it</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. You now have a console application ready</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Include IronXL.Excel -> proceed with installation</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Now, let's write some initial code to read the first cell of the first sheet in the Excel file and display it</p>

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

Here is the paraphrased version of the provided code snippet from the IronXL tutorial:

```cs
// Load the workbook from the current directory
var workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\HelloWorld.xlsx");

// Access the first worksheet in the workbook
var worksheet = workbook.WorkSheets.First();

// Retrieve the value from cell A1 as a string
string cellValue = worksheet["A1"].StringValue;

// Output the value to the console
Console.WriteLine(cellValue);
```

### 2.2. Constructing a New Excel Document

Explore the initial creation of an Excel document using IronXL in this segment:

```cs
WorkBook newWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
newWorkbook.Metadata.Title = "New IronXL Document";
WorkSheet newWorkSheet = newWorkbook.CreateWorkSheet("FirstSheet");
newWorkSheet["A1"].Value = "Hello World";

// Set a stylized bottom border on the next cell and choose a color
newWorkSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
newWorkSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
```

In this code block, we create a new `WorkBook` instance and set its format to `.XLSX`. A new worksheet named "FirstSheet" is created, and a greeting "Hello World" is inserted into cell A1. We then apply a dashed border with an orange color below the second cell.

<p class="list-description">Create a new Excel file using IronXL</p>

Here's the paraphrased section with the relative URL paths resolved against ironsoftware.com:

```cs
WorkBook newWorkBook = WorkBook.Create(ExcelFileFormat.XLSX);
newWorkBook.Metadata.Title = "IronXL New Document";
WorkSheet firstSheet = newWorkBook.CreateWorkSheet("Sheet1");
firstSheet["A1"].Value = "Hello World";
firstSheet["A2"].Style.BottomBorder.SetColor("#ff6600"); // Setting the bottom border color to orange
firstSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed; // Using dashed type for the border
```

### 2.3. Loading Different File Formats as Workbooks ###

This segment details the steps to import different data file types such as CSV, XML, and JSON into IronXL workbooks.

#### 2.3.1. Loading a CSV File ####

Initially, we start by loading a CSV file into IronXL. This straightforward process involves using the `WorkBook.Load` method to pull in data from a pre-existing CSV file stored in the designated directory.

#### 2.3.2. Composing and Saving a CSV List ####

To proceed, create a plain text file containing a collection of names and ages. Here’s a basic structure to follow, after which you can save this as `CSVList.csv`.

![CSV List Structure](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg)

#### 2.3.3. Opening an XML File ####

In this part, you will craft an XML file that encompasses a list of countries. Each country node should include attributes detailing specifics about the country such as its 'code', 'continent', etc.

An XML structure would look something like this:

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

Upon creating the XML file, utilize the following C# snippet to load the XML file into a workbook:

```cs
DataSet xmlDataSet = new DataSet();
xmlDataSet.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workbook = IronXL.WorkBook.Load(xmlDataSet);
WorkSheet worksheet = workbook.WorkSheets.First();
```

#### 2.3.5. Loading a JSON List as Workbook ####

For handling JSON data, begin by crafting a JSON representation of a country list:

```json
[
    {"name": "United Arab Emirates", "code": "AE"},
    {"name": "United Kingdom", "code": "GB"},
    {"name": "United States", "code": "US"},
    {"name": "United States Minor Outlying Islands", "code": "UM"}
]
```

#### 2.3.6 and Following Steps: Class Definition and Data Conversion ####

In subsequent steps, define a `CountryModel` class to map JSON data to a list, and use Newtonsoft library to parse this JSON into a `CountryModel` list. Add a custom extension class `ListConvertExtension` to facilitate converting a list to a dataset.

Here’s the code snippet to load the dataset as an IronXL workbook:

```cs
StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
var xmlDataSet = countryList.ToDataSet();
WorkBook workbook = IronXL.WorkBook.Load(xmlDataSet);
WorkSheet worksheet = workbook.WorkSheets.First();
```

<p class="list-decimal">2.3.1. Open CSV file</p>

<p class="list-decimal">2.3.2 Create a new text file and add to it a list of names and ages (see example) then save it as CSVList.csv</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Your code snippet should look like this</p>

Here's a paraphrased version of the provided C# code snippet:

```cs
// Load the workbook from a CSV file within the current working directory
WorkBook workbook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\CSVList.csv");

// Access the first worksheet in the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Retrieve the value of the first cell in the first column
string firstCellValue = worksheet["A1"].StringValue;

// Output the value of the cell to the console
Console.WriteLine(firstCellValue);
```

<p class="list-decimal">

### 2.3.3 Open an XML File

To work with XML data within Excel, you'll first need to create an XML file. This file should contain structured data, such as a list of countries. Here’s how you can structure your XML:

```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

After creating your XML file, you can then load it as an Excel workbook by utilizing the following steps:

```cs
// Load the XML data into a DataSet
DataSet xmldataset = new DataSet();
xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");

// Load the DataSet into an IronXL WorkBook
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();
``` 

This method transforms your structured XML data directly into an Excel format, making it easy to handle and manipulate within .NET applications using IronXL.

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
DataSet xmlDataSet = new DataSet();
xmlDataSet.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workbook = IronXL.WorkBook.Load(xmlDataSet);
WorkSheet worksheet = workbook.WorkSheets.First();
```

<p class="list-decimal">

### 2.3.5. Opening JSON Lists as Workbooks ###

For this demonstration, we'll be loading data from a JSON file into an Excel workbook.

```cs
[{
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
}]
```

#### Steps to Load JSON Data into a Workbook:

1. **Construct a Model for JSON Data**

First, define a class that reflects the structure of your JSON data:

```cs
public class CountryModel
{
    public string name { get; set; }
    public string code { get; set; }
}
```

2. **Parse JSON with Newtonsoft Library**

Include the Newtonsoft library to your project to facilitate JSON parsing. You can add this to your project from NuGet.

![Add Newtonsoft Library](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-newtonsoft-library-to-convert-json.png)

3. **Convert JSON List to Dataset**

Before loading the JSON data into a workbook, convert it into a dataset format using a custom extension method:

```cs
public static class ListConvertExtension
{
    public static DataSet ToDataSet<T>(this IList<T> list)
    {
        Type elementType = typeof(T);
        DataSet ds = new DataSet();
        DataTable t = new DataTable();
        ds.Tables.Add(t);
        foreach (var propInfo in elementType.GetProperties())
        {
            Type columnType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
            t.Columns.Add(propInfo.Name, columnType);
        }
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

4. **Load JSON as Excel Workbook**

Finally, read the JSON file, deserialize it into the list of `CountryModel`, convert it to a dataset, and load it as an Excel workbook:

```cs
StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<CountryModel>>(jsonFile.ReadToEnd());
var xmldataset = countryList.ToDataSet();
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();
```

<span class="list-description">Create JSON country list</span>
</p>

```cs
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
        "name": "USA",
        "code": "US"
    },
    {
        "name": "US Minor Outlying Islands",
        "code": "UM"
    }
]
```

<p class="list-decimal"></p>
<p class="list-decimal">2.3.6. Create a country model that will map to JSON</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Here is the class code snippet</p>

```cs
public class NationalEntity
{
    public string name { get; set; }
    public string code { get; set; }
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
public static class ConversionHelper
{
    /// <summary>
    /// Converts a list of objects to a DataTable within a DataSet.
    /// </summary>
    /// <typeparam name="T">The type of objects in the list.</typeparam>
    /// <param name="list">The list of objects to convert.</param>
    /// <returns>A DataSet containing the DataTable constructed from the list.</returns>
    public static DataSet ConvertToDataSet<T>(this IList<T> list)
    {
        Type itemType = typeof(T);
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable();
        dataSet.Tables.Add(table);

        // Create a DataColumn for each public property of T
        foreach (var propertyInfo in itemType.GetProperties())
        {
            Type columnType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
            table.Columns.Add(propertyInfo.Name, columnType);
        }

        // Populate the table with values from the list
        foreach (T item in list)
        {
            DataRow row = table.NewRow();
            foreach (var propertyInfo in itemType.GetProperties())
            {
                row[propertyInfo.Name] = propertyInfo.GetValue(item, null) ?? DBNull.Value;
            }
            table.Rows.Add(row);
        }

        return dataSet;
    }
}
```

<p class="list-decimal"></p>
<p class="list-decimal">And finally load this dataset as a workbook</p>

Here is the paraphrased section with resolved relative URL paths:

```cs
// Open the JSON file containing the country list
StreamReader jsonFileStream = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");

// Deserialize the JSON data into an array of CountryModel objects
var listOfCountries = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFileStream.ReadToEnd());

// Convert the list of CountryModel objects to a DataSet
DataSet countriesDataSet = listOfCountries.ToDataSet();

// Load the dataset into an IronXL WorkBook
WorkBook excelWorkbook = IronXL.WorkBook.Load(countriesDataSet);

// Select the first worksheet from the workbook
WorkSheet excelSheet = excelWorkbook.WorkSheets.First();
```

This section of the code demonstrates how to read a JSON file, convert its content into a dataset, and then load that dataset into an IronXL `WorkBook`. It finishes by obtaining the first worksheet from the workbook.

### 2.4. Saving and Exporting Files ###

This section covers methods for storing and exporting Excel documents into various formats using IronXL, providing flexibility depending on the file requirements.

#### 2.4.1. Saving as `.xlsx`####
To save an Excel file in `.xlsx` format, use the `SaveAs` method provided by IronXL:

```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "New IronXL File";

WorkSheet workSheet = workBook.CreateWorkSheet("FirstSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewHelloWorld.xlsx");
```

#### 2.4.2. Saving as CSV (`.csv`)####
To export data to a `.csv` file, specify a filename and a delimiter. You can adjust the delimiter to suit your data needs:

```cs
workBook.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\NewHelloWorld.csv", delimiter: "|");
```

#### 2.4.3. Saving as JSON (`.json`)####
To convert an Excel workbook into a JSON format, use `SaveAsJson`. This is beneficial for web applications and data exchange:

```cs
workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

The resulting JSON file will resemble:

```cs
[
    ["Hello World"],
    [""]
]
```

#### 2.4.4. Saving as XML (`.xml`)####
Finally, to save the workbook as an XML file which might be used for data storage or inter-software communication, use the `SaveAsXml` method:

```cs
workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.xml");
```

This example generates the following XML structure:

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

Each of these methods enhances the versatility of document management, allowing users to tailor the output to specific requirements of their systems or third-party applications.

<span class="list-description">We can save or export the Excel file to multiple file formats like (“.xlsx”,”.csv”,”.html”) using one of the following commands.</span>

<p class="list-decimal">

### 2.4.1. Saving as ".xlsx"

To save an Excel workbook in the `.xlsx` format using IronXL, simply utilize the `SaveAs` method. Below is a detailed guide on how to execute this operation effectively:

```cs
// Create a new workbook with a specified format
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "New IronXL File";

// Add a new worksheet and populate data
WorkSheet workSheet = workBook.CreateWorkSheet("FirstSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

// Save the workbook as .xlsx in the specified path
workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
```
This code snippet generates a new workbook and saves it as `HelloWorld.xlsx`, demonstrating the adaptability and ease of managing Excel files with IronXL.

<span class="list-description">To Save to “.xlsx” use saveAs function</span>
</p>

```cs
// Create a new Excel workbook
WorkBook myWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
// Set the title in workbook metadata
myWorkbook.Metadata.Title = "New IronXL File";

// Add and name a new worksheet
WorkSheet mySheet = myWorkbook.CreateWorkSheet("FirstSheet");
// Assign a value to cell A1
mySheet["A1"].Value = "Hello World";
// Style cell A2 with a bottom border
mySheet["A2"].Style.BottomBorder.SetColor("#ff6600");
mySheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

// Save the workbook to a file in the current directory
myWorkbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
```

<p class="list-decimal">

### 2.4.2 Saving as CSV

To export your workbook in a “.csv” format, utilize the `SaveAsCsv` method. This function requires two arguments: the desired filename along with its path, and a delimiter which can be a comma (`,`), pipe (`|`), or colon (`:`).
```

</p>

Here is the paraphrased section of the article, with relative URL paths resolved to `ironsoftware.com`:

```cs
// Save the workbook as a CSV file with a specified delimiter
workBook.SaveAsCsv($"{Directory.GetCurrentDirectory()}\\Files\\HelloWorld.csv", delimiter: ",");
```

<p class="list-decimal">

### 2.4.3 Export to JSON Format ".json" ###

This section details the process of exporting an Excel workbook to a JSON file format using IronXL. This capability allows for versatile data sharing and storage options that align with modern data handling standards.

```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "IronXL New File";

WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

Following the execution of this code, the resulting JSON file will appear as:

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

This function empowers developers to seamlessly transition workbook data into a widely used and easily consumable JSON format, enhancing integration with various applications and services.

<span class="list-description">To save to Json “.json” use SaveAsJson as follow</span>
</p>

Here's the paraphrased section with the resolved relative URL path:

```cs
// Export the Excel workbook to a JSON file in the current directory
workBook.ExportToJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
```

<p class="list-decimal">
  <span class="list-description">The result file should look like this</span>
</p>

```cs
[
    [
        "Hello World"
    ],
    [
        "Empty" // Previously there was an empty string here
    ]
]
```

<p class="list-decimal">

-----
### 2.4.4 Save to XML

To convert the Excel file into an XML format, you can use the `SaveAsXml` method as demonstrated below:

```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Title = "IronXL New File";

WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.xml");
```

Here's the XML output structure that results from the command:

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
This XML structure clearly represents the data in a well-defined, hierarchical way suitable for further processing or storage.

-----

<span class="list-description">To save to xml use SaveAsXml as follow</span>
</p>

Here is the paraphrased section of the article with the resolved URL path for the Iron Software domain:

```cs
// This command saves the workbook to an XML file in the current directory.
workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
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
    <Column1 xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></Column1>
  </FirstWorkSheet>
</FirstWorkSheet>
```

<hr class="separator">

## 3. Advanced Operations: Sum, Average, Count, and More ##

This section delves into the utilization of popular Excel functions such as SUM, AVG, and COUNT, with detailed coding examples provided for each.

### 3.1 Sum Operation Example ###

For demonstration, an Excel file named "Sum.xlsx" was prepared with a series of numbers. Below is the code snippet to calculate the sum of these numbers:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal sumTotal = workSheet["A2:A4"].Sum();
Console.WriteLine(sumTotal);
```

### 3.2 Average Calculation Example ###

Using the same Excel file "Sum.xlsx", the next snippet demonstrates how to compute the average of the values:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal averageValue = workSheet["A2:A4"].Avg();
Console.WriteLine(averageValue);
```

### 3.3 Count Entries Example ###

Also using "Sum.xlsx", this code snippet illustrates how to count the number of entries in a specified range:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
int totalEntries = workSheet["A2:A4"].Count();
Console.WriteLine(totalEntries);
```

### 3.4 Maximum Value Example ###

To find the maximum value in a range within "Sum.xlsx":

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal maxValue = workSheet["A2:A4"].Max();
Console.WriteLine(maxValue);
```

### 3.5 Minimum Value Example ###

Identifying the minimum value in a range using the same file:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal minValue = workSheet["A1:A4"].Min();
Console.WriteLine(minValue);
```

### 3.6 Ordering Cells ###

Ordering cells from the same file can be done either in ascending or descending:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
workSheet["A1:A4"].SortAscending();
// Use SortDescending() to order them in descending order
workBook.SaveAs("SortedSheet.xlsx");
```

### 3.7 Conditional Formulas ###

Applying conditional formulas to cells:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
int i = 1;
foreach (var cell in workSheet["B1:B4"])
{
    cell.Formula = "=IF(A" + i + ">=20,\"Pass\",\"Fail\")";
    i++;
}
workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
```

### 3.8 Trim Function Example ###

Finally, to apply the trim function and remove all extra spaces from cells in the modified file:

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
int index = 1;
foreach (var cell in workSheet["F1:F4"])
{
    cell.Formula = $"=TRIM(D{index})";
    index++;
}
workBook.SaveAs("editedFile.xlsx");
```

This section aims to give a practical view on applying standard Excel functionalities using code snippets in the IronXL framework.

### Example 3.1: Calculating Sum ###

<p class="list-description">This example demonstrates how to calculate the sum from a list of numbers that have been manually entered into an Excel file named "Sum.xlsx."</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

```cs
// Load the workbook from a specific path
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet
WorkSheet workSheet = workBook.WorkSheets.First();

// Calculate the sum of values from cells A2 to A4
decimal sum = workSheet["A2:A4"].Sum();

// Print the sum to the console
Console.WriteLine(sum);
```

<p class="list-description">Let’s find the sum for this list. I created an Excel file and named it “Sum.xlsx” and added this list of numbers manually</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

```cs
// Load the Workbook with the specified Excel file
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Get the first worksheet from the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Calculate the sum of values in the range A2 to A4
decimal totalSum = worksheet["A2:A4"].Sum();

// Display the total sum in the console
Console.WriteLine(totalSum);
```

### 3.2. Example: Calculating the Average ###

In this segment, we'll demonstrate how to compute the average from a set of values using IronXL. We utilize an Excel file named "Sum.xlsx" that contains a sequence of numbers.

```cs
// Load the workbook
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
// Access the first worksheet
WorkSheet workSheet = workBook.WorkSheets.First();
// Calculate the average of values in range A2 to A4
decimal avg = workSheet["A2:A4"].Avg();
// Display the average on the console
Console.WriteLine(avg);
```

<p class="list-description">Using the same file, we can get the average:</p>

```cs
// Load the workbook from a specific path
WorkBook workbook = IronXL.WorkBook.Load($@"{System.IO.Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
// Access the first worksheet
WorkSheet worksheet = workbook.WorkSheets.First();
// Calculate the average value of cells in the range A2 to A4
decimal average = worksheet["A2:A4"].Avg();
// Output the average to the console
Console.WriteLine(average);
```

### 3.3. Example: Counting Elements ###

Let's demonstrate how to determine the number of items in a sequence in an Excel file named “Sum.xlsx” that we prepared earlier.

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal elementCount = workSheet["A2:A4"].Count();
Console.WriteLine(elementCount);
```

<p class="list-description">Using the same file, we can also get the number of elements in a sequence:</p>

```cs
// Load the workbook from the specified location
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Select the first worksheet from the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Calculate the count of cells within the specified range
decimal cellCount = worksheet["A2:A4"].Count();

// Output the count to the console
Console.WriteLine(cellCount);
```

### 3.4. Example: Finding the Maximum Value ###

Discover how to find the highest value from a list within an Excel file. This tutorial uses an example Excel file named "Sum.xlsx" that contains a column of numbers entered manually.

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal maximumValue = workSheet["A2:A4"].Max();
Console.WriteLine(maximumValue);
```

With the `Max()` method, you can effortlessly retrieve the highest value from a range of cells. If you need to introduce more complex calculations or consider only specific cells based on a condition, IronXL also supports transformations right within the `Max` method:

```cs
WorkBook loadedWorkbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet activeWorksheet = loadedWorkbook.WorkSheets.First();
bool result = activeWorksheet["A1:A4"].Max(cell => cell.IsFormula);
Console.WriteLine(result);
```

This enhanced functionality checks if any cells in the specified range contain formulas by returning `true` or `false`, displayed in the console as necessary. The code example demonstrates a simplified way to augment Excel file interactions, showcasing that no formulae are detected in this instance by outputting "false."

<p class="list-description">Using the same file, we can get the max value of range of cells:</p>

Here's the paraphrased section with the URLs resolved:

```cs
// Load the workbook from a specified directory
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Calculate the maximum value from the range A2 to A4
decimal maximumValue = worksheet["A2:A4"].Max();

// Output the maximum value to the console
Console.WriteLine(maximumValue);
```

<p class="list-description">– We can apply the transform function to the result of max function:</p>

Here's the paraphrased snippet, with the relative URL paths resolved:

```cs
// Load the workbook from the current directory
WorkBook book = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
// Access the first worksheet in the workbook
WorkSheet sheet = book.WorkSheets.First();
// Evaluate if the maximum value in the range A1 to A4 is a formula
bool hasFormula = sheet["A1:A4"].Max(cell => cell.IsFormula);
// Output the result to the console
Console.WriteLine(hasFormula);
```

<p class="list-description">This example writes “false” in the console.</p>

### Example: Finding the Minimum Value in Excel ###

This section demonstrates how to identify the smallest value within a specified range of cells in an Excel file named `Sum.xlsx`. This process is achieved using the IronXL library, which facilitates the interaction with Excel files directly from .NET applications, without needing Excel installed on the machine.

#### Step-by-Step Process to Find the Minimum Value ####

1. **Load the Workbook**: First, we load the Excel workbook which contains the data.

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
```

2. **Select the Desired Worksheet**: By default, we access the first sheet in the workbook.

```cs
WorkSheet workSheet = workBook.WorkSheets.First();
```

3. **Retrieve and Compute the Minimum Value**: We then focus on a specific range of cells, from `A1` to `A4`, and apply the `Min()` method to find the minimum value.

```cs
decimal min = workSheet["A1:A4"].Min();
```

4. **Output the Result**: Finally, we output the minimum value to the console.

```cs
Console.WriteLine(min);
```

This simple sequence of commands makes it easy to integrate Excel data operations into .NET applications using IronXL, showcasing its utility in processing and analyzing spreadsheet data efficiently.

<p class="list-description">Using the same file, we can get the min value of range of cells:</p>

Here's the paraphrased section where we're finding the minimum value from a specified range in an Excel workbook using IronXL:

```cs
// Load the workbook from a specified path
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Calculate the minimum value from the range A1 to A4
decimal minimumValue = worksheet["A1:A4"].Min();

// Output the minimum value to the console
Console.WriteLine(minimumValue);
```

### 3.6. Example of Sorting Cells ###

In this demonstration, we'll showcase how to organize cells within a spreadsheet. We utilize an Excel file named "Sum.xlsx" that contains a sequence of data entries for this example.

Here's how you can sort these cells either in ascending or descending order based on your preference:

```cs
// Load the workbook from a file
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
var cells = workSheet["A1:A4"];

// Sort cells in ascending order
cells.SortAscending();
// Uncomment the following line to sort the cells in descending order
// cells.SortDescending();

// Save the sorted data to a new file
workBook.SaveAs("SortedSheet.xlsx");
```

This streamlined approach allows you to easily adjust the order of data in your Excel sheets using IronXL, enhancing data organization and analysis capabilities.

<p class="list-description">Using the same file, we can order cells by ascending or descending:</p>

Here’s the rewritten section of the code with an updated version that provides clearer comments and slightly varied logic to achieve similar outcomes.

```cs
// Load the workbook from the specified path
WorkBook workBook = IronXL.WorkBook.Load($@"{Environment.CurrentDirectory}\Files\Sum.xlsx");

// Access the first worksheet in the workbook
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Sort data in cells from A1 to A4 in ascending order. Uncomment the next line to sort in descending order.
workSheet["A1:A4"].SortAscending();
// workSheet["A1:A4"].SortDescending(); // Uncomment to sort in descending order

// Save the workbook with the sorted data under a new file name
workBook.SaveAs("SortedSheet.xlsx");
```

### 3.7. Example of Using Conditional Formulas ###

This section demonstrates how to incorporate conditional logic directly within your Excel file using IronXL. We use a sample file named "Sum.xlsx" to implement this operation.

#### Applying IF Statements in Cells ####

```cs
// Load the workbook from a specified file
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
int row = 1;

// Loop through cells in the specified range and set conditional formulas
foreach (var cell in workSheet["B1:B4"])
{
    // Set formula to check if value in column A is greater than or equal to 20
    cell.Formula = $"=IF(A{row}>=20, \" Pass\", \" Fail\")";
    row++;
}
// Save the workbook with the newly applied formulas
workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
```

#### Retrieving and Displaying Cell Formulas ####

```cs
// Reload the workbook to reflect the changes made above
WorkBook reloadWorkBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
WorkSheet reloadedWorkSheet = reloadWorkBook.WorkSheets.First();

// Output each formula in the specified cell range
foreach (var cell in reloadedWorkSheet["B1:B4"])
{
    Console.WriteLine(cell.Formula);
}
Console.ReadKey();
```

This tutorial illustrates the proficiency of IronXL in handling conditional logic within Excel cells, facilitating dynamic content generation based on specified criteria.

<p class="list-description">Using the same file, we can use the Formula property to set or get a cell’s formula:</p>

<p class="list-decimal">3.7.1. Save to XML “.xml”</p>

Below is a paraphrased version of the provided C# code section. I've adjusted the syntax and comments for clarity and a slight variation in the logic:

```cs
// Load the workbook from a specified path
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

// Access the first worksheet by default
WorkSheet worksheet = workbook.WorkSheets.First();

// Initialize counter for row index in Excel
int index = 1;

// Loop through a specific range of cells
foreach (var cell in worksheet["B1:B4"])
{
    // Set formula to evaluate if the corresponding value in column A is 20 or greater
    cell.Formula = $"=IF(A{index}>=20, \"Pass\", \"Fail\")";
    index++;
}

// Save the modifications in a new Excel file
workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
```

<p class="list-decimal">7.2. Using the generated file from the previous example, we can get the Cell’s Formula:</p>

Below is the paraphrased section of the article with updated relative URL paths resolved to `ironsoftware.com`:

```cs
// Load the workbook from the current directory
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");

// Get the first worksheet from the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Iterate through the range B1 to B4 in the worksheet
foreach (var cell in worksheet["B1:B4"])
{
    // Output the formula set in each cell to the console
    Console.WriteLine(cell.Formula);
}

// Wait for a user input to proceed
Console.ReadKey();
```

### 3.8. Example of Trimming Cell Values ###

To demonstrate the trimming function, which removes all extraneous spaces from cell values, we've added a column to the "sum.xlsx" file. The following C# code snippet shows how to apply the trim function to clean up cell contents effectively.

```cs
// Load the workbook where you want to apply trimming
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
int i = 1;

// Iterate through the cells and apply the trim function to each
foreach (var cell in workSheet["F1:F4"])
{
    cell.Formula = "=TRIM(D" + i + ")";
    i++;
}

// Save the modified file
workBook.SaveAs("editedFile.xlsx");
```
This method allows for efficient cleanup of data entries where extra spaces might cause data integrity issues or discrepancies. Here's a visual representation of where to apply the function:

<a href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

By utilizing this approach, you can ensure that text data within your Excel documents is formatted consistently, eliminating any unwanted spaces.

<p class="list-description">To apply trim function (to eliminate all extra spaces in cells), I added this column to the sum.xlsx file</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description">And use this code</p>

Here's the paraphrased section of the code, with updated comments and resolved relative URL paths:

```cs
// Load the workbook from a specific location on your drive
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");

// Access the first worksheet in the workbook
WorkSheet worksheet = workbook.WorkSheets.First();

// Initialize a counting variable
int index = 1;

// Loop through a specified range of cells
foreach (var cell in worksheet["f1:f4"])
{
    // Apply the trim function to each cell to remove excess spaces in the referenced cells
    cell.Formula = $"=trim(D{index})"; 
    index++;  // Increment the index after each iteration
}

// Save the modified workbook under a new name
workbook.SaveAs("editedFile.xlsx");
```

This snippet retains the original mechanics but uses slightly different variable names and comments for better clarity and understanding of the code actions.

<p class="list-description">Thus, you can apply formulas in the same way.</p>

<hr class="separator">

## 4. Managing Excel Workbooks with Multiple Sheets ##

This section will guide you on handling Excel workbooks that contain various sheets.

### 4.1. Accessing Data from Various Sheets Within a Single Workbook ###

This segment demonstrates how you can efficiently read from different sheets within the same `.xlsx` file using IronXL. By specifying the exact sheet name rather than defaulting to the first sheet, you're able to process data from any specified sheet directly.

```cs
WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
var range = workSheet["A2:D2"];
foreach (var cell in range)
{
    Console.WriteLine(cell.Text);
}
```

<p class="list-description">I created an xlsx file that contains two sheets: “Sheet1”,” Sheet2”</p>
<p class="list-description">Until now we used WorkSheets.First() to work with the first sheet. In this example we will specify the sheet name and work with it</p>

Here's the paraphrased section with properly resolved URL paths and formatted markdown:

```cs
// Load an existing Excel workbook using the IronXL library
WorkBook workbook = IronXL.WorkBook.Load($"{Directory.GetCurrentDirectory()}\\Files\\testFile.xlsx");

// Access the second worksheet named 'Sheet2' from the workbook
WorkSheet worksheet = workbook.GetWorkSheet("Sheet2");

// Select a range of cells from A2 to D2
IronXL.Range selectedRange = worksheet["A2:D2"];

// Iterate through each cell in the range and print its text
foreach (var cell in selectedRange)
{
    Console.WriteLine("Text in cell: " + cell.Text);
}
```

This script defines how to load an Excel file, access a specific worksheet, extract a range of cells, and print their contents using IronXL, a robust library for handling Excel files in .NET environments without the need for Microsoft Excel.

### 4.2. Incorporating a New Worksheet into your Workbook ###

In this section, you'll learn how to extend a workbook by adding a new worksheet to it:

```cs
// Load an existing workbook
WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create a new worksheet named 'new_sheet'
WorkSheet newSheet = workbook.CreateWorkSheet("new_sheet");

// Assign a value to a cell in the new worksheet
newSheet["A1"].Value = "Hello World";

// Save the workbook with the newly added worksheet
workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

In this example, after loading an existing workbook, a new worksheet named `new_sheet` is created. The cell `A1` on this new sheet is set with the text "Hello World," demonstrating a simple write operation. Finally, the workbook, now containing the new sheet, is saved to a specified path. This function allows for dynamic expansion of workbooks to accommodate new data or analysis sheets as needed.

<p class="list-description">We can also add new sheet to a workbook:</p>

Here is the paraphrased section of the article that deals with C# code for working with Excel files using IronXL, resolving relative URL paths as required:

```cs
// Initialize the workbook by loading an existing file
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create a new worksheet called 'new_sheet'
WorkSheet worksheet = workbook.CreateWorkSheet("new_sheet");

// Set the value of cell A1 to "Hello World"
worksheet["A1"].Value = "Hello World";

// Save this workbook to a new location on your drive
workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
```

<hr class="separator">

## 5. Database Integration with Excel ##

Explore how to transfer data between a database and Excel sheets.

I established a database named "TestDb" which includes a table called "Country" featuring two columns: Id (integer, identity) and CountryName (string).

### 5.1. Populate an Excel Sheet with Data from a Database ###

In this section, you'll learn how to populate an Excel sheet directly from database data.

We'll create a new Excel sheet within a workbook and fill it up using data obtained from a database table called `Country`. Below is the step-by-step guide on how to achieve this task:

```cs
// Establish a connection to the database
TestDbEntities dbContext = new TestDbEntities();

// Load an existing workbook or create a new one
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Create a new worksheet named 'FromDb'
WorkSheet sheet = workbook.CreateWorkSheet("FromDb");

// Retrieve country data from the database
List<Country> countryList = dbContext.Countries.ToList();

// Set up the headers for the new Excel sheet
sheet.SetCellValue(0, 0, "Id");
sheet.SetCellValue(0, 1, "Country Name");

// Populate the Excel sheet with the data
int row = 1;
foreach (var item in countryList)
{
    sheet.SetCellValue(row, 0, item.id);
    sheet.SetCellValue(row, 1, item.CountryName);
    row++;
}

// Save the modifications to a new file
workbook.SaveAs("FilledFile.xlsx");
```

This script retrieves a list of countries from a database, then iteratively fills up an Excel worksheet with country IDs and names before saving the workbook. This way, you can automate the process of exporting database information into an Excel document, ensuring data is readily accessible and well-organized.

<p class="list-description">Here we will create a new sheet and fill it with data from the Country Table</p>

```cs
TestDbEntities dbContext = new TestDbEntities();
WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.AddWorkSheet("DataFromDb");
IEnumerable<Country> countryData = dbContext.Countries.ToList();

// Insert column headers
sheet["A1"].Value = "Id";
sheet["B1"].Value = "Country Name";

int rowIndex = 1;
foreach (Country country in countryData)
{
    sheet[$"A{rowIndex + 1}"].Value = country.Id;
    sheet[$"B{rowIndex + 1}"].Value = country.CountryName;
    rowIndex++;
}

// Save the modified document
workbook.SaveAs("DatabaseExported.xlsx");
```

### 5.2. Populate Database from an Excel Sheet ###

<p class="list-description">Load data into the TestDb database’s Country table from an Excel sheet.</p>

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

Here is the paraphrased section:

```cs
// Initialize database context for TestDb
TestDbEntities databaseContext = new TestDbEntities();

// Load the workbook from the specified location
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

// Access the specific worksheet named "Sheet3"
WorkSheet currentSheet = workbook.GetWorkSheet("Sheet3");

// Convert the worksheet data into a DataTable
System.Data.DataTable data = currentSheet.ToDataTable(true);

// Iterate through each data row in the DataTable
foreach (DataRow row in data.Rows)
{
    // Create a new Country object and set its name
    Country country = new Country();
    country.CountryName = row[1].ToString();

    // Add the new Country to the Countries collection in database context
    databaseContext.Countries.Add(country);
}

// Commit all changes made in the context to the database
databaseContext.SaveChanges();
```

<hr class="separator">

### Additional Resources

For those interested in deepening their understanding of IronXL, consider exploring the supplementary tutorials available in this segment, as well as the sample implementations on our main page, which generally provide sufficient introduction for most developers.

For detailed insights into the `WorkBook` class and more, our [API Reference](https://ironsoftware.com/csharp/excel/object-reference/) is an invaluable resource.

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

