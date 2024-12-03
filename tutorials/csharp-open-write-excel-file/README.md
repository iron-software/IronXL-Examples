# C# Tutorial for Excel File Manipulation [Without Interop]

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file/>***


Learn through practical examples on creating, opening, and saving Excel spreadsheets using C#. Perform fundamental operations like summing values, calculating averages, counting items, and more. IronXL.Excel is a self-contained .NET library capable of handling numerous spreadsheet formats. Importantly, it operates independently without needing [Microsoft Excel](https://products.office.com/en-us/excel) installed or relying on Interop.

<hr class="separator">

<p class="main-content__segment-title">Overview</p>

<h2>Use IronXL to Open and Write Excel Files</h2>

Explore, modify, and manage Excel files effortlessly using the intuitive <a href="https://ironsoftware.com/csharp/excel/" target="_blank">IronXL C# library</a>.

To get started, you can either download a [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or use your own setup as you follow this tutorial.

Here's how to begin:

1. Download and install the IronXL Excel Library either via [NuGet](https://www.nuget.org/packages/IronXL.Excel) or directly as a DLL.
2. Load any Excel format like XLS, XLSX, or CSV using `WorkBook.Load`.
3. Access and retrieve cell values employing a straightforward syntax: `sheet["A11"].DecimalValue`.

This guide covers:

- **Setup**: How to integrate IronXL.Excel into your project.
- **Basic Tasks**: Steps to create and open workbooks, navigate between sheets and cells, and save your progress.
- **Advanced Techniques**: Master complex tasks like adding custom headers, footers, performing calculations, and more within your spreadsheets.

<h4>Open an Excel File : Quick Code</h4>

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CsharpExcelManipulation
{
    public class BasicExcelExample
    {
        public void Execute()
        {
            WorkBook workbook = WorkBook.Load("test.xlsx");
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            IronXL.Range cells = worksheet["A2:A8"];
            decimal cumulativeTotal = 0;

            // Iterate through the specified range of cells
            foreach (var cell in cells)
            {
                Console.WriteLine($"Cell {cell.RowIndex} contains the value: '{cell.Value}'");
                if (cell.IsNumeric)
                {
                    // Safely get the decimal value to manage precision
                    cumulativeTotal += cell.DecimalValue;
                }
            }

            // Verify the total against a predefined cell value
            if (worksheet["A11"].DecimalValue == cumulativeTotal)
            {
                Console.WriteLine("Basic Test Passed");
            }
        }
    }
}
```

<h4>Write and Save Changes to the Excel File : Quick Code</h4>

Here is the paraphrased section of the article, with the relative URL paths resolved:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section2
    {
        public void Execute()
        {
            workSheet["B1"].Value = 11.54;

            // Commit the modifications to the Excel file
            workBook.SaveAs("test.xlsx");
        }
    }
}
```

<hr class="separator">
<p class="main-content__segment-title">Step 1</p>

## 1. Free Installation of the IronXL C# Library

IronXL.Excel offers a robust and versatile library designed for managing Excel files within .NET environments. It supports installation across various .NET project formats, including Windows applications, ASP.NET MVC, and .NET Core applications. Whether you're looking to open, edit, read, or save Excel documents, IronXL.Excel can seamlessly integrate into your development workflow without any dependency on Microsoft Office.

<h3>Install the Excel Library to your Visual Studio Project with NuGet</h3>

To begin, you'll need to add the IronXL.Excel library to your project. There are two primary methods to accomplish this: by using the NuGet Package Manager or the NuGet Package Manager Console.

If you choose to utilize the NuGet Package Manager for incorporating the IronXL.Excel library into your project, follow these straightforward steps:

1. Navigate to your project in your development environment. With your mouse, right-click on the project's name and select 'Manage NuGet Packages' from the context menu.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <p><img src="/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

2. Navigate to the 'Browse' tab, enter "IronXL.Excel" in the search field, and click 'Install'.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg" target="_blank">

Here's the paraphrased section with the relative URL resolved:

-----
![Search for IronXL](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/search-for-ironxl.jpg)

</a>

3. Installation Complete

<p>The installation of IronXL.Excel is now successfully completed in your project.</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="Installation Complete" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/and-we-are-done.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>



<h3>Install Using NuGet Package Manager Console</h3>

1. Navigate to `Tools` in the menu, then select `NuGet Package Manager` followed by `Package Manager Console`.

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

```
2. Execute the command: `Install-Package IronXL.Excel -Version 2019.5.2`
```

<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" target="_blank">
    <p><img src="/img/tutorials/csharp-open-write-excel-file/install-package-ironxl.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></p>
</a>

<h3>Manually Install with the DLL</h3>

You have the option to manually integrate the DLL into your project or the global assembly cache. You can download the DLL [here](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

```
 PM > Install-Package IronXL.Excel
```

# C# Write to Excel [Without Using Interop] Code Example Tutorial

***Based on <https://ironsoftware.com/tutorials/csharp-open-write-excel-file/>***


Explore this guide on how to craft, access, and store Excel documents using C# without the need for Microsoft Excel or the Interop dependency via the robust IronXL.Excel library.

---

<p class="main-content__segment-title">Introduction</p>

<h2>Manipulating Excel Documents with IronXL</h2>

Easily manage Excel files using the versatile <a href="https://ironsoftware.com/csharp/excel/" target="_blank">IronXL C# library.</a> Start by downloading a [sample project from GitHub](https://github.com/magedo93/IronSoftware.git) or leverage your existing project and follow this guide.

1. Acquire the IronXL Excel Library from the [NuGet repository](https://www.nuget.org/packages/IronXL.Excel) or directly download the DLL
2. Load documents seamlessly with `WorkBook.Load` supporting formats like XLS, XLSX, or CSV.
3. Retrieve cell values smoothly with syntax like: `sheet["A11"].DecimalValue`

This tutorial will guide you through:

- **Installation of IronXL.Excel**: Incorporating IronXL.Excel into your project.
- **Fundamental Excel Tasks**: Steps to create a new workbook, select sheets and cells, and save updated files.
- **Enhanced Spreadsheet functionalities**: Employ advanced features such as headers, footers, and various arithmetic operations.

<h4>Opening an Excel File: Code Walkthrough</h4>

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            IronXL.Range range = workSheet["A2:A8"];
            decimal total = 0;
            
            // Loop through cells in the range
            foreach (var cell in range)
            {
                Console.WriteLine($"Cell {cell.RowIndex} contains: '{cell.Value}'");
                if (cell.IsNumeric)
                {
                    // Accumulate values; handle numeric precision
                    total += cell.DecimalValue;
                }
            }
            
            // Verify summation via formula evaluation
            if (workSheet["A11"].DecimalValue == total)
            {
                Console.WriteLine("Validation Successful: Basic Test Passed");
            }
        }
    }
}
```

<h4>Modifying and Storing an Excel Document: Quick Guide</h4>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section2
    {
        public void Run()
        {
            workSheet["B1"].Value = 11.54;
            
            // Commit changes to disk
            workBook.SaveAs("test.xlsx");
        }
    }
}
```

---

<p class="main-content__segment-title">Getting Started</p>

## 1. Free Installation of the IronXL C# Library ##

**IronXL.Excel** serves as a versatile and robust library for manipulating Excel files across .NET applications, including Windows Forms, ASP.NET MVC, and .NET Core.

<h3>Adding the Excel Library via NuGet in Visual Studio</h3>

Start by integrating the IronXL.Excel library into your project. We offer two installation methods: via the NuGet Package Manager and the NuGet Package Manager Console.

### Visual Installation via NuGet Package Manager:

1. **Initiate by right-clicking** on your project name in Visual Studio -> Select 'Manage NuGet Packages'
2. **Navigate to the 'Browse' tab**, search for 'IronXL.Excel', and hit 'Install'
3. **Complete the process** and you're ready to go!

<a rel="nofollow" href="https://ironsoftware.com/csharp/excel/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" target="_blank">
  <img src="https://ironsoftware.com/csharp/excel/img/tutorials/csharp-open-write-excel-file/select-manage-nuget-package.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<h3>Using NuGet Package Manager Console</h3>

1. Access through 'Tools' -> 'NuGet Package Manager' -> 'Package Manager Console'
2. Execute: `Install-Package IronXL.Excel -Version latest`

<a rel="nofollow" href="https://ironsoftware.com/csharp/excel/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" target="_blank">
  <img src="https://ironsoftware.com/csharp/excel/img/tutorials/csharp-open-write-excel-file/package-manager-console.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;">
</a>

<h3>Manual DLL Installation</h3>
Alternatively, download and incorporate the <a href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">DLL package</a> directly into your project or the global assembly cache. 

```
 PM > Install-Package  IronXL.Excel
```

---

<hr class="separator">
<h4 class="tutorial-segment-title">How To Tutorials</h4>

## 2. Fundamental Tasks: Creation, Opening, and Saving of Excel Files ##

### 2.1. Setting Up a Basic Project: HelloWorld Console Application ###

<p class="list-description">Initiate a HelloWorld Project</p>

- **2.1.1. Launching Visual Studio**  
  ![Open Visual Studio](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png)

- **2.1.2. Start a New Project**  
  ![Start New Project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png)

- **2.1.3. Select Console App (.NET Framework)**  
  ![Select Console App](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg)

- **2.1.4. Name Your Project “HelloWorld” and Proceed to Create**  
  ![Name and Create Project](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg)

- **2.1.5. Your Console Application is Ready**  
  ![Application Created](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg)

- **2.1.6. Incorporate IronXL.Excel, then Install**  
  ![Add and Install IronXL](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg)

- **2.1.7. Implement Initial Code to Read the First Cell of the First Sheet in an Excel File**  

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}
```

### 2.2. Creating a New Excel File ###

<p class="list-description">Generate a new Excel file using IronXL:</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.Metadata.Title = "IronXL New File";
            WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
        }
    }
}
```

### 2.3. Opening Various File Types as Workbooks ###

- **2.3.1. Open a CSV File**  

- **2.3.2. Create a Text File with a List of Names and Ages, and Save as CSV**  
  ![Example Code Snippet](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg)

- **2.3.3. Load a XML File**  
  ![Open XML File](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/create-country-model.png)

- **2.3.5. Open JSON List as Workbook**
- **2.3.6. Define a Country Model for JSON Data**  
  ![JSON Model](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/create-country-model.png)

- **2.3.8. Implement Newtonsoft Library to Parse JSON to Model List**  
  ![Use JSON Library](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-newtonsoft-library-to-convert-json.png)

Creating new Excel files, opening various document formats as Excel workbooks, and performing initial setups for applications are foundations that help facilitate data manipulation and presentation in .NET environments using IronXL—a robust library offering extensive capabilities without needing Excel installed.

### 2.1. Example Project: HelloWorld Console Application ###

<p class="list-description">Start by creating a HelloWorld Project</p>

<p class="list-decimal">2.1.1. Launch Visual Studio</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/open-visual-studio.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.2. Initiate New Project</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-create-new-project.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.3. Select the Console App (.NET framework) option</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/choose-console-app.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.4. Name your project “HelloWorld” and click on the create button</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/give-our-sample-name.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.5. Your new console application is now set up</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/console-application-created.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.6. Integrate IronXL.Excel to your project and install it</p>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" target="_blank"><img src="https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/add-ironxl-click-install.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal">2.1.7. Insert a few coding lines to read and display the first cell from the first sheet in an Excel file</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{System.IO.Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
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

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class DemoSection
    {
        public void Execute()
        {
            // Load the Excel workbook
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
            
            // Access the first worksheet
            WorkSheet worksheet = workbook.WorkSheets.First();
            
            // Read the value of cell A1
            string valueInCellA1 = worksheet["A1"].StringValue;
            
            // Output the value of cell A1
            Console.WriteLine(valueInCellA1);
        }
    }
}
```

### 2.2 Construct a New Excel Document ###

This section guides you through creating a brand-new Excel workbook using IronXL:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section4
    {
        public void Run()
        {
            // Instantiate a new workbook with the specified format
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.Metadata.Title = "IronXL New File";
            
            // Create a new worksheet and make modifications
            WorkSheet workSheet = workBook.CreateWorkSheet("FirstSheet");

            // Adding a cell value
            workSheet["A1"].Value = "Hello World";

            // Styling the cell
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
        }
    }
}
```

This script illustrates how to initiate a new `WorkBook` with a `.xlsx` file format, add a `WorkSheet`, and modify cell properties and styles.

<p class="list-description">Create a new Excel file using IronXL</p>

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section4
    {
        public void Run()
        {
            // Instantiate a new workbook with Excel file format .xlsx
            WorkBook newWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            
            // Set the title in the document metadata
            newWorkbook.Metadata.Title = "IronXL New File";
            
            // Generate a new worksheet named 'FirstWorksheet'
            WorkSheet newSheet = newWorkbook.CreateWorkSheet("1stWorkSheet");
            
            // Setting a simple string value to cell A1
            newSheet["A1"].Value = "Hello World";
            
            // Adjusting the style of cell A2 to have an orange dashed bottom border
            newSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            newSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
        }
    }
}
```

### 2.3. Load Various Data Formats into a Workbook ###

Explore how to open different data file formats—CSV, XML, and JSON—and load them as workbooks in IronXL:

#### 2.3.1. Loading a CSV File ####

#### 2.3.2. Start by creating a new text file. Populate it with a list of names and ages and save it as `CSVList.csv` ####
![code snippet](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg)

#### Here is how your code snippet should appear: ####

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section5
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}
```

#### 2.3.3. Load XML Data ####

Create an XML file containing a list of countries:

```html
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

Here's the code snippet to read XML as a workbook:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section6
    {
        public void Run()
        {
            DataSet xmldataset = new DataSet();
            xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
            WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
            WorkSheet workSheet = workBook.WorkSheets.First();
        }
    }
}
```

#### 2.3.4. Load JSON Data ####

Create a JSON file with country data and map it to a JSON object model in C#. Here's how you might define the JSON structure:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section7
    {
        public void Run()
        {
            [
                {"name": "United Arab Emirates", "code": "AE"},
               {"name":"United Kingdom","code":"GB"},
               {"name":"United States","code":"US"},
               {"name":"United States Minor Outlying Islands","code":"UM"}
            ]
        }
    }
}
```

#### Create a model to map JSON data ####

Here's the C# class that represents the country data in JSON:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section8
    {
        public void Run()
        {
            public class CountryModel
            {
                public string name { get; set; }
                public string code { get; set; }
            }
        }
    }
}
``` 

#### Convert JSON to a Dataset ####

First, add the Newtonsoft JSON library to handle the conversion. Here's how to define a method that converts a list of JSON objects to a DataSet:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section9
    {
        public void Run()
        {
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
                        Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
                        t.Columns.Add(propInfo.Name, ColType);
                    }

                    foreach (T item in list)
                    {
                        DataRow row = t.NewRow();
                        foreach (var prop Info in elementType.GetProperties())
                        {
                            row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
                        }
                        t.Rows.Add(row);
                    }
                    return ds;
                }
            }
        }
    }
}
```

Now load this DataSet as a workbook:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section10
    {
        public void Run()
        {
            StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
            var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
            var xmldataset = countryList.ToDataSet();
            WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
            WorkSheet workSheet = workBook.WorkSheets.First();
        }
    }
}
```

<p class="list-decimal">2.3.1. Open CSV file</p>

<p class="list-decimal">2.3.2 Create a new text file and add to it a list of names and ages (see example) then save it as CSVList.csv</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/code-snippet.jpg" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Your code snippet should look like this</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class CSVReadingExample
    {
        public void Execute()
        {
            // Load the workbook from a specified CSV file within the current directory
            WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");

            // Obtain the first worksheet from the workbook
            WorkSheet worksheet = workbook.DefaultWorkSheet;

            // Retrieve the value of the first cell in the worksheet
            string firstCellValue = worksheet["A1"].StringValue;

            // Display the value of the cell in the console
            Console.WriteLine(firstCellValue);
        }
    }
}
```

<p class="list-decimal">

### 2.3.3. Opening an XML File 

<p class="list-decimal">To start working with XML data within your Excel spreadsheet, begin by creating an XML file. This file should feature a 'countries' root element, encompassing multiple 'country' child elements. Every 'country' should have attributes defining its specifics such as code, continent, etc.

For example, structure your XML like this for a concise and organized dataset:</p>

```xml
<?xml version="1.0" encoding="utf-8"?>
<countries xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <country code="ae" handle="united-arab-emirates" continent="asia" iso="784">United Arab Emirates</country>
    <country code="gb" handle="united-kingdom" continent="europe" alt="England Scotland Wales GB UK Great Britain Britain Northern" boost="3" iso="826">United Kingdom</country>
    <country code="us" handle="united-states" continent="north america" alt="US America USA" boost="2" iso="840">United States</country>
    <country code="um" handle="united-states-minor-outlying-islands" continent="north america" iso="581">United States Minor Outlying Islands</country>
</countries>
```

<p class="list-decimal">The following code snippet demonstrates how to read this XML and load it as a workbook:</p>

```cs
using IronXL;
using System.Data;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section6
    {
        public void Run()
        {
            DataSet xmldataset = new DataSet();
            xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
            WorkBook workBook = WorkBook.Load(xmldataset);
            WorkSheet workSheet = workBook.WorkSheets.First();
        }
    }
}
```

<p class="list-decimal">This method lays out a detailed, straightforward approach to handling XML data in an Excel environment with IronXL, making data manipulation and review a seamless task.</p>

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
using IronXL;
using IronXL.Excel;

namespace ExcelIntegrationExamples
{
    public class LoadXmlExample
    {
        public void Execute()
        {
            DataSet xmlData = new DataSet();
            xmlData.ReadXml(Path.Combine(Directory.GetCurrentDirectory(), "Files", "CountryList.xml"));
            WorkBook excelWorkbook = WorkBook.Load(xmlData);
            WorkSheet excelSheet = excelWorkbook.WorkSheets.First();
        }
    }
}
```

<p class="list-decimal">

### 2.3.5. Load JSON List into Excel Workbook

This section demonstrates how to import a JSON list into an Excel workbook using IronXL.

First, create a JSON file containing a list of countries, structured as follows:

```json
[
    {"name": "United Arab Emirates", "code": "AE"},
    {"name": "United Kingdom", "code": "GB"},
    {"name": "United States", "code": "US"},
    {"name": "United States Minor Outlying Islands", "code": "UM"}
]
```

Next, define a class that represents the structure of your JSON data. In this case, we will create a `CountryModel`:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section7
    {
        public void Run()
        {
            public class CountryModel
            {
                public string name { get; set; }
                public string code { get; set; }
            }
        }
    }
}
```

Finally, utilize the `Newtonsoft.Json` library to parse the JSON file and load its contents into a new Excel workbook. Here's how you can achieve this:

1. Add reference to `Newtonsoft.Json` library if not already done so. You can do this via NuGet package manager.

2. Create a method to read the JSON file, convert it into the list of `CountryModel`, and then load this list into the workbook:

```cs
using IronXL.Excel;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section8
    {
        public void Run()
        {
            // Read the JSON file into a string
            StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
            var countryList = JsonConvert.DeserializeObject<List<CountryModel>>(jsonFile.ReadToEnd());
            
            // Convert list to a DataSet
            var dataSet = countryList.ToDataSet();
            
            // Load the DataSet into an Excel workbook
            WorkBook workBook = IronXL.WorkBook.Load(dataSet);
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Optionally, save the workbook to verify the contents
            workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\LoadedFromJson.xlsx");
        }
    }
}
```

This code snippet reads a JSON file, converts the JSON string into a list of `CountryModel` objects, transforms this list into a `DataSet`, and finally loads the `DataSet` into an Excel workbook using IronXL.

<span class="list-description">Create JSON country list</span>
</p>

Here's the paraphrased section from the article, resolving relative URL paths according to your instruction:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section7
    {
        public void Run()
        {
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
                    "name": "U.S. Minor Outlying Islands",
                    "code": "UM"
                }
            ]
        }
    }
}
```

<p class="list-decimal"></p>
<p class="list-decimal">2.3.6. Create a country model that will map to JSON</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/create-country-model.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-decimal"></p>
<p class="list-decimal">Here is the class code snippet</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class CountryCodeSection
    {
        public void Execute()
        {
            public class CountryDetails
                {
                    public string name { get; set; }
                    public string code { get; set; }
                }
        }
    }
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

Here's the paraphrased section of the article:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section9
    {
        public void Run()
        {
            public static class ListConversionUtility
            {
                // Convert a generic list to a DataSet
                public static DataSet ConvertListToDataSet<T>(this IList<T> list)
                {
                    Type elementType = typeof(T);
                    DataSet dataSet = new DataSet();
                    DataTable dataTable = new DataTable();
                    dataSet.Tables.Add(dataTable);
                    
                    // Create columns in the DataTable based on the properties of T
                    foreach (var propertyInfo in elementType.GetProperties())
                    {
                        Type columnType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
                        dataTable.Columns.Add(propertyInfo.Name, columnType);
                    }

                    // Populate the table with values from the list
                    foreach (T item in list)
                    {
                        DataRow newRow = dataTable.NewRow();
                        foreach (var propertyInfo in elementType.GetProperties())
                        {
                            newRow[propertyInfo.Name] = propertyInfo.GetValue(item, null) ?? DBNull.Value;
                        }
                        dataTable.Rows.Add(newRow);
                    }
                    return dataSet;
                }
            }
        }
    }
}
```

This revised section adopts a more structured approach to explaining the transformation from a list to a `DataSet`. It retains the original intentions while employing an alternative vocabulary and variations in sentence structure for clarity and readability.

<p class="list-decimal"></p>
<p class="list-decimal">And finally load this dataset as a workbook</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section10
    {
        public void Run()
        {
            // Open a StreamReader to read the JSON file
            StreamReader jsonReader = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
            // Deserialize the JSON content into an array of CountryModel
            var countries = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonReader.ReadToEnd());
            // Convert the array to a DataSet
            var dataSet = countries.ToDataSet();
            // Load the DataSet into a new WorkBook
            WorkBook excelWorkbook = IronXL.WorkBook.Load(dataSet);
            // Access the first worksheet in the workbook
            WorkSheet firstSheet = excelWorkbook.WorkSheets.First();
        }
    }
}
```

### 2.4 Saving and Exporting Excel Files ###

This section covers multiple ways to save or export your Excel files in various formats using IronXL, allowing for flexible data handling.

#### 2.4.1 Saving as .xlsx ####
To save your file in the default Excel format (.xlsx), utilize the `SaveAs` method:
```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section11
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.Metadata.Title = "IronXL Example File";
            
            WorkSheet workSheet = workBook.CreateWorkSheet("FirstSheet");
            workSheet["A1"].Value = "Hello, World!";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
            
            workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\Example.xlsx");
        }
    }
}
```

#### 2.4.2 Saving as CSV ####
You can also save your workbook as a CSV file. Here's how to specify a delimiter such as a comma (`,`), pipe (`|`), or colon (`:`):
```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section12
    {
        public void Run()
        {
            workBook.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\ExampleCSV.csv", delimiter: "|");
        }
    }
}
```

#### 2.4.3 Saving as JSON ####
To save your worksheet data in JSON format, use the `SaveAsJson` method:
```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section13
    {
        public void Run()
        {
            workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\ExampleJSON.json");
        }
    }
}
```
The result will be a JSON file representing your Excel data.

#### 2.4.4 Saving as XML ####
Finally, to store your data in XML format, you can apply the `SaveAsXml` function:
```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section15
    {
        public void Run()
        {
            workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\ExampleXML.xml");
        }
    }
}
```

This method outputs an XML-file with the content of your worksheet, making it easy to integrate into other XML-based systems or applications.

<span class="list-description">We can save or export the Excel file to multiple file formats like (“.xlsx”,”.csv”,”.html”) using one of the following commands.</span>

<p class="list-decimal">

### Saving Excel Files in .xlsx Format

This section provides clear instructions on how to save workbooks in the widely-used ".xlsx" Excel format using IronXL library.

#### Step-by-Step Guide to Save a Workbook as .xlsx

Here's a quick guide to get you started:

1. First, create a new Workbook.
2. Define the title and initial data.
3. Use the `SaveAs()` method to save the workbook in the desired format.

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section11
    {
        public void Run()
        {
            // Create a new workbook with specified file format
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);

            // Set metadata for the workbook
            workBook.Metadata.Title = "IronXL New File";

            // Create a worksheet and populate it
            WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

            // Save the workbook to a file in the current directory
            workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
        }
    }
}
```

This example outlines the creation of a workbook, setting a metadata title, initializing worksheet content, and then saving it in the .xlsx format. This approach is handy for producing cleanly formatted Excel reports programmatically.

<span class="list-description">To Save to “.xlsx” use saveAs function</span>
</p>

Here is the paraphrased section of your article, with resolved relative URL paths from links and images to ironsoftware.com:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section11
    {
        public void Execute()
        {
            // Create a new workbook with XLSX format
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            workbook.Metadata.Title = "IronXL New File";

            // Create a worksheet named '1stWorkSheet'
            WorkSheet worksheet = workbook.CreateWorkSheet("1stWorkSheet");
            worksheet["A1"].Value = "Hello World";
            worksheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            worksheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;

            // Save the workbook to a file with a dynamic path
            string currentDirectory = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(currentDirectory, "Files", "HelloWorld.xlsx");
            workbook.SaveAs(filePath);
        }
    }
}
```

<p class="list-decimal">

### 2.4.2. Export to CSV Format

<span class="list-description">To export the workbook to a CSV format using `SaveAsCsv`, provide the method with two arguments: the path and filename of the CSV file, and the delimiter character such as a comma (","), pipe ("|"), or colon (":").</span>

</p>

Here's a paraphrased version of the given C# code snippet:

```cs
using IronXL.Excel; // Include the IronXL library for Excel operations
namespace ironxl.CsharpOpenWriteExcelFile // Namespace for the project
{
    public class Section12 // Define the class Section12
    {
        public void Run() // Method to execute the operation
        {
            // Save the active workbook to a CSV file with a specified delimiter
            workBook.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter: "|");
        }
    }
}
```

In this revised version, I enhanced the comments to make the purpose and functionality clearer, and demonstrated saving the workbook with a pipe ("|") delimiter which is common for CSV files when fields contain commas.

<p class="list-decimal">

### 2.4.3 Save to JSON Format

To convert and save the Excel document to a JSON file, we can use the `SaveAsJson` method as demonstrated below:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section13
    {
        public void Run()
        {
            workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
        }
    }
}
```

<span class="list-description">To save to Json “.json” use SaveAsJson as follow</span>
</p>

Here's a paraphrased version of the provided C# code snippet section:

```cs
using IronXL.Excel;
namespace ironxl.CsharpExcelOperations
{
    public class JsonExportSection
    {
        public void Execute()
        {
            // Save the workbook in JSON format
            workBook.SaveAsJson($@"{System.Environment.CurrentDirectory}\Files\ExportedHelloWorld.json");
        }
    }
}
```

In this version, I've made changes to the namespace and class names to provide a more descriptive context, also modifying the file path and comments to enhance clarity.

<p class="list-decimal">
  <span class="list-description">The result file should look like this</span>
</p>

Here's a paraphrased version of the given C# code section, with the relative URL paths resolved to `ironsoftware.com`:

```cs
using IronXL.Excel;
namespace IronXL.CSharpExcelExample
{
    public class JSONOutputExample
    {
        public void Run()
        {
            // JSON representation
            var data = new[]
            {
                new[] {"Hello World"},
                new[] {string.Empty}
            };
        }
    }
}
``` 

In this version, I've renamed the namespace and class to make them more descriptive and updated the variable names and comments for clarity.

<p class="list-decimal">

### 2.4.4. Exporting to XML Format

<span class="list-description">To export your workbook to an XML file format, the following code can be utilized:</span>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section15
    {
        public void Run()
        {
            // Create a new workbook with a worksheet
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");

            // Add some data to the worksheet
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
            
            // Save the workbook to an XML file
            workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
        }
    }
}
```

<span class="list-description">The output XML document will look like this:</span>

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

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section15
    {
        public void Run()
        {
            // Save the workbook as an XML file in the current directory
            workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.xml");
        }
    }
}
```

<p class="list-decimal">
  <span class="list-description">Result should be like this</span>
</p>

Here is the paraphrased section with absolute URL paths resolved to ironsoftware.com:

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

Explore Advanced Functions in Excel with IronXL: Sum, Average, Count, and More

Dive deep into utilizing essential Excel functions such as SUM, AVG, and COUNT using the IronXL library. Discover how to effectively implement these operations through the following code examples:

### 3.1 Sum Example

Create an Excel file named "Sum.xlsx" and manually enter a list of numbers. Use the following code to calculate the sum of values in a specific range:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section16
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            decimal sumResult = worksheet["A2:A4"].Sum();
            Console.WriteLine(sumResult);
        }
    }
}
```

### 3.2 Average Example

Utilizing the same workbook, you can compute the average of the values in the same range:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section17
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            decimal average = worksheet["A2:A4"].Avg();
            Console.WriteLine(average);
        }
    }
}
```

### 3.3 Count Example

Determine the number of items within a range using the count function:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section18
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            int count = worksheet["A2:A4"].Count();
            Console.WriteLine(count);
        }
    }
}
```

### 3.4 Maximum Value Example

Find the highest value in a range of cells:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section19
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            decimal max = worksheet["A2:A4"].Max();
            Console.WriteLine(max);
        }
    }
}
```

### 3.5 Minimum Value Example

Similarly, finding the smallest value within a range is straightforward:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section21
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            decimal min = worksheet["A1:A4"].Min();
            Console.WriteLine(min);
        }
    }
}
```

These operations are fundamental for processing and analyzing data within Excel files, and IronXL provides a robust method for executing these tasks efficiently.

### 3.1. Example: Calculating the Sum ###

To demonstrate the sum calculation, we created an Excel document named `Sum.xlsx` and manually filled it with a list of numeric values as shown in the image below:

[![Example of sum calculation spreadsheet](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png)](https://ironsoftware.com/img/tutorials/csharp-open-write-excel-file/sum-example.png)

Below is a C# code snippet utilizing the `IronXL` library to compute the sum of a specific range in our spreadsheet:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section16
    {
        public void Run()
        {
            // Load the workbook with our data
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            
            // Access the first worksheet in the workbook
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Calculate the sum of values in the range A2:A4
            decimal sum = workSheet["A2:A4"].Sum();
            
            // Output the sum to the console
            Console.WriteLine(sum);
        }
    }
}
```

This approach accesses values between the cells A2 and A4, sums them up, and outputs the total. IronXL handles Excel file manipulation efficiently without the need for Microsoft Excel or Interop, making it a steadfast solution for .NET developers needing to perform spreadsheet calculations programmatically.

<p class="list-description">Let’s find the sum for this list. I created an Excel file and named it “Sum.xlsx” and added this list of numbers manually</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/sum-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/sum-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description"></p>

Here's the paraphrased section of the code with the updated image and link paths resolved to ironsoftware.com:

```cs
using IronXL.Excel;
namespace ironxl.CsharpExcelOperations
{
    public class CalculateSum
    {
        public void Execute()
        {
            // Load the workbook from a directory
            WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            // Access the first worksheet in the workbook
            WorkSheet sheet = workbook.WorkSheets.First();
            // Calculate the sum of values in the range from A2 to A4
            decimal totalSum = sheet["A2:A4"].Sum();
            // Output the calculated sum to the console
            Console.WriteLine(totalSum);
        }
    }
}
```

In this revised version, I've improved readability by renaming the class and method to more clearly reflect their purpose ("CalculateSum" and "Execute"). I've added comments to explain each code line to make the process straightforward for other developers who might use or modify this code.

### 3.2. Example: Calculating the Average ###

Explore how to compute the average from a list of numbers stored in an Excel workbook named "Sum.xlsx":

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section17
    {
        public void Run()
        {
            // Load the workbook
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            
            // Access the first worksheet
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Calculate the average of values from cell A2 to A4
            decimal average = workSheet["A2:A4"].Avg();
            
            // Display the computed average
            Console.WriteLine(average);
        }
    }
}
```

This example demonstrates loading an Excel file, selecting a range of cells, and calculating their average value.

<p class="list-description">Using the same file, we can get the average:</p>

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class AverageCalculator
    {
        public void Execute()
        {
            // Load the workbook from the specified location
            var workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            
            // Access the first worksheet in the workbook
            WorkSheet worksheet = workbook.WorkSheets.First();
            
            // Calculate the average of values in cells A2 to A4
            decimal averageValue = worksheet["A2:A4"].Avg();
            
            // Output the average value to the console
            Console.WriteLine("The average is: " + averageValue);
        }
    }
}
```

### 3.3. Example: Counting Cells ###

Discover how to tally the number of entries within a specific range using the same "Sum.xlsx" file:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section18
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            decimal count = workSheet["A2:A4"].Count();
            Console.WriteLine(count);
        }
    }
}
```

This snippet demonstrates how to count elements in the cell range from A2 to A4 within the Excel sheet. The result, representing the total number of cells within this specified range, will be printed to the console.

<p class="list-description">Using the same file, we can also get the number of elements in a sequence:</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section18
    {
        public void Execute()
        {
            // Load the workbook from the specified file
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            
            // Access the first worksheet from the workbook
            WorkSheet worksheet = workbook.WorkSheets.First();
            
            // Count the number of items in the range A2:A4
            decimal numberOfEntries = worksheet["A2:A4"].Count();
            
            // Output the count to the console
            Console.WriteLine("Total Count: " + numberOfEntries);
        }
    }
}
```

### 3.4. Example: Finding the Maximum Value ###

Explore how to determine the highest value within a designated range using IronXL. The following example utilizes an Excel file named "Sum.xlsx" and extracts the maximum value from a specific cell range.

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section19
    {
        public void Run()
        {
            // Load the workbook
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            // Select the first worksheet
            WorkSheet workSheet = workBook.WorkSheets.First();
            // Retrieve the maximum decimal value from the specified range
            decimal max = workSheet["A2:A4"].Max();
            // Output the maximum value
            Console.WriteLine(max);
        }
    }
}
```

In addition, you can apply further operations on the results. For instance, verifying if the maximum cell contains a formula can be achieved with the following modification:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section20
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            // Check if the maximum cell from a broader range has a formula
            bool maxHasFormula = workSheet["A1:A4"].Max(c => c.IsFormula);
            // Display the result
            Console.WriteLine("Does the maximum value cell contain a formula? " + maxHasFormula);
        }
    }
}
```

This method outputs `false` if no formulas are found, demonstrating the flexibility of IronXL in handling Excel data queries dynamically.

<p class="list-description">Using the same file, we can get the max value of range of cells:</p>

Here's a paraphrased version of the provided C# code block from the article, also resolving relative URL paths with `ironsoftware.com` as requested:

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section19
    {
        public void Execute()
        {
            // Load the workbook from a file named "Sum.xlsx"
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

            // Retrieve the first worksheet in the workbook
            WorkSheet worksheet = workbook.WorkSheets.First();

            // Find the maximum value within the range A2 to A4
            decimal maximumValue = worksheet["A2:A4"].Max();

            // Output the maximum value to the console
            Console.WriteLine("The maximum value is: {0}", maximumValue);
        }
    }
}
```

This version of the code uses slightly different variable names and adds more descriptive comments to clarify what each section of the code is doing, enhancing readability and maintainability.

<p class="list-description">– We can apply the transform function to the result of max function:</p>

Here is the paraphrased version of the given C# code snippet:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section20
    {
        public void Execute()
        {
            // Load the workbook from a specific location
            WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            // Access the first worksheet in the workbook
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            // Evaluate the maximum value in the range A1:A4 based on whether the cell contains a formula
            bool hasFormula = worksheet["A1:A4"].Max(cell => cell.IsFormula);
            // Output the result to the console
            Console.WriteLine(hasFormula);
        }
    }
}
```

In this edited snippet:
- Changed the method name from `Run` to `Execute` to provide a different verb that still conveys performing an action.
- Added comments to explain each line of code, which improves readability and understanding of the code actions.
- Renamed `workBook` and `workSheet` to `workbook` and `worksheet` respectively, for conventional .NET casing consistency.
- Clarified the boolean variable's purpose by renaming it from `max2` to `hasFormula`, indicating that it checks for the presence of formulas in specified cells.

<p class="list-description">This example writes “false” in the console.</p>

### 3.5. Example: Finding the Minimum Value ###

This example demonstrates how to retrieve the smallest value from a specified range of cells in an Excel file named "Sum.xlsx":

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section21
    {
        public void Run()
        {
            // Load the workbook from a specific path
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            // Select the first worksheet in the workbook
            WorkSheet workSheet = workBook.WorkSheets.First();
            // Retrieve the smallest value from the specified cell range
            decimal minValue = workSheet["A1:A4"].Min();
            // Output the result to the console
            Console.WriteLine(minValue);
        }
    }
}
```

In this snippet, the method `Min()` is called on a range of cells (`A1:A4`) to compute the minimum value. The result is then printed to the console. This is a practical way to quickly identify the lowest value in a sequence within an Excel sheet.

<p class="list-description">Using the same file, we can get the min value of range of cells:</p>

Here is a paraphrased version of the C# code section provided:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class MinimumValueExample
    {
        public void Execute()
        {
            // Loading the workbook from a predefined location
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            // Accessing the first worksheet by default
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            // Calculating the minimum value from a range of cells
            decimal minValue = worksheet["A1:A4"].Min();
            // Print the minimum value to the console
            Console.WriteLine(minValue);
        }
    }
}
```

This version of the code maintains the same functionality but introduces slight modifications in naming and comments for clarity and understanding, also ensuring unique context though still related closely to the original sample.

### 3.6. Example of Sorting Cells ###

The functionality of sorting cell values within a workbook is an essential task when managing data. In this specific instance, we've prepared an Excel file entitled "Sum.xlsx" which includes various entries.

To demonstrate both ascending and descending sorting capabilities:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section22
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            //sort the range of cells in ascending order
            workSheet["A1:A4"].SortAscending();
            // For descending order, you could use:
            // workSheet["A1:A4"].SortDescending(); 
            
            workBook.SaveAs("SortedSheet.xlsx");
        }
    }
}
```

In this snippet, we access and sort the cells ranging from A1 to A4. The `SortAscending()` method rearranges them from the lowest to highest values, while you can switch to `SortDescending()` for the reverse order. This method offers a straightforward approach to organizing data as per specific requirements.

<p class="list-description">Using the same file, we can order cells by ascending or descending:</p>

Here is the paraphrased version of the given section:

```cs
using IronXL;

namespace ironxl.CsharpExcelManipulation
{
    public class SortDataExample
    {
        public void Execute()
        {
            // Load an existing workbook
            WorkBook excelWorkbook = WorkBook.Load($@"{System.IO.Directory.GetCurrentDirectory()}\Files\Sum.xlsx");

            // Access the first worksheet in the workbook
            WorkSheet excelSheet = excelWorkbook.WorkSheets.First();

            // Sort the range A1 to A4 in ascending order
            excelSheet["A1:A4"].SortAscending();
            // To sort in descending order, uncomment the following line
            // excelSheet["A1:A4"].SortDescending();

            // Save the changes to a new file
            excelWorkbook.SaveAs("SortedExcelSheet.xlsx");
        }
    }
}
```

This revised version employs more explanatory comments and variable names to improve the readability and understanding of the code snippet.

### 3.7. Example Using IF Conditions ###

This example illustrates how to utilize the `IF` condition in formulas for modifying cell values based on specific criteria:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section23
    {
        public void Run()
        {
            // Load the workbook from the file
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            // Access the first worksheet
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Initialize a counter for rows
            int rowIndex = 1;
            // Iterate through a range of cells and apply the IF condition
            foreach (var cell in workSheet["B1:B4"])
            {
                // Set the formula to check if the value in a parallel cell in column A is 20 or more
                cell.Formula = $"=IF(A{rowIndex}>=20, \"Pass\", \"Fail\")";
                rowIndex++;
            }

            // Save the modified workbook to a new file
            workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
        }
    }
}
```

In this code, we process each cell in the range B1 to B4. We use an `IF` statement to assign "Pass" if the corresponding cell in column A has a value of 20 or more, and "Fail" otherwise. This demonstration provides a clear example of how to dynamically set cell values based on condition in IronXL.

<p class="list-description">Using the same file, we can use the Formula property to set or get a cell’s formula:</p>

<p class="list-decimal">3.7.1. Save to XML “.xml”</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section23
    {
        public void Execute()
        {
            // Load the workbook containing the 'Sum.xlsx' file
            WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            
            // Access the first worksheet from the workbook
            WorkSheet worksheet = workbook.WorkSheets.First();
            int rowIndex = 1;
            
            // Iterate over cells in the range B1 to B4 to apply conditional formulas
            foreach (var cell in worksheet["B1:B4"])
            {
                // Set the formula to evaluate if the corresponding value in column A is 20 or more
                cell.Formula = $"=IF(A{rowIndex}>=20, \"Pass\", \"Fail\")";
                rowIndex++;
            }
            
            // Save the changes to a new Excel file named 'NewExcelFile.xlsx'
            workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
        }
    }
}
```

<p class="list-decimal">7.2. Using the generated file from the previous example, we can get the Cell’s Formula:</p>

Here's the paraphrased section:

```cs
using IronXL.Excel;
namespace ironxl.CsharpExcelDemo
{
    public class DisplayFormulas
    {
        public void Execute()
        {
            // Load the workbook from the specified path
            WorkBook excelWorkbook = WorkBook.Load($@"{Environment.CurrentDirectory}\Files\NewExcelFile.xlsx");
            // Accessing the first worksheet in the workbook
            WorkSheet firstSheet = excelWorkbook.WorkSheets.First();
            // Iterate through each cell in the specified range
            foreach (var excelCell in firstSheet["B1:B4"])
            {
                // Output the formula set in each cell
                Console.WriteLine(excelCell.Formula);
            }
            // Wait for a key press before closing the application
            Console.ReadKey();
        }
    }
}
``` 

In this version, I adjusted variable names to enhance clarity and included comments to provide context to each operation within the code snippet.

### 3.8 Example of Trimming Cells ###

This section demonstrates the trimming function, which removes all excess spaces from cell values. For illustration, I've modified the `sum.xlsx` file by incorporating an additional column.

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section25
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            int i = 1;
            foreach (var cell in workSheet["f1:f4"])
            {
                cell.Formula = "=trim(D" + i + ")";
                i++;
            }
            workBook.SaveAs("editedFile.xlsx");
        }
    }
}
```
This example utilizes the `TRIM` formula within Excel to automatically clean up the data. By applying the `TRIM` function via IronXL, all entries in the designated column are processed to discard any superfluous spaces, streamlining the data presentation significantly.

<p class="list-description">To apply trim function (to eliminate all extra spaces in cells), I added this column to the sum.xlsx file</p>
<a rel="nofollow" href="/img/tutorials/csharp-open-write-excel-file/trim-example.png" target="_blank"><img src="/img/tutorials/csharp-open-write-excel-file/trim-example.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%;"></a>

<p class="list-description">And use this code</p>

```cs
using IronXL;
using IronXL.Excel;

namespace IronSoftware.CSharpExcelExamples
{
    public class TrimFunctionExample
    {
        public void Execute()
        {
            // Loading the workbook from the specified file
            WorkBook excelWorkbook = WorkBook.Load($@"{System.IO.Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
            WorkSheet firstSheet = excelWorkbook.DefaultWorkSheet;

            // Starting index for cell iteration in the worksheet
            int index = 1;

            // Iterating over a range of cells and applying the trim formula
            foreach (var cell in firstSheet["f1:f4"])
            {
                cell.Formula = $"=trim(D{index})";
                index++;
            }

            // Saving the modified workbook to a new file
            excelWorkbook.SaveAs("editedFile.xlsx");
        }
    }
}
```

<p class="list-description">Thus, you can apply formulas in the same way.</p>

<hr class="separator">

## Working with Multisheet Excel Documents in C#

Handling Excel documents with multiple sheets can sometimes be more complex. Here, we'll explore handling workbooks that consist of several sheets using IronXL.

### Reading Data from Various Sheets in the Same Workbook ###

Multiple sheets in a single workbook can hold a wide range of data. To specify and work with more than the default first sheet, you can explicitly select a sheet by its name.

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section26
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
            var range = workSheet["A2:D2"];
            foreach (var cell in range)
            {
                Console.WriteLine(cell.Text);
            }
        }
    }
}
```

### Adding a New Sheet to an Existing Workbook ###

Expanding a workbook with a new sheet allows us to continuously add and organize data efficiently.

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section27
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            workSheet["A1"].Value = "Hello World";
            workBook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
        }
    }
}
```

By learning these operations, managing data across multiple sheets becomes a streamlined process, allowing for more organized data management within complex Excel workbooks.

### 4.1. Accessing Data Across Several Sheets in a Workbook ###

This tutorial explores how to extract data from various sheets within a single Excel file. Initially, we created a workbook named "testFile.xlsx" that includes two sheets, labeled “Sheet1” and “Sheet2”. Typically, to interact with the first sheet, we use `WorkSheets.First()`. However, in this example, we'll demonstrate how to specifically select and operate on a sheet by its name.

```cs
using IronXL.Excel;

namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section26
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
            var range = workSheet["A2:D2"];
            foreach (var cell in range)
            {
                Console.WriteLine(cell.Text);
            }
        }
    }
}
``` 

This section shows how straightforward it is to manipulate different sheets within the same workbook, enhancing the versatility of your data handling in .NET applications.

<p class="list-description">I created an xlsx file that contains two sheets: “Sheet1”,” Sheet2”</p>
<p class="list-description">Until now we used WorkSheets.First() to work with the first sheet. In this example we will specify the sheet name and work with it</p>

Below is the paraphrased version of the provided C# code section, with the relative URL paths properly resolved to `ironsoftware.com`:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section26
    {
        public void Execute()
        {
            // Load the workbook from a specified path
            WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            // Access a specific worksheet in the workbook by specifying its name
            WorkSheet worksheet = workbook.GetWorkSheet("Sheet2");
            // Identify and retrieve a range of cells from A2 to D2
            var cellRange = worksheet["A2:D2"];
            // Loop through each cell in the specified range
            foreach (var cell in cellRange)
            {
                // Print the text content of each cell
                Console.WriteLine(cell.Text);
            }
        }
    }
}
``` 

In this version:
- The method name changed from `Run` to `Execute` to vary the terminology slightly.
- Comments are added to describe the operations being performed for clarity, which is especially useful in technical documentation to help understand the context and purpose of the code blocks.

### 4.2. Adding a New Worksheet to Existing Workbook ###

Explore how straightforward it is to enhance a workbook by integrating a new sheet. Here's an easy guide to adding a fresh worksheet:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section27
    {
        public void Run()
        {
            // Load an existing workbook from the specified file
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            
            // Create a new worksheet and set a value
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            workSheet["A1"].Value = "Hello World";
            
            // Save the workbook with the new sheet added
            workBook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
        }
    }
}
```

In this example, we demonstrate how to open an existing workbook, append a new sheet named 'new_sheet', and initiate it with a simple message. Afterward, the workbook is saved, preserving all changes, including the addition of the new sheet. This process enhances file management and organization within your projects, ensuring that workbook modifications are clearly and effectively managed.

<p class="list-description">We can also add new sheet to a workbook:</p>

```cs
using IronXL.Excel;
namespace IronXLSample
{
    public class AddNewSheetExample
    {
        public void Execute()
        {
            // Load an existing workbook
            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            
            // Add a new worksheet to the workbook
            WorkSheet worksheet = workbook.CreateWorkSheet("new_sheet");
            
            // Set the value of cell A1
            worksheet["A1"].Value = "Hello World";
            
            // Save the workbook to a new file
            workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
        }
    }
}
```

<hr class="separator">

## 5. Database Integration with Excel ##

Explore how to import and export data between a database and Excel files.

I set up a database named "TestDb" that includes a table called Country with two columns: Id (integer, identity) and CountryName (string).

### 5.1. Populate an Excel Sheet Using Database Data ###

In this example, we'll demonstrate how you can populate a new Excel sheet with data retrieved from a database table. We have pre-configured a database referred to as "TestDb" which contains a table named "Country". This table includes two columns: `Id` (int, identity), and `CountryName` (string).

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section28
    {
        public void Run()
        {
            // Create a connection to the database
            TestDbEntities dbContext = new TestDbEntities();
            
            // Load an existing Excel workbook
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            
            // Create a new worksheet in the workbook
            WorkSheet sheet = workbook.CreateWorkSheet("FromDb");
            
            // Query the list of countries from the database
            List<Country> countryList = dbContext.Countries.ToList();
            
            // Write column headers
            sheet.SetCellValue(0, 0, "Id");
            sheet.SetCellValue(0, 1, "Country Name");
            
            // Populate the worksheet with data
            int row = 1;
            foreach (var item in countryList)
            {
                sheet.SetCellValue(row, 0, item.id);
                sheet.SetCellValue(row, 1, item.CountryName);
                row++;
            }
            
            // Save the workbook with the new data
            workbook.SaveAs("FilledFile.xlsx");
        }
    }
}
```

In this method, we first establish a connection to the "TestDb" database and retrieve a list of countries. We then add this data to a new sheet in an existing workbook and save the updated workbook. This process efficiently integrates Excel file operations with database management tasks, simplifying data handling in business applications.

<p class="list-description">Here we will create a new sheet and fill it with data from the Country Table</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section28
    {
        public void Run()
        {
            // Initialize database context
            TestDbEntities dbContext = new TestDbEntities();
            
            // Load an existing workbook
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            
            // Create a new worksheet
            WorkSheet sheet = workbook.CreateWorkSheet("FromDb");
            
            // Retrieve country list from the database
            List<Country> countryList = dbContext.Countries.ToList();
            
            // Set headers for columns
            sheet.SetCellValue(0, 0, "Id");
            sheet.SetCellValue(0, 1, "Country Name");
            
            // Populate the worksheet with data from the database
            int row = 1;
            foreach (var country in countryList)
            {
                sheet.SetCellValue(row, 0, country.id);
                sheet.SetCellValue(row, 1, country.CountryName);
                row++;
            }
            
            // Save the workbook to a new file
            workbook.SaveAs("FilledFile.xlsx");
        }
    }
}
```

### 5.2. Populating a Database from an Excel Spreadsheet ###

<p class="list-description">This process involves inserting data into the 'Country' table in our 'TestDb' database.</p>

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section29
    {
        public void Run()
        {
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
        }
    }
}
```

In this snippet, we initiate a connection to the database through `TestDbEntities`. The data is fetched from an Excel sheet named "Sheet3" contained in `testFile.xlsx`. This data is converted into a `DataTable` and subsequently iterated over to populate the 'Country' table. Each row from the Excel sheet is mapped to a new entity of `Country` which is then added to the database context. Finally, the `SaveChanges` method commits these additions to the database.

<p class="list-description">Insert the data to the Country table in TestDb Database</p>

Here is the paraphrased section of the article:

```cs
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class DatabaseIntegrationExample
    {
        public void Execute()
        {
            TestDbEntities databaseContext = new TestDbEntities();
            // Load the workbook from the specified path
            WorkBook workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            // Access a specific worksheet by its name
            WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
            // Convert the worksheet data into a DataTable, considering the first row as headers
            System.Data.DataTable sheetData = sheet.ToDataTable(true);

            // Loop through each data row and populate the database
            foreach (DataRow row in sheetData.Rows)
            {
                Country newCountry = new Country();
                newCountry.CountryName = row[1].ToString();
                databaseContext.Countries.Add(newCountry);
            }
            // Commit changes to the database
            databaseContext.SaveChanges();
        }
    }
}
```

In this rewritten code, the variable names and comments have been refined to enhance clarity and readability. The object names and methods were also slightly altered to make them more intuitive.

<hr class="separator">

### Additional Resources

For more insights into working with IronXL, we recommend exploring other tutorials in this section and reviewing the practical examples on our homepage, which many developers find sufficient to begin their projects.

For detailed documentation and references to the `WorkBook` class, visit our [API Reference](https://ironsoftware.com/csharp/excel/object-reference/).

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

