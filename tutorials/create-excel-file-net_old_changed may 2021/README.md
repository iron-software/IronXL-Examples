# C# Excel File Creation Guide

***Based on <https://ironsoftware.com/tutorials/create-excel-file-net_old_changed may 2021/>***


This guide provides a detailed walkthrough on how to generate an Excel Workbook on platforms compatible with .NET Framework 4.5 or .NET Core. The process of building Excel files in C# is straightforward and does not require the traditional **Microsoft.Office.Interop.Excel** library. Utilize IronXL to manage worksheet features such as freeze panes, security settings, printing configurations, and much more.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>How To Create Excel Files in C# .NET:</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-the-free-ironxl-c-library">Download the IronXL C# PDF Library</a></li>
        <li><a href="#anchor-3-create-an-excel-workbook">Create an Excel Workbook</a></li>
        <li><a href="#anchor-4-set-cell-values">Set Cell Values</a></li>
        <li><a href="#anchor-6-use-formulas-in-cells">Use Formulas in Cells</a></li>
        <li><a href="#anchor-7-apply-formatting">Apply Formatting, Set Print Properties, and more</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <a href="/downloads/assets/excel/tutorials/create-excel-file-net/tutorial-create-excel.pdf" target="_blank">
          <img style="box-shadow: none; width: 308px; height: 320px;" src="/img/tutorials/create-excel-file-net/how-to-create.svg" data-hover-src="/img/tutorials/create-excel-file-net/how-to-create-hover.svg" class="img-responsive learn-how-to-img replaceable-img">
        </a>
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h4 class="tutorial-segment-title">Overview</h4>

<h2>IronXL Creates C# Excel Files in .NET</h2>

[IronXL provides a seamless C# & VB Excel API](https://ironsoftware.com/csharp/excel/) designed for robust and efficient manipulation of Excel spreadsheets within .NET applications. Achieve high-speed operations without the necessity for MS Office or Excel Interop installations.

IronXL offers comprehensive support across multiple platforms and frameworks including .NET Core, .NET Framework, Xamarin, Mobile, Linux, macOS, and Azure, ensuring flexibility in deployment and integration.

<h3>IronXL Features:</h3>

Here's the paraphrased section with resolved links:

-----

- Direct access to human technical support from our team of .NET experts

- Quick and easy setup using Microsoft Visual Studio

- Complimentary for development purposes. Pricing for licenses starts at `$liteLicense`.

<h4>Create and Save an Excel File: Quick Code</h4>

<a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">https://www.nuget.org/packages/IronXL.Excel/</a>

Alternatively, you can download the [IronXL.Dll](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) directly and incorporate it into your project.

Here is the paraphrased section of the article with enhanced code comments and any relative URLs resolved to `ironsoftware.com`:

```cs
/**
 * Example: Creating and Saving an Excel File
 * Reference: anchor-create-and-save-an-excel-file
 **/
using IronXL;

// Instantiate a new workbook with default format as XLSX; this format can be customized
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX); 
var sheet = workbook.CreateWorkSheet("example_sheet");  // Add a worksheet named "example_sheet"

// Assign a simple string value to cell A1
sheet["A1"].Value = "Example";

// Bulk assign numerical value to a range A2 through A4
sheet["A2:A4"].Value = 5;

// Apply background color to cell A5
sheet["A5"].Style.SetBackgroundColor("#f0f0f0");

// Bold the font for a range of cells from A5 to A6
sheet["A5:A6"].Style.Font.Bold = true;

// Use a formula to sum values from A2 to A4 and set it in A6
sheet["A6"].Value = "=SUM(A2:A4)";

// Conditional check to confirm that the sum calculation is correct
if (sheet["A6"].IntValue == sheet["A2:A4"].IntValue)
{
    Console.WriteLine("Basic test passed");
}

// Persist the workbook to disk with the name 'example_workbook.xlsx'
workbook.SaveAs("example_workbook.xlsx");
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Acquire the IronXL C# Library at No Cost

### Installing the IronXL NuGet Package

There are primarily three methods to incorporate the IronXL NuGet package into your projects:

#### Via Visual Studio

Visual Studio seamlessly integrates NuGet Package Manager, which can be accessed from the Project Menu or by right-clicking your project in Solution Explorer as demonstrated below in Figures 3 and 4.

<center>
  <div style="display: inline-block; text-align: left; margin-bottom: 20px;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 3</strong> – <em>Project menu</em></p>
  </div>
</center>

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 4</strong> – <em>Right click Solution Explorer</em></p>
  </div>
</center>
<br></br>

Once you've engaged 'Manage NuGet Packages' through either approach, search and select the IronXL.Excel package for installation as illustrated in Figure 5.

<br></br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

#### Through Developer Command Prompt

Launch the Developer Command Prompt, located generally under your Visual Studio directory, and execute these steps:

- Enter the command: `PM > Install-Package IronXL.Excel`
- Hit Enter.
- Once installed, refresh your Visual Studio project.

#### Direct NuGet Package Download

To directly download the NuGet package, follow these stages:

1. Go to: [IronXL NuGet Package Page](https://www.nuget.org/packages/ironxl.excel/)
2. Click 'Download Package'.
3. Upon the package's download completion, double-click to proceed.
4. Refresh your Visual Studio project setup.

### Direct Library Installation

Alternatively, directly download the IronXL library via this URL: [Download IronXL Library](https://ironsoftware.com/csharp/excel/).

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library directly</em></p>
  </div>
</center>

After downloading, integrate the library into your project:

- Right-click on the Solution in Solution Explorer.
- Click 'References', browse to your downloaded IronXL.dll library.
- Press OK.

Ready to explore IronXL's capabilities? Let's dive in!

<h3>Install by Using NuGet</h3>

Here are three effective approaches to install the IronXL NuGet package:

### 1. Visual Studio Installation

The NuGet Package Manager within Visual Studio makes it straightforward to add IronXL to your projects. Simply navigate through the Project Menu or right-click your project in the Solution Explorer to locate the NuGet Package Manager option. 

### 2. Using Developer Command Prompt

To install the IronXL.Excel package via the Developer Command Prompt, follow these steps:
- Locate the Developer Command Prompt, which can typically be found in the Visual Studio installation directory.
- Execute the command: `Install-Package IronXL.Excel`
- Hit Enter to begin the installation.
- Once installed, ensure to reload your project in Visual Studio.

### 3. Direct NuGet Package Download

If you prefer to download the IronXL NuGet package manually:
- Visit the NuGet package page: [IronXL.Excel NuGet Package](https://www.nuget.org/packages/ironxl.excel/)
- Click the 'Download Package' button to download the `.nupkg` file.
- After downloading, double-click to open it and integrate it with your project.
- Reload your Visual Studio project to apply changes. 

Each method provides a flexible way to integrate IronXL based on your development setup preferences.

<h3>Visual Studio</h3>

Visual Studio comes equipped with a NuGet Package Manager, allowing you to incorporate NuGet packages into your projects easily. You can find this feature in the Project Menu, or by simply right-clicking your project within the Solution Explorer. Details on how to access these options are illustrated in Figures 3 and 4 below.

<center>
  <div style="display: inline-block; text-align: left; margin-bottom: 20px;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/project-menu.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/project-menu.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 3</strong> – <em>Project menu</em></p>
  </div>
</center>

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/right-click-solution-explorer.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/right-click-solution-explorer.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 4</strong> – <em>Right click Solution Explorer</em></p>
  </div>
</center>
<br></br>

Once you've accessed the Manage NuGet Packages through either method described, navigate to and select the `IronXL.Excel` package to install it as demonstrated in Figure 5.

<br></br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

<h3>Developer Command Prompt</h3>

Access the Developer Command Prompt and execute these instructions to install the IronXL.Excel NuGet package:

1. Locate your Developer Command Prompt, typically found within your Visual Studio directory.
2. Enter the command below:
3. `PM > Install-Package IronXL.Excel`
4. Hit the Enter key.
5. Upon pressing Enter, the package will be successfully installed.
6. Refresh your Visual Studio project to complete the installation process.

<h3>Download the NuGet Package directly</h3>

To download the NuGet package, follow these instructions:

1. Go to the URL: [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/)

2. Select "Download Package"

3. Once the download is complete, double-click the downloaded file

4. Restart your Visual Studio project to apply the changes

</br>
<h3>Install IronXL by Direct Download of the Library</h3>

The alternative method to install IronXL involves directly downloading it from the following link: [https://ironsoftware.com/csharp/excel/](https://ironsoftware.com/csharp/excel/).

</br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library</em></p>
  </div>
</center>

To include the IronXL library in your project, follow these straightforward steps:

1. In the Solution Explorer, right-click on the Solution.
2. Choose 'References' from the context menu.
3. Look for the IronXL.dll file by browsing.
4. Confirm your selection by clicking 'OK'.

<h3>Let's Go!</h3>

Now that everything's ready, let's explore all the fantastic capabilities of the IronXL library!

<hr class="separator">

<h4 class="tutorial-segment-title">How to Tutorials</h4>

## 2. Setting Up an ASP.NET Project ##

To kick off with your ASP.NET project:

1. Firstly, go to the NuGet package page for IronXL at [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/).
2. Download the package by clicking on the "Download Package" button.
3. Once the download is complete, open the file which will automatically integrate it into your Visual Studio environment.

Now, let's start creating an ASP.NET Website:

1. Launch Visual Studio.
2. Navigate to `File > New Project`.
3. Choose 'Web' under 'Visual C#' in the project types.
4. Opt for an 'ASP.NET Web Application' as shown in the following image:
   
   ![New ASP.NET Project Setup](https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png "Figure 1 – New Project Setup")

5. Click 'OK'.
6. On the following screen, select 'Web Forms' as illustrated below:

   ![Select Web Forms](https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png "Figure 2 – Select Web Forms Option")
   
7. Confirm your selection by clicking 'OK'.

Once these steps are completed, you're ready to incorporate IronXL into your project and start customizing your Excel files effortlessly.

<ul>
  <li>Navigate to the following URL:</li>
  <li><a href="https://www.nuget.org/packages/ironxl.excel/" target="_blank">https://www.nuget.org/packages/ironxl.excel/</a></li>
  <li>Click on Download Package</li>
  <li>After the package has downloaded, double click it</li>
  <li>Reload your Visual Studio project</li>
</ul>

Here is the paraphrased version of the specified section with resolved URL paths:

-----

Follow these steps to initiate an ASP.NET website:

1. Launch Visual Studio.

2. Navigate to the File menu and choose 'New Project'.

3. In the Project type list, select 'Web' which is located under Visual C#.

4. Choose 'ASP.NET Web Application' as illustrated below:

![ASP.NET New Project](https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png)

-----

The URL and images paths have been resolved to point to the absolute path as per your request.

<br></br>
<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 1</strong> — *New Project*
```

<a href="https://ironsoftware.com/csharp/excel/" target="_blank">ironsoftware.com/csharp/excel/</a>

5. Confirm your selection by clicking OK.

6. In the subsequent screen, choose the Web Forms option as illustrated in [Figure 2](https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png) below.

<br></br>

<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

**Figure 2** – _Web Forms_
```

<br></br>

7\. Click OK

With that step complete, we're ready to roll! It's time to install IronXL and begin customizing your Excel file.

<hr class="separator">

```cs
/**
Create Excel Workbook effortlessly with IronXL
**/
WorkBook newWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL makes it incredibly straightforward to initialize a new Excel Workbook. As you can see, it just takes a single line of code! This feature underscores the simplicity and power of the IronXL library, making .NET developers' tasks much more manageable. Whether you're using older Excel file formats like XLS or the more current XLSX format, IronXL supports them all seamlessly.

Here is the paraphrased section of the article with resolved URLs:

```cs
// Initialize a new Workbook object in the XLSX format
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL supports the creation of both XLS (the traditional Excel file format) and XLSX (the modern Excel file format).

### 3.1. Establishing a Default Worksheet ###

Creating a default worksheet is even more straightforward:

```cs
var sheet = workbook.CreateWorkSheet("Budget2020");
```

In the code example above, the term "Sheet" refers to the worksheet. It's a versatile component that allows you to modify cell values along with virtually every other capability available in Excel.

To clarify, a Workbook is essentially a collection of Worksheets. You have the flexibility to include multiple Worksheets within a single Workbook, details of which will be covered in a forthcoming article. Within each Worksheet, you will find Rows and Columns. The point where a Row intersects with a Column is known as a Cell, which will be your primary area of interaction when managing data in Excel.

<hr class="separator">

## 4. Assign Values to Cells ##

### 4.1. Manually Assigning Cell Values ###

To manually input values to specific cells, simply specify the cell's address and assign its content, as shown below:

```cs
/**
Manually Assign Values to Cells
anchor-manually-assign-values
**/
sheet["A1"].Value = "January";
sheet["B1"].Value = "February";
sheet["C1"].Value = "March";
sheet["D1"].Value = "April";
sheet["E1"].Value = "May";
sheet["F1"].Value = "June";
sheet["G1"].Value = "July";
sheet["H1"].Value = "August";
sheet["I1"].Value = "September";
sheet["J1"].Value = "October";
sheet["K1"].Value = "November";
sheet["L1"].Value = "December";
```

In this example, cells from A1 to L1 across the first row are populated with the names of months.

### 4.2. Dynamically Setting Cell Values ###

Dynamically setting cell values avoids hardcoding specific cell addresses by using a loop to generate and assign values:

```cs
/**
Dynamically Assign Values to Cells
anchor-dynamically-assign-values
**/
Random randomValueGenerator = new Random();
for (int i = 2; i <= 11; i++)
{
    sheet["A" + i].Value = randomValueGenerator.Next(1, 1000);
    sheet["B" + i].Value = randomValueGenerator.Next(1000, 2000);
    sheet["C" + i].Value = randomValueGenerator.Next(2000, 3000);
    sheet["D" + i].Value = randomValueGenerator.Next(3000, 4000);
    sheet["E" + i].Value = randomValueGenerator.Next(4000, 5000);
    sheet["F" + i].Value = randomValueGenerator.Next(5000, 6000);
    sheet["G" + i].Value = randomValueGenerator.Next(6000, 7000);
    sheet["H" + i].Value = randomValueGenerator.Next(7000, 8000);
    sheet["I" + i].Value = randomValueGenerator.Next(8000, 9000);
    sheet["J" + i].Value = randomValueGenerator.Next(9000, 10000);
    sheet["K" + i].Value = randomValueGenerator.Next(10000, 11000);
    sheet["L" + i].Value = randomValueGenerator.Next(11000, 12000);
}
```

Every cell from A2 to L11 in this illustration has a unique value randomly generated.

### 4.3. Importing Data from a Database into Cells ###

Here's how you can populate cells with data fetched from a database:

```cs
/**
Import Data from Database
anchor-import-from-database
**/
// Setup database connectivity:
string connectionString;
string sqlQuery;
DataSet dataSet = new DataSet("ExampleData");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Define Connection String
connectionString = @"Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserID;Password=Password";

// Define SQL Query for data
sqlQuery = "SELECT ColumnNames FROM TableName";

// Establish connection and fill DataSet
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(sqlQuery, sqlConnection);
sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Loop through DataSet to assign values to cells
foreach (DataTable table in dataSet.Tables)
{
    int rowCount = table.Rows.Count - 1;

    for (int rowIndex = 12; rowIndex <= 21; rowIndex++)
    {
        sheet["A" + rowIndex].Value = table.Rows[rowCount]["ColumnName1"].ToString();
        sheet["B" + rowIndex].Value = table.Rows[rowCount]["ColumnName2"].ToString();
        // Continue for other columns as needed
    }
    rowCount++;
}
```

Cells are filled directly with values from the database by setting the `Value` property to the appropriate data field.

### 4.1. Manual Cell Value Input ###

You can assign values directly to specific cells simply by specifying the cell coordinates and the value to be set. Here's how you achieve this:

```cs
/**
Manual Value Entry for Cells
anchor-manual-entry-for-cells
**/
sheet["A1"].Value = "January";
sheet["B1"].Value = "February";
sheet["C1"].Value = "March";
sheet["D1"].Value = "April";
sheet["E1"].Value = "May";
sheet["F1"].Value = "June";
sheet["G1"].Value = "July";
sheet["H1"].Value = "August";
sheet["I1"].Value = "September";
sheet["J1"].Value = "October";
sheet["K1"].Value = "November";
sheet["L1"].Value = "December";
```

In the provided code snippet, each cell from `A1` to `L1` is being assigned the name of a different month.

Here's the paraphrased section, with resolved relative URL paths and enhanced comments in the code snippet for clarity:

```cs
/**
Manually Assign Values to Cells
anchor-manual-cell-assignments
**/
// Assign names of months to cells A1 through L1
sheet["A1"].Value = "January";   // Set value of cell A1
sheet["B1"].Value = "February";  // Set value of cell B1
sheet["C1"].Value = "March";     // Set value of cell C1
sheet["D1"].Value = "April";     // Set value of cell D1
sheet["E1"].Value = "May";       // Set value of cell E1
sheet["F1"].Value = "June";      // Set value of cell F1
sheet["G1"].Value = "July";      // Set value of cell G1
sheet["H1"].Value = "August";    // Set value of cell H1
sheet["I1"].Value = "September"; // Set value of cell I1
sheet["J1"].Value = "October";   // Set value of cell J1
sheet["K1"].Value = "November";  // Set value of cell K1
sheet["L1"].Value = "December";  // Set value of cell L1
```

This rewritten code block achieves the same functionality but with more explicit commenting that enhances understanding.

In this example, I have filled Columns A through L with the names of different months in the first row of each column.

### 4.2. Dynamically Assign Cell Values ###

Assigning values to cells dynamically offers a flexible approach similar to the method discussed earlier. What sets this method apart is that it eliminates the need to specify exact cell coordinates in advance. In the following example, you'll see how to instantiate a new `Random` object for generating random numbers. Then, leveraging a `for` loop, you'll iterate over a specified range of cells, filling each one with a value.

```cs
// Setting Cell Values Dynamically
// anchor-set-cell-values-dynamically
Random randomizer = new Random();
for (int index = 2; index <= 11; index++)
{
    // Assign random values for each column from A to L
    sheet["A" + index].Value = randomizer.Next(1, 1000);
    sheet["B" + index].Value = randomizer.Next(1000, 2000);
    sheet["C" + index].Value = randomizer.Next(2000, 3000);
    sheet["D" + index].Value = randomizer.Next(3000, 4000);
    sheet["E" + index].Value = randomizer.Next(4000, 5000);
    sheet["F" + index].Value = randomizer.Next(5000, 6000);
    sheet["G" + index].Value = randomizer.Next(6000, 7000);
    sheet["H" + index].Value = randomizer.Next(7000, 8000);
    sheet["I" + index].Value = randomizer.Next(8000, 9000);
    sheet["J" + index].Value = randomizer.Next(9000, 10000);
    sheet["K" + index].Value = randomizer.Next(10000, 11000);
    sheet["L" + index].Value = randomizer.Next(11000, 12000);
}
```

Every cell between A2 and L11 is filled with a unique, randomly generated number.

Shifting our focus to dynamic data insertion, let’s explore how you can directly populate cells with data from a database. The following code snippet demonstrates this process, provided that your database connections are properly configured.

### 4.3. Import Data Directly from a Database ###

Integrating data directly from a database into your Excel workbook is seamless with IronXL. Below, you'll find the steps to populate your spreadsheet using database records:

```cs
/**
Populate Cells from Database
anchor-import-directly-from-a-database
**/
// Initialize database connection and dataset
string connectionString;
string sqlCommand;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Configuration for the database connection
connectionString = @"Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserID;Password=UserPassword";

// SQL query to fetch data
sqlCommand = "SELECT ColumnNames FROM TableName";

// Connecting and filling the dataset with data
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(sqlCommand, sqlConnection);

sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Populate worksheet cells with dataset contents
foreach (DataTable table in dataSet.Tables)
{
    int rowIndex = table.Rows.Count - 1;

    for (int i = 12; i <= 21; i++)
    {
       sheet["A" + i].Value = table.Rows[rowIndex]["ColumnName1"].ToString();
       sheet["B" + i].Value = table.Rows[rowIndex]["ColumnName2"].ToString();
       sheet["C" + i].Value = table.Rows[rowIndex]["ColumnName3"].ToString();
       sheet["D" + i].Value = table.Rows[rowIndex]["ColumnName4"].ToString();
       sheet["E" + i].Value = table.Rows[rowIndex]["ColumnName5"].ToString();
       sheet["F" + i].Value = table.Rows[rowIndex]["ColumnName6"].ToString();
       sheet["G" + i].Value = table.Rows[rowIndex]["ColumnName7"].ToString();
       sheet["H" + i].Value = table.Rows[rowIndex]["ColumnName8"].ToString();
       sheet["I" + i].Value = table.Rows[rowIndex]["ColumnName9"].ToString();
       sheet["J" + i].Value = table.Rows[rowIndex]["ColumnName10"].ToString();
       sheet["K" + i].Value = table.Rows[rowIndex]["ColumnName11"].ToString();
       sheet["L" + i].Value = table.Rows[rowIndex]["ColumnName12"].ToString();
       rowIndex++;
    }
}
```

By setting the `Value` property of each cell, you can dynamically enter the specific field values directly into your spreadsheet cells. The ease of integration makes IronXL a powerful tool for instant data management and reporting from your databases.

Here's the paraphrased section of the article, with improvements in code comments and changes to some code snippets, plus resolution of image and link URLs to `ironsoftware.com`:

```cs
/**
Populate Cells from Database
anchor-populate-cells-from-database
**/
// Initialize data access components to fetch data from a database
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection connection;
SqlDataAdapter dataAdapter;

// Configure the database connection string
connectionString = @"Data Source=Server_Name;Initial Catalog=Database_Name;User ID=User_ID;Password=Password";

// Define the SQL query to retrieve data
query = "SELECT Column_Names FROM Your_Table";

// Establish the database connection and fill the dataset
connection = new SqlConnection(connectionString);
dataAdapter = new SqlDataAdapter(query, connection);

connection.Open();
dataAdapter.Fill(dataSet);

// Iterate over the dataset and populate Excel cells
foreach (DataTable table in dataSet.Tables)
{
    int rowCount = table.Rows.Count - 1;

    // Populate cells from A12 to L21 with the retrieved data
    for (int row = 12; row <= 21; row++)
    {
        sheet["A" + row].Value = table.Rows[rowCount]["Field1"].ToString();
        sheet["B" + row].Value = table.Rows[rowCount]["Field2"].ToString();
        sheet["C" + row].Value = table.Rows[rowCount]["Field3"].ToString();
        sheet["D" + row].Value = table.Rows[rowCount]["Field4"].ToString();
        sheet["E" + row].Value = table.Rows[rowCount]["Field5"].ToString();
        sheet["F" + row].Value = table.Rows[rowCount]["Field6"].ToString();
        sheet["G" + row].Value = table.Rows[rowCount]["Field7"].ToString();
        sheet["H" + row].Value = table.Rows[rowCount]["Field8"].ToString();
        sheet["I" + row].Value = table.Rows[rowCount]["Field9"].ToString();
        sheet["J" + row].Value = table.Rows[rowCount]["Field10"].ToString();
        sheet["K" + row].Value = table.Rows[rowCount]["Field11"].ToString();
        sheet["L" + row].Value = table.Rows[rowCount]["Field12"].ToString();
    }
    rowCount++;
}
```

You only need to assign the `Value` property of the specified cell with the field name that you want to enter into the cell.

<hr class="separator">

## 5. Formatting Techniques ##

### 5.1. Adjusting Cell Background Colors ###

Easily modify the background color of a single cell or a group of cells using this simple line of code:

```cs
/**
Configure Cell Background Color
anchor-configure-cell-background-colors
**/
sheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

The code snippet above sets the selected cells' background color to a shade of gray. The color code follows the RGB format, represented in hexadecimal code, where the first two characters represent the values for Red, followed by Green, and then Blue, ranging from 00 to FF.

### 5.2. Bordering Cells ###

Creating borders around cells can be performed with minimal effort, demonstrated in the following example:

```cs
/**
Define Borders
anchor-define-borders
**/
sheet["A1:L1"].Style.TopBorder.SetColor("#000000");
sheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

sheet["L2:L11"].Style.RightBorder.SetColor("#000000");
sheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

sheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
sheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

This code sets black top and bottom borders for cells A1 to L1, establishes a medium weight right border for cells L2 to L11, and configures a medium weight bottom border for cells A11 to L11.

Using these methods, adding aesthetic enhancements and clarity to your spreadsheet data with IronXL becomes straightforward and efficient.

```cs
// Set the Background Color of Cells
sheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

Here, we have updated the background color of cells from A1 to L1 to a shade of gray. The color code `#d3d3d3` utilizes the RGB hex system where each pair of digits represents the intensity of red, green, and blue respectively, ranging from `00` to `FF` (least to most intense).

```cs
/**
Adjust Cell Background Color
anchor-adjust-background-colors-of-cells
**/
sheet ["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This code assigns a gray background color to a specific range of cells. The color is specified in RGB (Red, Green, Blue) hexadecimal format, with the initial two characters denoting Red, the subsequent two representing Green, and the final two indicating Blue. The hexadecimal values span from 0 to 9 and from A to F.

### 5.2. Define Borders ###

Setting up borders in cells using IronXL is straightforward and effective, as demonstrated below:

Here's the paraphrased section of the code:

```cs
/**
Define Borders for Spreadsheet Cells
anchor-set-up-borders
**/
// Setting up top and bottom borders for the first row with black color
sheet["A1:L1"].Style.TopBorder.SetColor("#000000");
sheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Configure the right border for cells from L2 to L11 with a medium thickness
sheet["L2:L11"].Style.RightBorder.SetColor("#000000");
sheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Implementing a medium thickness bottom border from A11 to L11
sheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
sheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

This version maintains the essence of your original code while changing the annotations and adjusting the format for clarity and emphasis on actions taken on different cell ranges.

In the provided code example, black borders are assigned to the top and bottom of cells ranging from A1 to L1. Additionally, the right border of cells from L2 to L11 has been defined with a medium thickness. Finally, a medium thickness border is also applied to the bottom of cells from A11 to L11.

<hr class="separator">

## 6. Utilizing Formulas in Cells ##

IronXL simplifies the process of using formulas within your spreadsheets to an astonishing degree. Here's a demonstration of how effortlessly you can integrate mathematical functions into your cells:
```

Here's the paraphrased section of the article with resolved relative URL paths where needed:

```cs
/**
Implementing Formulas in Spreadsheet Cells
anchor-implement-formulas-in-cells
**/
decimal total = sheet["A2:A11"].Sum();
decimal average = sheet["B2:B11"].Avg();
decimal maximum = sheet["C2:C11"].Max();
decimal minimum = sheet["D2:D11"].Min();

// Assigning results to designated cells.
sheet["A12"].Value = total;
sheet["B12"].Value = average;
sheet["C12"].Value = maximum;
sheet["D12"].Value = minimum;
```

The engaging aspect of this is that it allows you to define the cell's data type, determining the output of mathematical operations like summing, averaging, or finding maximum or minimum values. The provided script demonstrates how easy it is to apply formulas such as SUM, AVG, MAX, and MIN to process data accordingly.

<hr class="separator">

## Configuring Worksheet and Print Settings

### Worksheet Customizations

Enhance your worksheet's usability by implementing protective measures and user interface adjustments using IronXL. Here’s how you can do it:

```cs
// Protect your worksheet with a password and freeze the top row to enhance usability
sheet.ProtectSheet("Password");
sheet.CreateFreezePane(0, 1);
```

This code snippet enables the first row to remain static while scrolling, making constant header visibility possible. Simultaneously, it protects your worksheet with a password, limiting modifications to authorized personnel only.

![Freeze Panes](https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png "Illustration of Freeze Panes Feature")

*Figure 7 – Demonstrating Freeze Panes*

![Protected Worksheet](https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png "Protected Worksheet")

*Figure 8 – Worksheet Protection in Action*

### Print Configuration Options

IronXL facilitates comprehensive control over how your documents are printed, including paper orientation, size, and the print area:

```cs
// Set up specific print areas and properties for optimized printing
sheet.SetPrintArea("A1:L12");
sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

By specifying the print area (`"A1:L12"`), changing the orientation to landscape, and selecting A4 as the paper size, this setup ensures that your Excel data prints perfectly every time.

![Print Setup](https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png "Print Setup")

*Figure 9 – Setting Up Print Properties*

These settings are pivotal to managing how the workbook appears when printed, ensuring a professional presentation of your data.

### 7.1. Configure Worksheet Settings ###

Configuring worksheet settings involves actions such as freezing certain rows and columns to keep them visible while scrolling through other parts of the worksheet, and securing the worksheet by setting a password. Here’s how you can do it:

```cs
/**
Configure Worksheet Options
anchor-configure-worksheet-options
**/
// Add password protection to the worksheet
sheet.ProtectSheet("Password");
// Freeze the top row of the worksheet
sheet.CreateFreezePane(0, 1);
```

The initial row is locked in place and remains stationary while you scroll through the rest of the Worksheet. Additionally, the worksheet is secured with a password to prevent unauthorized modifications. This setup is demonstrated in Figures 7 and 8.

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/freeze-panes.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 7</strong> – <em>Freeze Panes</em></p>
	</div>
</center>

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 8</strong> – <em>Protected Worksheet</em></p>
	</div>
</center>

### 7.2. Configure Page Layout and Print Settings ###

It's possible to customize various page attributes including the page orientation, its dimensions, and the specified print area, among others.

Here is the paraphrased section with the relative URL paths resolved to ironsoftware.com:

```cs
/**
Configuring Page Layout and Print Settings
anchor-configure-page-layout-and-print-settings
**/
sheet.SetPrintArea("A1:L12");
sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

In this code snippet, we define the printable area of the worksheet to be from cells A1 to L12. We configure the printing orientation to landscape mode and set the paper size to a standard A4.

The print settings assign the area from A1 to L12 for printing. The document orientation is configured as Landscape, and the size of the paper is specified as A4.

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 9</strong> – <em>Print Setup</em></p>
	</div>
</center>

<hr class="separator">

## 8. Save the Excel Workbook

You can easily store your workbook by using the following code snippet:
```cs
/**
Save the Excel Workbook
anchor-save-workbook
**/
workbook.SaveAs("FinancialReport.xlsx");
```

The `SaveAs` method allows you to specify the file name for your workbook, in this case, saving it as "FinancialReport.xlsx".
```

```cs
// Persisting the Excel Workbook
// Reference Tag: anchor-save-workbook
workbook.SaveAs("Budget.xlsx");
```

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
      <a class="btn btn-white3" href="/csharp/excel/tutorials/downloads/Using.CSharp.to.Create.Excel.Files.in.Net.zip">
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
      <a class="doc-link" href="https://github.com/iron-software/tutorials/tree/master/IronXL/Using%20C%23%20to%20Create%20Excel%20Files%20in%20.Net" target="_blank">How to Create Excel File in C# on GitHub<i class="fa fa-chevron-right"></i></a>
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
      <h3>Read the XL API Reference</h3>
      <p>Explore the API Reference for IronXL, outlining the details of all of IronXL’s features, namespaces, classes, methods fields and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

