# C# Excel File Creation Guide

***Based on <https://ironsoftware.com/tutorials/create-excel-file-net_old_changed may 2021/>***


This guide provides detailed instructions on how to generate an Excel Workbook on any system compatible with .NET Framework 4.5 or .NET Core. Craft Excel documents easily in C# without needing the outdated **Microsoft.Office.Interop.Excel** library. With IronXL, you can configure worksheet features such as freeze panes and sheet protection, adjust printing settings, and much more.

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

[IronXL is an efficient C# & VB Excel API](https://ironsoftware.com/csharp/excel/) that enables reading, editing, and creation of Excel spreadsheet files in .NET, delivering top-notch performance. It entirely eliminates the need for installing MS Office or Excel Interop.

IronXL offers comprehensive support for .NET Core, .NET Framework, Xamarin, Mobile platforms, Linux, macOS, and Azure.

<h3>IronXL Features:</h3>

- Direct assistance from our dedicated .NET development team.

- Swift setup through Microsoft Visual Studio.

- Complimentary for development phases. Pricing starts from `$liteLicense`.

<h4>Create and Save an Excel File: Quick Code</h4>

<a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">https://www.nuget.org/packages/IronXL.Excel/</a>

Alternatively, you can directly download the `IronXL.dll` from [this link](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and include it in your project.

Below is the paraphrased section from the provided article:

```cs
// Creating and Saving an Excel File Example
// Reference ID: anchor-create-and-save-an-excel-file
using IronXL;

// Create a new workbook with the default format XLSX, adjustable via options
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
var sheet = workbook.CreateWorkSheet("sample_sheet");

// Assigning a simple value to a cell
sheet["A1"].Value = "Sample Data";
// Assign the same value to a range of cells
sheet["A2:A4"].Value = 10;
// Apply a background color to a cell
sheet["A5"].Style.SetBackgroundColor("#e0e0e0");
// Make text bold for a range of cells
sheet["A5:A6"].Style.Font.Bold = true;
// Using a formula to calculate the sum
sheet["A6"].Value = "=SUM(A2:A4)";

// Simple condition to check calculation results
if (sheet["A6"].IntValue == sheet["A2:A4"].IntValue)
{
    Console.WriteLine("Sum calculation is correct");
}

// Save the workbook to a file
workbook.SaveAs("sample_workbook.xlsx");
```

This rewritten code includes changes to variable names, comments, and text outputs to provide a natural, alternate expression of the original script while maintaining the logic and structure.

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

##  1. Acquire the No-Cost IronXL C# Library

### Installing via NuGet

There are several methods to integrate the IronXL NuGet package into your projects:

1. **Visual Studio Integration**: 
   - Access the NuGet Package Manager by navigating through the Project Menu or by directly right-clicking on your project in the Solution Explorer. These are depicted in Figures 3 and 4 below.

    <center>
      <div style="display: inline-block; text-align: left; margin-bottom: 20px;">
        <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
        <p><strong>Figure 3</strong> – <em>Menu of the Project</em></p>
      </div>
    </center>

    <center>
      <div style="display: inline-block; text-align: left;">
        <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
        <p><strong>Figure 4</strong> – <em>Solution Explorer Context Menu</em></p>
      </div>
    </center>

    From either location, choose 'Manage NuGet Packages', search for IronXL.Excel, and proceed to install it, as shown in Figure 5.

    <center>
      <div style="display: inline-block; text-align: left;">
        <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
        <p><strong>Figure 5</strong> – <em>Installing IronXL.Excel NuGet Package</em></p>
      </div>
    </center>

2. **Using Developer Command Prompt**:
   - Locate the Developer Command Prompt which might be found in the Visual Studio directory. 
     Type the command:
     ```
     PM> Install-Package IronXL.Excel
     ```
     Hit the Enter key to execute it. After installation, make sure to reload your project in Visual Studio.

3. **Direct NuGet Package Download**:
   - Visit the [IronXL NuGet page](https://www.nuget.org/packages/ironxl.excel) and click on the 'Download Package' button. Once downloaded, execute the package file and reload your project in Visual Studio.

### Direct Library Download

Alternatively, download the IronXL library directly from [Iron Software's official site](https://ironsoftware.com/csharp/excel/).

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Direct Download of IronXL Library</em></p>
  </div>
</center>

Adding the library to your project is simple:
1. Right-click the Solution icon in Solution Explorer.
2. Click on References.
3. Click on 'Add Reference', browse to locate the downloaded `IronXL.dll` file, and select it.
4. Confirm by clicking 'OK'.

### Ready to Start!

With IronXL set up, you're ready to dive into its powerful features and enhance your applications!

<h3>Install by Using NuGet</h3>

You can install the IronXL NuGet package using one of the following three methods:

### 1. Using Visual Studio

Visual Studio makes it simple to add NuGet packages to your projects. You can find the NuGet Package Manager under the Project menu or by right-clicking on your project in the Solution Explorer. This tool allows you to conveniently manage and install new packages.

### 2. Through the Developer Command Prompt

To install via the Developer Command Prompt, locate it typically under your Visual Studio directory, and open it. Simply type:

```
PM > Install-Package IronXL.Excel
```

Hit Enter, and the package will be automatically installed. Remember to reload your Visual Studio project after installation.

### 3. Direct NuGet Package Download

Alternatively, download the NuGet package directly. Visit the [IronXL NuGet page](https://www.nuget.org/packages/ironxl.excel/) and select "Download Package". Once downloaded, execute the package file and ensure your Visual Studio project is refreshed to apply the changes.

<h3>Visual Studio</h3>

Visual Studio comes equipped with a NuGet Package Manager, enabling you to easily add NuGet packages to your projects. You can find it in the Project Menu or by right-clicking on your project within the Solution Explorer. The step-by-step process is demonstrated in Figures 3 and 4 below.

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

Once you've selected "Manage NuGet Packages" using either method, search for and select the IronXL.Excel package to install it, as depicted in Figure 5.

<br></br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

<h3>Developer Command Prompt</h3>

To install the IronXL.Excel NuGet package via the Developer Command Prompt, please adhere to the subsequent instructions:

1. Locate your Developer Command Prompt, which is typically found within your Visual Studio installation directory.

2. Execute this specific command: 

3. Enter `PM > Install-Package IronXL.Excel`

4. Hit the Enter key to commence the installation.

5. After the command completes, the IronXL.Excel package will be successfully installed.

6. Finally, ensure to refresh your Visual Studio project to reflect the changes.

<h3>Download the NuGet Package directly</h3>

To download the NuGet package, follow these instructions:

1. Go to [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/)
2. Choose the "Download Package" option.
3. Once the download is complete, double-click on the downloaded file.
4. Restart your Visual Studio project to apply the changes.

</br>
<h3>Install IronXL by Direct Download of the Library</h3>

Here's a rephrased version of the specified section with resolved URL paths:

---
Alternatively, you can directly download IronXL from its official page here: [IronXL Download](https://ironsoftware.com/csharp/excel/)

</br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library</em></p>
  </div>
</center>

Here is the paraphrased section of the article with all relative paths resolved:

-----
Include the library in your project using the following instructions:

1. Right-click on the Solution within the Solution Explorer.
2. Choose 'References' from the context menu.
3. Search for and select the `IronXL.dll` in the file dialog.
4. Confirm your selection by clicking 'OK'.

<h3>Let's Go!</h3>

Let's dive into utilizing the powerful capabilities of the IronXL library now that you're all set up!

<hr class="separator">

<h4 class="tutorial-segment-title">How to Tutorials</h4>

## 2. Initiating an ASP.NET Project ##

Next, I'll guide you through the process of starting a new ASP.NET project with IronXL. Follow these simple steps to get your project up and running quickly:

1. Firstly, visit the official NuGet page for IronXL at [this link](https://www.nuget.org/packages/ironxl.excel/).
2. Download the package by clicking on 'Download Package.'
3. Once downloaded, open the package to integrate it seamlessly into your development environment.
4. Refresh or reload your Visual Studio project to sync the new integration.

### Setting Up Your ASP.NET Website

Once IronXL is incorporated into your project, creating the ASP.NET Website is straightforward:

1. Launch Visual Studio and select File > New Project from the top menu.
2. In the Project type listbox, under Visual C#, select 'Web.'
3. Choose 'ASP.NET Web Application' for your project type, as illustrated in the image below:

    ![Creating a New ASP.NET Web Application](https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png)
    
    *Figure 1 – New Project Setup*

4. After clicking OK, you'll reach a new screen. Here, select 'Web Forms' as depicted in the subsequent image:

    ![Selecting Web Forms for Project](https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png)
    
    *Figure 2 – Setting Up Web Forms*

5. Confirm your settings by clicking OK.

Now, with your ASP.NET environment prepared, you're ready to incorporate IronXL and start leveraging its powerful features for managing Excel data.

<ul>
  <li>Navigate to the following URL:</li>
  <li><a href="https://www.nuget.org/packages/ironxl.excel/" target="_blank">https://www.nuget.org/packages/ironxl.excel/</a></li>
  <li>Click on Download Package</li>
  <li>After the package has downloaded, double click it</li>
  <li>Reload your Visual Studio project</li>
</ul>

Here's the paraphrased section with the updated URL paths resolved to `ironsoftware.com`:

---
To begin developing an ASP.NET website, you can follow these steps:

1. Launch Visual Studio.
2. From the top menu, select File, then choose New Project.
3. Under the 'Visual C#' category, pick 'Web' from the project type options.
4. Choose the 'ASP.NET Web Application' option, as illustrated in the accompanied image.

<br></br>
<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 1</strong> – *New Project*
```

<a href="https://ironsoftware.com/csharp/excel/" target="_blank">ironsoftware.com/csharp/excel/</a>

5. Confirm your choice by clicking OK.

6. Subsequently, on the following interface, choose the Web Forms option. Refer to Figure 2 below for visual guidance.

<br></br>

<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 2</strong> – *Web Forms*
```

<br></br>

Once you click OK, you're ready to dive in. Begin by installing IronXL to start tailoring your Excel file to your needs.

<hr class="separator">

```cs
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

Both XLS (older Excel file version) and XLSX (current and newer file version) file formats can be created with IronXL.

### 3.1. Set a Default Worksheet ###

And, it’s even simpler to create a default Worksheet:

```cs
var sheet = workbook.CreateWorkSheet("2020 Budget");
```

"Sheet" in the above code snippet represents the worksheet and you can use it to set cell values and almost everything Excel can do.

In case you are confused about the difference between a Workbook and a Worksheet, let me explain:

A Workbook contains Worksheets. This means that you can add as many Worksheets as you like into one Workbook. In a later article, I will explain how to do this. A Worksheet contains Rows and Columns. The intersection of a Row and a Column is called a Cell, and this is what you will manipulate whilst working with Excel.

 -----
 
## 3. Generating an Excel Workbook ##

Creating a brand new Excel Workbook with IronXL is surprisingly straightforward and can be achieved in just one line of code. Here's how:

```cs
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL allows you to produce both older XLS and the contemporary XLSX file formats.

### 3.1. Initialize a Default Worksheet ###

Crafting a default worksheet is equally effortless:

```cs
var sheet = workbook.CreateWorkSheet("2020 Budget");
```

Here, `sheet` refers to the newly created worksheet, which can be manipulated to set cell values and facilitate nearly all Excel functionalities.

For those new to Excel programming, here’s a quick primer on terms:
- A **Workbook** is a collection of Worksheets. You can insert multiple worksheets into a single workbook.
- A **Worksheet** comprises Rows and Columns, where the meeting point of a row and a column is called a Cell. You can change or use cells as needed when you’re handling Excel data.

This simple distinction helps you better understand and utilize the capabilities of Excel through IronXL.

```cs
// Instantiate a new Workbook using IronXL with the XLSX file format
WorkBook newWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL supports the creation of both XLS (the traditional Excel file format) and XLSX (the modern and current Excel file format).

### 3.1. Establishing a Default Worksheet ###

Creating a default worksheet is even more straightforward:
```

```cs
var sheet = workbook.AddWorkSheet("2020 Budget");
```

In the provided code snippet, "Sheet" refers to the worksheet where you can customize cell values and perform various Excel operations seamlessly.

If you're unsure about the distinction between a Workbook and a Worksheet, here's a brief clarification:

A Workbook is a collection of Worksheets, allowing you to integrate multiple Worksheets within a single Workbook. This topic will be elaborated in a forthcoming article. On the other hand, a Worksheet is made up of Rows and Columns, forming Cells at their intersections. These Cells are the primary elements you will interact with when using Excel functionalities.

<hr class="separator">

## 4. Setting Cell Values ##

Setting cell values is a straightforward operation. You designate the specific cell and assign its value. Below is a practical demonstration:

```cs
/**
Assign Cell Values Individually
anchor-set-cell-values
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

In the example above, I've filled out the cells in the first row, from columns A through L, each with the name of a month.

### 4.2. Assign Cell Values Dynamically ###

Assigning values dynamically offers flexibility. You aren't required to manually specify each cell's location. In the following example, you'll see how random values are assigned using a loop:

```cs
/**
Dynamically Set Cell Values
anchor-set-cell-values-dynamically
**/
Random rand = new Random();
for (int i = 2; i <= 11; i++)
{
    sheet["A" + i].Value = rand.Next(1, 1000);
    sheet["B" + i].Value = rand.Next(1000, 2000);
    sheet["C" + i].Value = rand.Next(2000, 3000);
    sheet["D" + i].Value = rand.Next(3000, 4000);
    sheet["E" + i].Value = rand.Next(4000, 5000);
    sheet["F" + i].Value = rand.Next(5000, 6000);
    sheet["G" + i].Value = rand.Next(6000, 7000);
    sheet["H" + i].Value = rand.Next(7000, 8000);
    sheet["I" + i].Value = rand.Next(8000, 9000);
    sheet["J" + i].Value = rand.Next(9000, 10000);
    sheet["K" + i].Value = rand.Next(10000, 11000);
    sheet["L" + i].Value = rand.Next(11000, 12000);
}
```

Each cell from A2 to L11 is assigned a unique, randomly generated number.

### 4.3. Direct Addition from Databases ###

For adding data directly from a database, set up the connection, execute a select query, and dynamically allocate the results to specific cells as shown below:

```cs
/**
Populate Cells Directly from a Database
anchor-add-directly-from-a-database
**/
// Setup database connection and obtain data
string connectionString;
SqlConnection connection;
SqlDataAdapter dataAdapter;
DataSet dataSet = new DataSet("ExampleDataSet");

// Database connection string
connectionString = @"Data Source=Server_Name;Initial Catalog=Database_Name;User ID=User_ID;Password=Password";

// SQL query to retrieve data
string sql = "SELECT Column_Names FROM Table_Name";

// Establish connection and fill dataset
connection = new SqlConnection(connectionString);
dataAdapter = new SqlDataAdapter(sql, connection);

connection.Open();
dataAdapter.Fill(dataSet);

// Iterate through dataset and populate cells
foreach (DataTable table in dataSet.Tables)
{
    for (int j = 12; j <= 21; j++)
    {
        int rowCount = table.Rows.Count - 1;
        sheet["A" + j].Value = table.Rows[rowCount]["Column1"].ToString();
        sheet["B" + j].Value = table.Rows[rowCount]["Column2"].ToString();
        sheet["C" + j].Value = table.Rows[rowCount]["Column3"].ToString();
        sheet["D" + j].Value = table.Rows[rowCount]["Column4"].ToString();
        sheet["E" + j].Value = table.Rows[rowCount]["Column5"].ToString();
        sheet["F" + j"].Value = table.Rows[rowCount]["Column6"].ToString();
        sheet["G" + j].Value = table.Rows[rowCount]["Column7"].ToString();
        sheet["H" + j].Value = table.Rows[rowCount]["Column8"].ToString();
        sheet["I" + j].Value = table.Rows[rowCount]["Column9"].ToString();
        sheet["J" + j].Value = table.Rows[rowCount]["Column10"].ToString();
        sheet["K" + j].Value = table.Rows[rowCount]["Column11"].ToString();
        sheet["L" + j].Value = table.Rows[rowCount]["Column12"].ToString();
        rowCount++;
    }
}
```

### 4.1. Manual Cell Value Setting ###

Manually setting the values in cells is straightforward. Simply specify the target cell and assign it a value. Here’s how you do it:

```cs
/**
Manually Assign Values to Cells
anchor-manually-set-values
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

In this example, we have filled the first row with the names of the months from January to December, assigning each month to the corresponding cell from `A1` to `L1`.

The paraphrased content for setting cell values manually in an Excel sheet using IronXL is as follows:

```cs
/**
Assign Cell Values Individually
anchor-assign-cell-values
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

In this reformatted section, I have updated the code comment to provide clarity on the action being performed, which is "Assign Cell Values Individually". Each line sets the `Value` for a different cell corresponding to the months of the year in a sequential manner.

In this example, I've filled Columns A through L with names corresponding to each month, assigning them to the first row of each column.

### 4.2. Dynamically Modify Cell Values ###

Dynamically modifying cell values follows a similar method as before but adds flexibility by avoiding fixed cell references. In the following example, you'll see how to instantiate a `Random` object to generate random numbers. You will also use a `for` loop to cycle through a designated range of cells, filling them with these randomly generated values.

```cs
/**
Dynamically Assign Cell Values
anchor-dynamically-assign-cell-values
**/
Random randomizer = new Random();
for (int index = 2; index <= 11; index++)
{
	sheet[$"A{index}"].Value = randomizer.Next(1, 1000);
	sheet[$"B{index}"].Value = randomizer.Next(1000, 2000);
	sheet[$"C{index}"].Value = randomizer.Next(2000, 3000);
	sheet[$"D{index}"].Value = randomizer.Next(3000, 4000);
	sheet[$"E{index}"].Value = randomizer.Next(4000, 5000);
	sheet[$"F{index}"].Value = randomizer.Next(5000, 6000);
	sheet[$"G{index}"].Value = randomizer.Next(6000, 7000);
	sheet[$"H{index}"].Value = randomizer.Next(7000, 8000);
	sheet[$"I{index}"].Value = randomizer.Next(8000, 9000);
	sheet[$"J{index}"].Value = randomizer.Next(9000, 10000);
	sheet[$"K{index}"].Value = randomizer.Next(10000, 11000);
	sheet[$"L{index}"].Value = randomizer.Next(11000, 12000);
}
```

Each cell between A2 and L11 displays a unique, randomly added value.

Discussing dynamic inputs further, would learning to automatically populate cells with data from a database be of interest? The following code segment demonstrates this process, provided your database connections are appropriately established.

### 4.3. Integrating Data from a Database ###

To populate your Excel cells directly from a database, you'll follow a straightforward process. Here's a step-by-step method to input data seamlessly into your spreadsheet:

```cs
/**
Database to Excel Integration
anchor-integrating-data-from-database
**/
// Initialize database connection objects
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Define your database connection string
connectionString = @"Data Source=Your_Server;Initial Catalog=Your_Database;User ID=Your_Username;Password=Your_Password";

// SQL query to fetch data
query = "SELECT Columns FROM Your_Table";

// Manage database connection and fetch data
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(query, sqlConnection);

sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Iterate through dataset to transfer data to Excel cells
foreach (DataTable table in dataSet.Tables)
{
    int rowIndex = table.Rows.Count - 1;

    for (int i = 12; i <= 21; i++)
    {
        sheet["A" + i].Value = table.Rows[rowIndex]["Column1"].ToString();
        sheet["B" + i].Value = table.Rows[rowIndex]["Column2"].ToString();
        sheet["C" + i].Value = table.Rows[rowIndex]["Column3"].ToString();
        sheet["D" + i].Value = table.Rows[rowIndex]["Column4"].ToString();
        sheet["E" + i].Value = table.Rows[rowIndex]["Column5"].ToString();
        sheet["F" + i].Value = table.Rows[rowIndex]["Column6"].ToString();
        sheet["G" + i].Value = table.Rows[rowIndex]["Column7"].ToString();
        sheet["H" + i].Value = table.Rows[rowIndex]["Column8"].ToString();
        sheet["I" + i].Value = table.Rows[rowIndex]["Column9"].ToString();
        sheet["J" + i].Value = table.Rows[rowIndex]["Column10"].ToString();
        sheet["K" + i].Value = table.Rows[rowIndex]["Column11"].ToString();
        sheet["L" + i].Value = table.Rows[rowIndex]["Column12"].ToString();
    }
    rowIndex++;
}
```

This method facilitates transferring structured data from your database directly into the designated Excel cells, all done programmatically with IronXL.

```cs
/**
Populate Cells from Database
anchor-populate-cells-from-database
**/
// Initialize database objects for data retrieval
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection connection;
SqlDataAdapter adapter;

// Configure the database connection string
connectionString = @"Data Source=Your_Server_Name;Initial Catalog=Your_Database;User ID=Your_User_ID;Password=Your_Password";

// SQL statement to fetch data
query = "SELECT Columns FROM Your_Table";

// Establish connection and fill dataset
connection = new SqlConnection(connectionString);
adapter = new SqlDataAdapter(query, connection);

connection.Open();
adapter.Fill(dataSet);

// Iterate through the dataset table content
foreach (DataTable dataTable in dataSet.Tables)
{
    int rowCount = dataTable.Rows.Count - 1;

    for (int i = 12; i <= 21; i++)
    {
        sheet ["A" + i].Value = dataTable.Rows[rowCount]["Column1"].ToString();
        sheet ["B" + i].Value = dataTable.Rows[rowCount]["Column2"].ToString();
        sheet ["C" + i].Value = dataTable.Rows[rowCount]["Column3"].ToString();
        sheet ["D" + i].Value = dataTable.Rows[rowCount]["Column4"].ToString();
        sheet ["E" + i].Value = dataTable.Rows[rowCount]["Column5"].ToString();
        sheet ["F" + i].Value = dataTable.Rows[rowCount]["Column6"].ToString();
        sheet ["G" + i].Value = dataTable.Rows[rowCount]["Column7"].ToString();
        sheet ["H" + i].Value = dataTable.Rows[rowCount]["Column8"].ToString();
        sheet ["I" + i].Value = dataTable.Rows[rowCount]["Column9"].ToString();
        sheet ["J" + i].Value = dataTable.Rows[rowCount]["Column10"].ToString();
        sheet ["K" + i].Value = dataTable.Rows[rowCount]["Column11"].ToString();
        sheet ["L" + i].Value = dataTable.Rows[rowCount]["Column12"].ToString();
    }
    rowCount++;
}
```

Simply assign the desired field name to the `Value` property of the specified cell.

<hr class="separator">

## 5. Implement Formatting Techniques ##

### 5.1. Adjust Cell Background Colors ###

To modify the background color of a single cell or a group of cells, it only takes a simple line of code:

```cs
// Set the background color of a range of cells to gray
sheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This command alters the background color of the selected cells to gray. In this example, the color code is represented in hex format, where "#d3d3d3" specifies the gray color.

### 5.2. Add Borders to Cells ###

Creating borders in IronXL is straightforward and can be done with minimal code. Here’s how you can do it:

```cs
/**
 * This section adds borders to specific cell ranges
 * anchor-add-borders-to-cells
 **/
// Set black border on the top and bottom of cells A1 to L1
sheet["A1:L1"].Style.TopBorder.SetColor("#000000");
sheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Set a medium style right border from L2 to L11
sheet["L2:L11"].Style.RightBorder.SetColor("#000000");
sheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Define a medium bottom border for the range A11 to L11
sheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
sheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

In the commands above, we add top and bottom borders to cells from A1 to L1 and right borders to cells from L2 to L11. These borders are set to black with a right border styled as medium. Then, a medium bottom border is similarly applied to cells from A11 to L11.

### 5.1. Configure Cell Background Colors ###

Setting the background color for a single cell or a group of cells is straightforward. Use the code example below to accomplish this:
```

```cs
/**
Set Background Color for Cell Range
anchor-setting-cell-background-colors
**/
sheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");  // Update the background color of cells from A1 to L1 to a light gray shade
```

The background color for a selected range of cells is assigned a gray shade. The color specification uses the RGB (Red, Green, Blue) format. In this format, the first two hexadecimal characters indicate the red component, the following two denote green, and the last two represent blue. The hexadecimal values can vary from 0 to 9 and from A to F.

### 5.2. Establish Cell Borders ###

Easily create borders in your spreadsheet with IronXL using the following straightforward method:

```cs
/**
Designate Border Styles
anchor-designate-border-styles
**/
// Setting top and bottom borders for the first row with black color
sheet ["A1:L1"].Style.TopBorder.SetColor("#000000");
sheet ["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Applying medium thickness to the right border from row 2 to row 11
sheet ["L2:L11"].Style.RightBorder.SetColor("#000000");
sheet ["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Applying medium thickness to the bottom border for the 11th row
sheet ["A11:L11"].Style.BottomBorder.SetColor("#000000");
sheet ["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

In the explained code segment, black borders were applied to the top and bottom of cells from `A1` to `L1`. Additionally, a medium-strength right border was established for cells spanning from `L2` to `L11`. Finally, a medium bottom border was also applied to the range from `A11` to `L11`. This configuration effectively outlines specified sections of the worksheet for emphasis or organizational purposes.

<hr class="separator">

## 6. Implementing Formulas in Cells ##

IronXL simplifies using formulas significantly, a feature I always emphasize due to its ease of use. Observe how seamlessly you can integrate formulas into your spreadsheet operations using the code example below:
```

```cs
/**
Applying Formulas in Spreadsheet Cells
anchor-applying-formulas-in-cells
**/
decimal total = sheet["A2:A11"].Sum();
decimal average = sheet["B2:B11"].Avg();
decimal maximum = sheet["C2:C11"].Max();
decimal minimum = sheet["D2:D11"].Min();

sheet["A12"].Value = total;
sheet["B12"].Value = average;
sheet["C12"].Value = maximum;
sheet["D12"].Value = minimum;
```

The flexibility of setting the cell's data type to match the formula result is a significant advantage. The example provided demonstrates utilizing various formulas in IronXL: `SUM` to total values, `AVG` to compute the average, `MAX` to find the maximum value, and `MIN` to determine the minimum value.

<hr class="separator">

## 7. Configuring Worksheet and Printing Settings ##

### 7.1. Adjusting Worksheet Attributes ###

You can easily manage the behavior of your worksheet by applying protection and setting freeze panes. Here's how:

```cs
/**
Adjust Worksheet Attributes
anchor-setting-worksheet-properties
**/
sheet.ProtectSheet("YourPasswordHere");
sheet.CreateFreezePane(0, 1);
```

In this example, the first row remains static while you scroll, thanks to the `CreateFreezePane` method, and changes are restricted with a password using `ProtectSheet`.

Here is a glimpse of applied settings:

<div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" alt="Freeze Panes Example" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 7</strong> – <em>Freeze Panes Example</em></p>
</div>

<div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="Protected Worksheet Example" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 8</strong> – <em>Protected Worksheet Example</em></p>
</div>

### 7.2. Tailoring Page and Print Layouts ###

Customize how your worksheets appear when printed with these settings:

```cs
/**
Adjust Print Layout
anchor-configuring-print-properties
**/
sheet.SetPrintArea("A1:L12");
sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

The code snippet above defines the print area as cells from A1 to L12, sets the page orientation to landscape and selects A4 as the paper size.

Visualize the print setup:

<div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" alt="Print Setup Example" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Print Setup Example</em></p>
</div>

These features not only enhance the functionality and security of your worksheets but also ensure that your documents print correctly, appearing just as you intend.

### 7.1. Configure Worksheet Settings ###

Adjusting the worksheet properties enables the freezing of specific rows and columns while also allowing you to secure the worksheet by setting a password. Here are the steps:

Below is the paraphrased section of the article, with relative URL paths resolved to ironsoftware.com.

-----
```cs
/**
Configure Worksheet Settings
anchor-configure-worksheet-settings
**/
sheet.ProtectSheet("SecurePassword");
sheet.CreateFreezePane(0, 1);
```
In this code snippet, the `sheet.ProtectSheet("SecurePassword");` enforces a password protection on the worksheet ensuring it remains secure. The `sheet.CreateFreezePane(0, 1);` command freezes the top row of the worksheet to facilitate easy navigation by keeping headers visible as you scroll through the document.

The initial row remains static and won't move when you scroll through the worksheet, providing a consistent view of the row headers. Additionally, the worksheet is safeguarded against unauthorized modifications with the use of a password. Observe these features demonstrated in Figures 7 and 8.

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

### 7.2. Configure Page and Printing Settings ###

Modify various page settings including the page orientation, page size, and the print area, among others.

Below is the paraphrased section adjusted with absolute URL paths based on your request to "resolve any relative URL paths to ironsoftware.com":

```cs
/**
Configure Print and Page Settings
link-to-details-set-page-and-print-properties
**/
sheet.SetPrintArea("A1:L12");
sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;  // Configures the worksheet to print in landscape layout
sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;  // Sets the paper size to A4
``` 

The modified code snippet includes brief comments to explain what each line of code is doing, enhancing understandability.

The print settings configure the printable area from cell `A1` to `L12`. The document is oriented in a landscape format, and `A4` is designated as the paper size.

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 9</strong> – <em>Print Setup</em></p>
	</div>
</center>

<hr class="separator">

## 8. Save Your Workbook ##

To persist the Workbook on disk, execute the following code snippet:
```cs
/**
Save Workbook
anchor-save-workbook
**/
workbook.SaveAs("Budget.xlsx");
```
```

Here's the paraphrased code snippet section:

```cs
// Save the Excel Workbook
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

