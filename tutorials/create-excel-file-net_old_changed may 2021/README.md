# C# Create Excel File Tutorial

This guide provides a detailed, step-by-step approach to creating an Excel Workbook file across all platforms supporting .NET Framework 4.5 or .NET Core. Simplify your C# Excel file creation without relying on the older **Microsoft.Office.Interop.Excel** library by leveraging IronXL. This allows you to configure worksheet options such as freeze panes and protection, manage print settings, and much more.

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

[IronXL offers a streamlined C# and VB Excel API](https://ironsoftware.com/csharp/excel/) that facilitates the reading, editing, and creation of Excel spreadsheets within .NET, ensuring quick performance. It operates independently, with no requirement for MS Office or Excel Interop installations.

Furthermore, IronXL provides comprehensive support across multiple platforms and technologies including .NET Core, .NET Framework, Xamarin, Mobile devices, Linux, macOS, and Azure environments.

<h3>IronXL Features:</h3>

- Personalized human support directly from our .NET development experts

- Quick and easy setup using Microsoft Visual Studio

- Complimentary for development purposes. Pricing for licenses starts at `$liteLicense`.

<h4>Create and Save an Excel File: Quick Code</h4>

<a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">https://www.nuget.org/packages/IronXL.Excel/</a>

Alternatively, the `IronXL.Dll` library can be obtained and incorporated into your project via the following [download link](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).

Here's the paraphrased section of the article:

```cs
/**
Initialization & Persisting Excel File
anchor-initialization-and-persisting-excel-file
**/
using IronXL;

// By default, the workbook will use the XLSX format, but you can specify a different format with the CreatingOptions if needed
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
var sheet = workbook.CreateWorkSheet("sample_sheet");

// Assigning a simple value to a cell
sheet["A1"].Value = "Example";

// Assign values to a range of cells simultaneously
sheet["A2:A4"].Value = 5;

// Apply background color to a cell
sheet["A5"].Style.SetBackgroundColor("#f0f0f0");

// Apply bold styling to text in a range of cells
sheet["A5:A6"].Style.Font.Bold = true;

// Use a formula to compute values
sheet["A6"].Value = "=SUM(A2:A4)";

// A simple assertion to check if the formula computes as expected
if (sheet["A6"].IntValue == sheet["A2:A4"].IntValue)
{
    Console.WriteLine("Validation succeeded");
}

// Save the workbook to a file
workbook.SaveAs("sample_workbook.xlsx");
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## Install IronXL C# Library for Free

Getting started with the IronXL C# library is straightforward and offers multiple installation methods. Here’s how you can incorporate this powerful tool into your projects.

### Installing via Visual Studio

IronXL can be seamlessly integrated with your .NET projects using Visual Studio. To begin, you can add the package directly through the NuGet Package Manager:

1. Open the NuGet Package Manager by either navigating through the **Project Menu** or by right-clicking your project in the **Solution Explorer**.
   ![Project Menu](https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png)
   *Figure 3 – The Project Menu*

   ![Right Click Solution Explorer](https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png)
   *Figure 4 – Right Click in Solution Explorer*

2. After accessing the NuGet Package Manager, search for `IronXL.Excel` and proceed to install it as demonstrated below:
   ![Install IronXL.Excel NuGet Package](https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png)
   *Figure 5 – Installing the IronXL.Excel NuGet Package*

### Using the Developer Command Prompt

For those who prefer using the command line, the Developer Command Prompt is an excellent alternative:

1. Find the Developer Command Prompt in your Visual Studio installation directory.
2. Execute the following command:
   ```
   PM> Install-Package IronXL.Excel
   ```
3. Hit Enter to install the package and then reload your Visual Studio project.

### Direct Download

Alternatively, you can directly download the NuGet package:

1. Visit [IronXL on NuGet](https://www.nuget.org/packages/ironxl.excel/) and click the **Download Package**.
2. Once downloaded, open the package and refresh your Visual Studio project setup.

### Direct Library Download

If you prefer downloading the library files directly, IronXL is also available for direct download:

1. Go to the IronXL download page: [Download IronXL Library](https://ironsoftware.com/csharp/excel/)
   ![Download IronXL Library](https://ironsoftware.com/img/tutorials/create-excel-file-net/download-ironxl-library.png)
   *Figure 6 – Downloading IronXL directly*

2. Once downloaded, add the `IronXL.dll` to your project’s references by right-clicking the solution in the Solution Explorer, selecting **References**, then **Browse**, and navigating to the downloaded library.

With IronXL installed, you’re equipped to leverage its full capabilities in your C# projects to manipulate Excel data with ease. Let’s get started!

<h3>Install by Using NuGet</h3>

Here are three distinct methods to integrate the IronXL NuGet package into your development environment:

1. **Visual Studio Integration**  
   Utilize the NuGet Package Manager available within Visual Studio for straightforward installation. You can find this option under the Project Menu or by right-clicking on your project in the Solution Explorer.

2. **Using Developer Command Prompt**  
   Launch the Developer Command Prompt, which is typically found in your Visual Studio installation directory. Enter the following command to install the package:
   ```
   PM> Install-Package IronXL.Excel
   ```
   Simply hit Enter, and the package will automatically be incorporated into your project.

3. **Direct NuGet Package Download**  
   If you prefer direct downloads, you can obtain the NuGet package manually. Navigate to [https://www.nuget.org/packages/IronXL.Excel](https://www.nuget.org/packages/IronXL.Excel), click 'Download Package', and once downloaded, execute the file to bring it into your project.

<h3>Visual Studio</h3>

Visual Studio includes the NuGet Package Manager, enabling you to add NuGet packages to your projects conveniently. Access it through the Project Menu or by right-clicking your project in the Solution Explorer. These methods are illustrated in Figures 3 and 4 below.

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

Once you've accessed the Manage NuGet Packages via either method described, search for and select the `IronXL.Excel` package. Proceed to install it as depicted in Figure 5.

<br></br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

<h3>Developer Command Prompt</h3>

Access the Developer Command Prompt and proceed with the following instructions to integrate the IronXL.Excel NuGet package into your project:

1. Locate the Developer Command Prompt, typically found within your Visual Studio installation directory.

2. Enter this command:
   
   ```
   PM > Install-Package IronXL.Excel
   ```

3. Hit the Enter key to execute the command.

4. The package will automatically be downloaded and installed.

5. Once the installation is complete, refresh your Visual Studio project to apply the changes.

<h3>Download the NuGet Package directly</h3>

Follow these simple steps to download the NuGet package:

1. Go to the URL: [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/)
2. Select the 'Download Package' option.
3. Once the download is complete, double-click on the downloaded file.
4. Restart your Visual Studio project to apply changes.

</br>
<h3>Install IronXL by Direct Download of the Library</h3>

The alternative method for installing IronXL involves directly downloading it from the website at [ironsoftware.com/csharp/excel](https://ironsoftware.com/csharp/excel/).

</br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library</em></p>
  </div>
</center>

Here's the paraphrased section with the relative URL paths resolved:

-----
To add the library to your project, simply follow these steps:

1. In Solution Explorer, perform a right-click on the Solution.
2. Click on 'References' from the context menu.
3. Navigate through your files to locate the IronXL.dll library.
4. Confirm by clicking 'OK'.

<h3>Let's Go!</h3>

Now that your setup is complete, let’s dive into exploring the powerful capabilities of the IronXL library!

<hr class="separator">

<h4 class="tutorial-segment-title">How to Tutorials</h4>

## 2. Start an ASP.NET Project ##

Initiate your ASP.NET project by following these straightforward steps:

1. Open Visual Studio on your computer.
2. From the top menu, choose File > New Project.
3. In the Project type list, select Web under the Visual C# section.
4. Then, pick ASP.NET Web Application as displayed below.
   
    <center>
      <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank">
        <p><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p>
      </a>
    </center>
    <strong style="margin-left: 40px;">Figure 1</strong> – *New Project Setup*

5. Confirm your selection by clicking OK.
6. On the subsequent screen, select Web Forms as explained in the next step:

    <center>
      <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
    </center>
    <strong style="margin-left: 40px;">Figure 2</strong> – *Selecting Web Forms Option*
    
7. Click OK to finalize the setup.

You are now ready to integrate IronXL and explore its functionalities as you develop your project.

<ul>
  <li>Navigate to the following URL:</li>
  <li><a href="https://www.nuget.org/packages/ironxl.excel/" target="_blank">https://www.nuget.org/packages/ironxl.excel/</a></li>
  <li>Click on Download Package</li>
  <li>After the package has downloaded, double click it</li>
  <li>Reload your Visual Studio project</li>
</ul>

Follow these steps to create an ASP.NET Website:

1. Start by launching Visual Studio.

2. Navigate to `File` > `New Project`.

3. In the Project type listbox, choose `Web` located under `Visual C#`.

4. Opt for `ASP.NET Web Application`, depicted in the following image.

<br></br>
<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

**Figure 1** – _New Project_
```

<a href="https://ironsoftware.com/csharp/excel/" target="_blank">ironsoftware.com/csharp/excel/</a>

5. Once you have made your selection, click "OK."

6. On the following interface, choose "Web Forms" as depicted in Figure 2 below.

<br></br>

<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 2</strong> – *Web Forms*
```

<br></br>

Once you click OK, you'll have the basic structure required. The next step is to install IronXL, which will enable you to begin customizing your Excel file to meet your specific needs.

<hr class="separator">

```cs
// Create a new Excel Workbook
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

Both the XLS (an earlier version of Excel file format) and the newer XLSX file formats are supported by IronXL.

### 3.1. Set a Default Worksheet ###

Creating a default worksheet is just as straightforward:

```cs
// Create a default worksheet named '2020 Budget'
var sheet = workbook.CreateWorkSheet("2020 Budget");
```

Here, "sheet" refers to the newly created worksheet, which you can now use to manage cell values and perform nearly any Excel-related operation.

If you're unsure about the distinction between a Workbook and a Worksheet, here's a quick clarification:

A Workbook is a collection of one or more Worksheets. You can add multiple Worksheets to a single Workbook. A Worksheet is composed of Rows and Columns, where the intersection of a Row and a Column forms a Cell. These cells are what you will manipulate when working with IronXL to manage Excel data.

The code snippet is rephrased as follows:

```cs
// Initialize a new WorkBook for the XLSX format
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL supports both the traditional XLS format and the modern XLSX file format for Excel documents.

### 3.1. Establish a Default Worksheet ###

Creating a default Worksheet is even more straightforward:

```cs
var sheet = workbook.AddSheet("2020 Budget");
```

In the provided code example, "Sheet" refers to a worksheet where you can modify cell values and perform almost all the functions that Excel offers.

If you're wondering about the distinction between a Workbook and a Worksheet, here's a quick primer:

A Workbook is essentially a container for Worksheets, allowing you to incorporate multiple Worksheets within a single Workbook. Details on how to add more Worksheets will be covered in a forthcoming article. Each Worksheet is composed of Rows and Columns, with a Cell located at the intersection of a Row and a Column. These Cells are the primary elements you interact with when handling data in Excel.

<hr class="separator">

## Modify Cell Values ##

### Manually Alter Cell Data ###

To manually assign values to specific cells in an Excel worksheet, you can directly specify the cell and its new content. Below is an illustration:

```cs
// Set individual cell values
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
The example fills the first row, columns A to L, setting each to a different month name.

### Dynamic Cell Value Assignment ###

For dynamic cell value assignment, you can create varied data entries without hard-coding each cell's value. The code below demonstrates this concept using a loop:

```cs
// Dynamically assigning random values to cells
Random randomGenerator = new Random();
for (int rowIndex = 2; rowIndex <= 11; rowIndex++)
{
    sheet["A" + rowIndex].Value = randomGenerator.Next(1, 1000);
    sheet["B" + rowIndex].Value = randomGenerator.Next(1000, 2000);
    sheet["C" + rowIndex].Value = randomGenerator.Next(2000, 3000);
    sheet["D" + rowIndex].Value = randomGenerator.Next(3000, 4000);
    sheet["E" + rowIndex].Value = randomGenerator.Next(4000, 5000);
    sheet["F" + rowIndex].Value = randomGenerator.Next(5000, 6000);
    sheet["G" + rowIndex].Value = randomGenerator.Next(6000, 7000);
    sheet["H" + rowIndex].Value = randomGenerator.Next(7000, 8000);
    sheet["I" + rowIndex].Value = randomGenerator.Next(8000, 9000);
    sheet["J" + rowIndex].Value = randomGenerator.Next(9000, 10000);
    sheet["K" + rowIndex].Value = randomGenerator.Next(10000, 11000);
    sheet["L" + rowIndex].Value = randomGenerator.Next(11000, 12000);
}
```
This fills cells from A2 to L11 with unique, randomly generated numbers.

### Loading Data from a Database ###

To insert data from a database directly, the following snippet outlines the procedure assuming database connection setup is correct:

```cs
// Populate cells from a database
string connectionString;
string sqlQuery;
DataSet dataSet = new DataSet("DataSetName");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Setting Database Connection
connectionString = @"Data Source=Server_Name;Initial Catalog=Database_Name;User ID=User_ID;Password=Password";

// SQL command for data retrieval
sqlQuery = "SELECT Fields FROM Table_Name";

// Initializing connection
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(sqlQuery, sqlConnection);

// Opening connection and filling the dataset
sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Loop through dataset to fill cells
foreach (DataTable table in dataSet.Tables)
{
    int rowCounter = table.Rows.Count - 1;
    for (int cellCounter = 12; cellCounter <= 21; cellCounter++)
    {
       sheet["A" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_1"].ToString();
       sheet["B" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_2"].ToString();
       sheet["C" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_3"].ToString();
       sheet["D" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_4"].ToString();
       sheet["E" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_5"].ToString();
       sheet["F" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_6"].ToString();
       sheet["G" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_7"].ToString();
       sheet["H" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_8"].ToString();
       sheet["I" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_9"].ToString();
       sheet["J" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_10"].ToString();
       sheet["K" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_11"].ToString();
       sheet["L" + cellCounter].Value = table.Rows[rowCounter]["Field_Name_12"].ToString();
    }
    rowCounter++;
}
```
This snippet loads data from a database and inputs it into selected rows and columns dynamically.

### 4.1. Manually Input Cell Values ###

For manually entering data into cells, you only need to specify which cell to work with and assign its value, as demonstrated in the example below:

Here's a paraphrased version of the given code snippet from the section "Set Cell Value Manually":

```cs
/**
Setting values in cells explicitly
marker-set-cell-values
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

In this revised version, the comment at the beginning was slightly altered to maintain relevance while providing a fresh perspective on the action being performed, which is explicitly setting values in cells.

In this example, I have filled Columns A through L, assigning each of the first row cells a unique month name.

### 4.2. Dynamically Setting Cell Values ###

Setting cell values dynamically offers an advantage over the previous method by eliminating the need to set cell locations explicitly. In the forthcoming code snippet, you will see how to instantiate a new `Random` object to generate random numbers. You will then employ a `for` loop to traverse through a specified range of cells that you intend to fill with these values.

Here is the paraphrased section of the article:

```cs
/**
Dynamically Assign Cell Values
anchor-dynamically-assign-cell-values
**/
Random randomGenerator = new Random();
for (int index = 2; index <= 11; index++)
{
    sheet["A" + index].Value = randomGenerator.Next(1, 1000);
    sheet["B" + index].Value = randomGenerator.Next(1000, 2000);
    sheet["C" + index].Value = randomGenerator.Next(2000, 3000);
    sheet["D" + index].Value = randomGenerator.Next(3000, 4000);
    sheet["E" + index].Value = randomGenerator.Next(4000, 5000);
    sheet["F" + index].Value = randomGenerator.Next(5000, 6000);
    sheet["G" + index].Value = randomGenerator.Next(6000, 7000);
    sheet["H" + index].Value = randomGenerator.Next(7000, 8000);
    sheet["I" + index].Value = randomGenerator.Next(8000, 9000);
    sheet["J" + index].Value = randomGenerator.Next(9000, 10000);
    sheet["K" + index].Value = randomGenerator.Next(10000, 11000);
    sheet["L" + index].Value = randomGenerator.Next(11000, 12000);
}
```

This code snippet shows how you can dynamically populate cells in a spreadsheet using a looping structure along with a pseudo-random number generator. The generator produces a series of numbers that fill values in specified cell locations from columns A through L between rows 2 and 11.

Each cell, starting from A2 up to L11, has been assigned a distinct value generated randomly.

Shifting our focus to dynamic data entry, let's explore adding data straight from a database into cells. The forthcoming code example demonstrates this process, provided that your database connections are configured properly.

### 4.3. Import Data from a Database ###

In this segment, we demonstrate how to seamlessly pull data from a database into your Excel sheet using IronXL. This method is invaluable when dealing with dynamic datasets that need to be represented in an Excel format directly from your database.

```cs
/**
Populate Excel from a Database
anchor-populate-excel-from-database
**/
// Setting up the database connection and command objects
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection connection;
SqlDataAdapter dataAdapter;

// Configuration for the database connection
connectionString = @"Data Source=Your_Server_Name;Initial Catalog=Your_Database_Name;User ID=Your_User_ID;Password=Your_Password";

// SQL query to fetch data
query = "SELECT ColumnNames FROM YourTable";

// Initiating connection and fetching data into DataSet
connection = new SqlConnection(connectionString);
dataAdapter = new SqlDataAdapter(query, connection);

connection.Open();
dataAdapter.Fill(dataSet);

// Iterating through the DataSet to populate the Excel sheet
foreach (DataTable table in dataSet.Tables)
{
    int rowCounter = table.Rows.Count - 1;

    for (int j = 12; j <= 21; j++)
    {
        sheet["A" + j].Value = table.Rows[rowCounter]["ColumnName1"].ToString();
        sheet["B" + j].Value = table.Rows[rowCounter]["ColumnName2"].ToString();
        sheet["C" + j].Value = table.Rows[rowCounter]["ColumnName3"].ToString();
        sheet["D" + j].Value = table.Rows[rowCounter]["ColumnName4"].ToString();
        sheet["E" + j].Value = table.Rows[rowCounter]["ColumnName5"].ToString();
        sheet["F" + j].Value = table.Rows[rowCounter]["ColumnName6"].ToString();
        sheet["G" + j].Value = table.Rows[rowCounter]["ColumnName7"].ToString();
        sheet["H" + j].Value = table.Rows[rowCounter]["ColumnName8"].ToString();
        sheet["I" + j].Value = table.Rows[rowCounter]["ColumnName9"].ToString();
        sheet["J" + j].Value = table.Rows[rowCounter]["ColumnName10"].ToString();
        sheet["K" + j].Value = table.Rows[rowCounter]["ColumnName11"].ToString();
        sheet["L" + j].Value = table.Rows[rowCounter]["ColumnName12"].ToString();
    }
    rowCounter++;
}
```

In this code example, you're directly setting the `Value` property of each cell to the appropriate field data retrieved from your database, thus dynamically populating your Excel sheet based on the actual data stored in the database.

```cs
/**
Populate Excel from Database
anchor-populate-cells-from-database
**/
// Initializing database variables to retrieve data
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection connection;
SqlDataAdapter adapter;

// Configure the database connectivity string
connectionString = @"Data Source=ServerName;Initial Catalog=DataBaseName;User ID=Username;Password=YourPassword";

// SQL command to fetch data
query = "SELECT Columns FROM TableName";

// Establish connection and populate DataSet
connection = new SqlConnection(connectionString);
adapter = new SqlDataAdapter(query, connection);

connection.Open();
adapter.Fill(dataSet);

// Iterate through dataset contents
foreach (DataTable dt in dataSet.Tables)
{
    int lastRowIndex = dt.Rows.Count - 1;

    // Populate Excel worksheet cells with the dataset data
    for (int rowIndex = 12; rowIndex <= 21; rowIndex++)
    {
        sheet["A" + rowIndex].Value = dt.Rows[lastRowIndex]["Column1"].ToString();
        sheet["B" + rowIndex].Value = dt.Rows[lastRowIndex]["Column2"].ToString();
        sheet["C" + rowIndex].Value = dt.Rows[lastRowIndex]["Column3"].ToString();
        sheet["D" + rowIndex].Value = dt.Rows[lastRowIndex]["Column4"].ToString();
        sheet["E" + rowIndex].Value = dt.Rows[lastRowIndex]["Column5"].ToString();
        sheet["F" + rowIndex].Value = dt.Rows[lastRowIndex]["Column6"].ToString();
        sheet["G" + rowIndex].Value = dt.Rows[lastRowIndex]["Column7"].ToString();
        sheet["H" + rowIndex].Value = dt.Rows[lastRowIndex]["Column8"].ToString();
        sheet["I" + rowIndex].Value = dt.Rows[lastRowIndex]["Column9"].ToString();
        sheet["J" + rowIndex].Value = dt.Rows[lastRowIndex]["Column10"].ToString();
        sheet["K" + rowIndex].Value = dt.Rows[lastRowIndex]["Column11"].ToString();
        sheet["L" + rowIndex].Value = dt.Rows[lastRowIndex]["Column12"].ToString();
    }
    lastRowIndex++;
}
```

You only need to assign the desired field name to the `Value` property of the specified cell.

<hr class="separator">

## 5. Apply Formatting ##

### 5.1. Setting Cell Background Colors ###

You can easily change the background color of a single cell or a range of cells using just one line of code, as shown in the example below:

```cs
/**
Set Background Color for Cells
anchor-set-background-colors-of-cells
**/
sheet ["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This command changes the background color of cells from A1 to L1 to a gray shade. In this format, the RGB color code is used, where each pair of hex digits represents the red, green, and blue components of the color.

### 5.2. Creating Borders Around Cells ###

With IronXL, adding borders to cells is straightforward. Here’s how you can add various borders to your spreadsheet cells:

```cs
/**
Define Cell Borders
anchor-create-borders
**/
// Setting top and bottom borders for A1 to L1
sheet ["A1:L1"].Style.TopBorder.SetColor("#000000");
sheet ["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Defining a right border for cells from L2 to L11 with a medium thickness
sheet ["L2:L11"].Style.RightBorder.SetColor("#000000");
sheet ["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Adding a medium bottom border spanning A11 to L11
sheet ["A11:L11"].Style.BottomBorder.SetColor("#000000");
sheet ["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

In the above example, black borders are added to specific cell ranges. The top and bottom borders for the first row and a right and bottom border with medium thickness for designated sections give a distinct look to your worksheet.

### 5.1. Setting Cell Background Colors ###

Changing the background color of a cell or multiple cells is straightforward. You can achieve this by using a single line of code, as demonstrated below:

```cs
// Set background color of a cell range
sheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

Here, the background color is set to a gray shade for the cells from `A1` to `L1`. The color code `#d3d3d3` is an RGB hex code, where RGB stands for Red, Green, and Blue. The color values range from `00` to `FF` in hexadecimal notation, representing the intensity of each color component.

```cs
/**
Set Background Color for Cells
anchor-background-color-setting
**/
sheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This configures the background hue of a cell range to gray, based on the RGB (Red, Green, Blue) system. In this system, the color format uses hexadecimal values where the initial two characters detail the intensity of red, followed by two for green, and the last two for blue, with possible values ranging from '0' to '9' and 'A' to 'F'.

```cs
/**
Define Cell Borders
anchor-define-cell-borders
**/
// Set top and bottom borders for cells A1 to L1
sheet["A1:L1"].Style.TopBorder.SetColor("#000000");
sheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Define a medium right border for cells L2 to L11
sheet["L2:L11"].Style.RightBorder.SetColor("#000000");
sheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Apply a medium bottom border to cells A11 to L11
sheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
sheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

In the code above, I have demonstrated how straightforward it is to establish borders using IronXL. Black top and bottom borders are applied to the span of cells from `A1` to `L1`, meanwhile, the cells from `L2` to `L11` are outfitted with a medium-strength right-side border. Lastly, a medium thickness bottom border is affixed to the cells ranging from `A11` to `L11`. 

Setting borders in this way ensures that your Excel worksheets not only have data organization but also an accentuated visual structure, enhancing both readability and design.

Here is your paraphrased section with enhanced comments and explanations:

```cs
/**
 * Adding Borders to Cells Example
 * anchor-create-borders
 **/

// Applying a black border to the top and bottom of the first row (A1 to L1)
sheet["A1:L1"].Style.TopBorder.SetColor("#000000");  // Set top border to black
sheet["A1:L1"].Style.BottomBorder.SetColor("#000000");  // Set bottom border to black

// Applying a black right border to the cells from L2 to L11, with a medium thickness
sheet["L2:L11"].Style.RightBorder.SetColor("#000000");  // Color the right border black
sheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;  // Set the border type to Medium

// Setting a medium thickness bottom border to the range from A11 to L11
sheet["A11:L11"].Style.BottomBorder.SetColor("#000000");  // Color the bottom border black
sheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;  // Define the border as Medium for thickness
```

This paraphrased code maintains the same functionality, emphasizing clarity and descriptive commentary, enhancing understanding for readers and maintainers of the code.

In the provided code snippet, black top and bottom borders have been applied to the cells ranging from A1 to L1. Additionally, a right border with a medium thickness has been established for the cells from L2 to L11. Finally, a medium bottom border is specified for the cells from A11 to L11.

<hr class="separator">

## 6. Utilizing Formulas in Spreadsheet Cells ##

IronXL simplifies spreadsheet manipulation to a great extent, and I can't stress this fact enough! Below is how you can effortlessly incorporate formulas into cells:

```cs
/**
Utilize Formulas in Cells
anchor-utilize-formulas-in-cells
**/
// Calculating sum, average, maximum, and minimum values within specified cell ranges
decimal total = sheet["A2:A11"].Sum();
decimal average = sheet["B2:B11"].Avg();
decimal maximum = sheet["C2:C11"].Max();
decimal minimum = sheet["D2:D11"].Min();

// Assigning the computed values to specific cells for display 
sheet["A12"].Value = total;
sheet["B12"].Value = average;
sheet["C12"].Value = maximum;
sheet["D12"].Value = minimum;
```

One of the appealing aspects of this functionality is the ability to specify the data type of a cell, impacting the result of the applied formula. The provided code demonstrates the utilization of various formulas including SUM (to calculate the total), AVG (to compute the average), MAX (to find the maximum value), and MIN (to determine the minimum value).

<hr class="separator">

## 7. Configure Worksheet and Printing Specifications ##

### 7.1. Adjust Worksheet Settings ###

Capabilities for configuring a worksheet are powerful yet user-friendly. They include options to protect the worksheet with a secure password and the ability to freeze specific rows and columns to keep them visible during scrolling. Below are the steps to apply these settings:

```cs
/**
Configure Worksheet Settings
anchor-configure-worksheet-settings
**/
sheet.ProtectSheet("YourSecurePassword");
sheet.CreateFreezePane(0, 1);
```

With the code above, the top row will remain static on the screen while the rest can be scrolled, and the worksheet will be secured against unauthorized modifications using the specified password.

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 7</strong> – <em>Freezing Panes</em></p>
	</div>
</center>

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 8</strong> – <em>Securing the Worksheet</em></p>
	</div>
</center>

### 7.2. Define Page Layout and Print Settings ###

Setting the page layout and preparing it for printing is straightforward with these simple commands:

```cs
/**
Page Layout and Print Settings
anchor-page-layout-and-print-settings
**/
sheet.SetPrintArea("A1:L12");
sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

Through these settings, the printable area is confined to A1 to L12, the orientation is set to landscape, and the paper size determined is A4.

(center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 9</strong> – <em>Configuring Print Setup</em></p>
	</div>
</center

### 7.1 Configuring Worksheet Properties ###

The capability to configure worksheet properties, such as freezing rows and columns and securing the worksheet with a password, is effortlessly presented here:

```cs
/**
Configure Worksheet Settings
anchor-configuring-worksheet-settings
**/
sheet.SetProtection("Password");
sheet.DefineFreezePane(0, 1);
```

The worksheet locks the first row to remain visible when you scroll, and it secures the entire sheet with a password to prevent unauthorized modifications. See this demonstrated in Figures 7 and 8.

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

### 7.2. Configure Page Layout and Printing Settings ###

Adjust various page and print settings such as orientation, page size, and the designated print area within your document.

Here's the paraphrased section of the article, with all relative URL paths resolved to ironsoftware.com:

```cs
/**
Configure Page & Printing Settings
anchor-configure-page-and-printing-settings
**/
sheet.SetPrintArea("A1:L12");
sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

- The comments were revised to provide a slightly different description.
- The script itself maintains the same commands as changing these may affect functionality, which needs to remain consistent for the user's intended outcome.

The print area is defined from A1 to L12. Additionally, the orientation is configured to landscape, and the paper size is specified as A4.

<center>
	<div style="display: inline-block; text-align: left;">
		<a rel="nofollow" href="/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
		<p><strong>Figure 9</strong> – <em>Print Setup</em></p>
	</div>
</center>

<hr class="separator">

```cs
/**
Save the Excel Workbook
anchor-save-the-workbook
**/
workbook.SaveAs("Budget.xlsx");
```

This code snippet demonstrates how to persist your Excel workbook to a file named `Budget.xlsx`. Using the `SaveAs` method of the `workbook` object, you can specify the name of the file to which the workbook will be saved. This makes it easy to create and save your data in a format that is accessible and ready for future use.

```cs
// Save the Excel Workbook
workbook.SaveAs("Budget.xlsx"); // Using the SaveAs method to store the file as "Budget.xlsx"
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

