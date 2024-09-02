# C# Excel File Creation Tutorial

In this guide, we'll walk you through the process of generating an Excel workbook on any system that supports .NET Framework 4.5 or .NET Core. With the use of C#, creating Excel files becomes straightforward without needing to rely on the **Microsoft.Office.Interop.Excel** library. Leverage IronXL to configure worksheet features such as pane freezing, worksheet protection, and printing options, among others.

<hr class="separator">

<h4 class="tutorial-segment-title">Overview</h4>




<h2>IronXL Creates C&num; Excel Files in .NET</h2>

[IronXL is a comprehensive C# & VB Excel API](https://ironsoftware.com/csharp/excel/) designed to enable reading, editing, and creating Excel files in .NET with exceptional speed. It eliminates the need for installing MS Office or the Excel Interop.

IronXL offers extensive support for multiple platforms and frameworks, including .NET Core, .NET Framework, Xamarin, Mobile, Linux, macOS, and Azure.

<h3>IronXL Features:</h3>

- Direct assistance from our .NET developers
- Quick setup using Microsoft Visual Studio
- No cost during development phase. Licensing starts at `$liteLicense`.

<h4>Create and Save an Excel File: Quick Code</h4>

<a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">https://www.nuget.org/packages/IronXL.Excel/</a>

Alternatively, you can directly download the `IronXL.dll` file from [this link](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and include it into your project.

```cs
using IronXL;

// Initialize a new workbook with default XLSX format, this can be customized with CreatingOptions
WorkBook newWorkBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet newSheet = newWorkBook.CreateWorkSheet("example_sheet");
newSheet["A1"].Value = "Example";

// Assigning a value to a range of cells
newSheet["A2:A4"].Value = 5;
newSheet["A5"].Style.SetBackgroundColor("#f0f0f0");

// Apply bold styling to a range of cells
newSheet["A5:A6"].Style.Font.Bold = true;

// Evaluating a formula
newSheet["A6"].Value = "=SUM(A2:A4)";
if (newSheet["A6"].IntValue == newSheet["A2:A4"].IntValue)
{
    Console.WriteLine("Basic test passed");
}

// Save the workbook to a file
newWorkBook.SaveAs("example_workbook.xlsx");
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Acquire the Free IronXL C# Library

Explore diverse methods for integrating the IronXL library into your projects:
1. **Visual Studio Installation**
2. **Developer Command Prompt Installation**
3. **Direct NuGet Package Download**

### Visual Studio Integration

Leverage Visual Studio’s built-in NuGet Package Manager to easily incorporate the IronXL library. Access the package manager through the Project Menu or by right-clicking your project within the Solution Explorer.

![Project Menu Visualization](https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png)
*<center><strong>Figure 3</strong> – The Project Menu Interface</center>*

![Solution Explorer Right-Click Menu](https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png)
*<center><strong>Figure 4</strong> – Right-clicking in the Solution Explorer</center>*

After selecting "Manage NuGet Packages", search and install the IronXL.Excel package as depicted below:

![Install IronXL Excel via NuGet Package Manager](https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png)
*<center><strong>Figure 5</strong> – Installing IronXL.Excel NuGet Package</center>*

### Developer Command Prompt Installation

Find the Developer Command Prompt under your Visual Studio installations, open it, and execute the following steps:

1. Input the command: `PM > Install-Package IronXL.Excel`
2. Press Enter.
3. The package is then installed.
4. Refresh your Visual Studio project session.

### Direct NuGet Package Download

Navigate and download the IronXL.Excel package directly from NuGet at:
[https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/)

Post-download:
- Execute the downloaded package.
- Reload the Visual Studio project to recognize the new library.

### Install IronXL via Direct Library Download

Alternatively, download the IronXL library directly using this link:
[Download IronXL Library](https://ironsoftware.com/csharp/excel/)

To incorporate IronXL into your project:

1. Right-click the Solution in the Solution Explorer.
2. Choose 'References'.
3. Locate and select the downloaded `IronXL.dll`.
4. Confirm by clicking 'OK'.

**Setup Complete!**
With IronXL installed, you’re ready to dive into creating, manipulating, and styling Excel files in your .NET applications with ease!

<h3>Install by Using NuGet</h3>

Here are three distinct methods you can use to install the IronXL NuGet package for enhanced control and functionality within your projects:

1. **Through Visual Studio**:  
   Utilize the NuGet Package Manager in Visual Studio to seamlessly integrate IronXL into your development environment. You can locate this option within the Project Menu or by right-clicking on your project in the Solution Explorer.

2. **Using Developer Command Prompt**:  
   Begin by finding the Developer Command Prompt, typically located in your Visual Studio directory. Once opened, type the command `Install-Package IronXL.Excel` and hit Enter. This straightforward command will handle the installation for you, after which, you should reload your Visual Studio project.

3. **Direct NuGet Package Download**:  
   If you prefer a direct download, visit the IronXL page on the NuGet website at [this link](https://www.nuget.org/packages/ironxl.excel/). From there, click on "Download Package". Once the download is complete, proceed to open the package file, and don’t forget to reload your Visual Studio project afterwards to complete the integration process.

<h3>Visual Studio</h3>

Visual Studio offers a convenient NuGet Package Manager that helps in integrating NuGet packages into your projects. This can be accessed through the Project Menu or by right-clicking on your project within the Solution Explorer. These methods are depicted below in Figures 3 and 4.

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
<br>

After selecting the 'Manage NuGet Packages' option through any of the methods described, search for and install the IronXL.Excel package, as demonstrated in Figure 5.

<br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

<h3>Developer Command Prompt</h3>

Launch the Developer Command Prompt by searching for it typically within your Visual Studio directory. Here's how to proceed with installing the IronXL.Excel NuGet package: 

1. Locate your Developer Command Prompt, usually found in the Visual Studio installation directory.
2. Enter the command:
   
   ```
   PM> Install-Package IronXL.Excel
   ```

3. Hit the Enter key to execute.
4. The IronXL.Excel package will now be installed to your project.
5. Refresh or reload your Visual Studio project to apply changes.

<h3>Download the NuGet Package directly</h3>

Here's the paraphrased section of the article:

-----

To acquire and install the NuGet package, follow these instructions:

1. Go to [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/) in your web browser.

2. Select the 'Download Package' option.

3. Once the download is complete, open the package by double-clicking on the file.

4. Finally, restart your Visual Studio project to complete the setup.

<h3>Install IronXL by Direct Download of the Library</h3>

An alternative method for installing IronXL involves directly downloading it from the official website. You can access the download link here: [https://ironsoftware.com/csharp/excel/](https://ironsoftware.com/csharp/excel/).

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library</em></p>
  </div>
</center>

Here's a paraphrased version of the given section, with relative URL paths resolved to ironsoftware.com:

---

To include the library in your project, follow these steps:

1. In the Solution Explorer, right-click on the Solution.
2. Choose 'References' from the context menu.
3. Look for the `IronXL.dll` file in the dialogue.
4. Confirm by clicking 'OK'.

<h3>Let's Go!</h3>

Now that everything is configured, let's dive into exploring the fantastic capabilities of the IronXL library!

<hr class="separator">

<h4 class="tutorial-segment-title">How to Tutorials</h4>

## 2. Establish an ASP.NET Project ##

Initiate your journey by setting up an ASP.NET project following these steps:

1. Begin by navigating to [IronXL NuGet Package](https://www.nuget.org/packages/ironxl.excel/).
2. Proceed to download the package.
3. Once downloaded, execute the package to integrate it into your environment.
4. Ensure to refresh your Visual Studio instance to reflect the changes.

Next, let’s proceed to structuring an ASP.NET website:

1. Open Microsoft Visual Studio.
2. From the menu, select `File > New Project`.
3. In the project type selections, choose Web under Visual C#.
4. Opt for 'ASP.NET Web Application' as shown below.

![Creating a new ASP.NET project](https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png)
**Figure 1** – *Creating a New Project*

5. Confirm your selection by clicking `OK`.
6. On the subsequent screen, select 'Web Forms' illustrated here:

![Selecting Web Forms in project setup](https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png)
**Figure 2** – *Opting for Web Forms*
7. Finalize by clicking `OK`.

With the setup complete, you are now ready to elevate your project by integrating IronXL, enabling you to customize your Excel handling capabilities efficiently.

<ul>
  <li>Navigate to the following URL:</li>
  <li><a href="https://www.nuget.org/packages/ironxl.excel/" target="_blank">https://www.nuget.org/packages/ironxl.excel/</a></li>
  <li>Click on Download Package</li>
  <li>After the package has downloaded, double click it</li>
  <li>Reload your Visual Studio project</li>
</ul>

Here's the paraphrased section with links and image paths resolved to ironsoftware.com:

------

Follow these steps to set up an ASP.NET Website:

1. Launch Visual Studio.
2. Navigate to `File` and choose `New Project`.
3. Choose 'Web' from the Project type options under Visual C#.
4. Select 'ASP.NET Web Application', as illustrated below:

   ![ASP.NET Web Application Setup](https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png)

   **Figure 1** – Setting up a New Project

<br>
<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

Here's the paraphrased section:

---
<strong style="margin-left: 40px;">Figure 1</strong> – *New Project*

5. Press OK.

6. In the subsequent interface, choose Web Forms as illustrated below in Figure 2.

<br>

<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

**Figure 2** – *Web Forms*
```
The paraphrase transforms the styling and reaffirms the text, focusing solely on the label of the figure referenced in the document, without altering the essential content.

<br>

After clicking OK, you're all set to begin. Install IronXL and start personalizing your Excel file.

<hr class="separator">

Creating a new Excel Workbook with IronXL is remarkably straightforward—it only takes a single line of code! Here's how you can do it:
```cs
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
```

Here's the paraphrased version of the given code snippet:

```cs
// Initialize a new Excel workbook with an XLSX file format
WorkBook myWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL supports the creation of both XLS (the traditional Excel file format) and XLSX (the modern, XML-based file format) for Excel files.

### 3.1. Creating a Default Worksheet ###

Creating a default worksheet is incredibly straightforward:

```cs
WorkSheet workSheet = workBook.CreateWorkSheet("2020 Budget");
```

In the example above, `"Sheet"` represents the worksheet where you'll be able to assign values to cells and perform nearly every function available in Excel. 

If you're wondering about the difference between a Workbook and a Worksheet, here's a quick clarification: a Workbook is comprised of one or more Worksheets. You can add multiple Worksheets to a single Workbook. Later articles will explore how to do this. Each Worksheet consists of Rows and Columns, where the intersection of a Row and a Column forms a Cell. These Cells are what you'll manipulate as you work with Excel data.

```cs
// Initialize a new worksheet with the name "2020 Budget"
WorkSheet workSheet = workBook.AddWorkSheet("2020 Budget");
```

In the code example above, "Sheet" refers to the Excel worksheet where you can assign values to individual cells and perform nearly all operations available in Excel.

For those unsure of the distinction between a Workbook and a Worksheet, here's a straightforward clarification:

A Workbook is essentially a collection of Worksheets—it's the entire file containing your sheets. You can include multiple Worksheets in a single Workbook as needed, which will be covered in depth in a forthcoming discussion. Within each Worksheet, you have Rows and Columns that intersect to form Cells—the fundamental elements you interact with while using Excel functionalities.

<hr class="separator">

## 4. Assigning Values to Cells ##

### 4.1. Manually Assign Values to Cells ###

Setting values to cells directly is straightforward. Specify the cell by its address, and then assign its value. Below is how you populate the cells from Columns A to L in the first row, giving each a different month name:

```cs
workSheet["A1"].Value = "January";
workSheet["B1"].Value = "February";
workSheet["C1"].Value = "March";
workSheet["D1"].Value = "April";
workSheet["E1"].Value = "May";
workSheet["F1"].Value = "June";
workSheet["G1"].Value = "July";
workSheet["H1"].Value = "August";
workSheet["I1"].Value = "September";
workSheet["J1"].Value = "October";
workSheet["K1"].Value = "November";
workSheet["L1"].Value = "December";
```

### 4.2. Dynamically Assign Values to Cells ###

Dynamic assignment of values is similar to manual assignment but eliminates the need to specify cell addresses explicitly. You can use a loop along with a `Random` object to fill cells with random numbers:

```cs
Random random = new Random();
for (int i = 2; i <= 11; i++)
{
    workSheet["A" + i].Value = random.Next(1, 1000);
    workSheet["B" + i].Value = random.Next(1000, 2000);
    workSheet["C" + i].Value = random.Next(2000, 3000);
    workSheet["D" + i].Value = random.Next(3000, 4000);
    workSheet["E" + i].Value = random.Next(4000, 5000);
    workSheet["F" + i].Value = random.Next(5000, 6000);
    workSheet["G" + i].Value = random.Next(6000, 7000);
    workSheet["H" + i].Value = random.Next(7000, 8000);
    workSheet["I" + i].Value = random.Next(8000, 9000);
    workSheet["J" + i].Value = random.Next(9000, 10000);
    workSheet["K" + i].Value = random.Next(10000, 11000);
    workSheet["L" + i].Value = random.Next(11000, 12000);
}
```

Each cell between A2 and L11 will now contain a unique number generated randomly.

### 4.3. Inserting Data Directly from a Database ###

To dynamically populate cells straight from a database, set up the appropriate connections and execute a SQL query that fetches the data. Once retrieved, loop through the results to assign values to cells:

```cs
// Create database objects to retrieve data
string connectionString;
string sqlQuery;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection connection;
SqlDataAdapter dataAdapter;

// Specify database connection details
connectionString = @"Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserId;Password=Password";

// SQL query to fetch data
sqlQuery = "SELECT ColumnNames FROM TableName";

// Prepare and execute query
connection = new SqlConnection(connectionString);
dataAdapter = new SqlDataAdapter(sqlQuery, connection);
connection.Open();
dataAdapter.Fill(dataSet);

// Assign values to cells from dataset
foreach (DataTable table in dataSet.Tables)
{
    int rowCount = table.Rows.Count - 1;
    for (int j = 12; j <= 21; j++)
    {
        workSheet["A" + j].Value = table.Rows[rowCount]["ColumnName1"].ToString();
        workSheet["B" + j].Value = table.Rows[rowCount]["ColumnName2"].ToString();
        // Continue for other columns
    }
    rowCount++;
}
```

This approach allows you to set the `Value` property of cells to the respective field names from your database effortlessly.

### 4.1. Manually Inputting Cell Values ###

When manually inputting data into cells, designate the specific cell and assign it a value. Here's a brief demonstration:

```cs
workSheet["A1"].Value = "January";
workSheet["B1"].Value = "February";
workSheet["C1"].Value = "March";
workSheet["D1"].Value = "April";
workSheet["E1"].Value = "May";
workSheet["F1"].Value = "June";
workSheet["G1"].Value = "July";
workSheet["H1"].Value = "August";
workSheet["I1"].Value = "September";
workSheet["J1"].Value = "October";
workSheet["K1"].Value = "November";
workSheet["L1"].Value = "December";
```
In this example, the cells from `A1` to `L1` across the first row are populated with the names of each month, providing a clear, month-by-month layout.
```

```cs
// Assigning month names to the first row of respective columns
workSheet["A1"].Value = "January";
workSheet["B1"].Value = "February";
workSheet["C1"].Value = "March";
workSheet["D1"].Value = "April";
workSheet["E1"].Value = "May";
workSheet["F1"].Value = "June";
workSheet["G1"].Value = "July";
workSheet["H1"].Value = "August";
workSheet["I1"].Value = "September";
workSheet["J1"].Value = "October";
workSheet["K1"].Value = "November";
workSheet["L1"].Value = "December";
```

I have filled each column from A to L with the names of different months, each appearing in the first row.

### 4.2. Dynamically Assigning Values to Cells ###

Adding values dynamically offers a flexible alternative, sparing the need to specify fixed cell locations. This next code sample demonstrates setting up a new `Random` instance to generate random numbers. You'll use a loop to traverse a specified range of cells where these generated values will be inserted.

```cs
Random random = new Random();
for (int i = 2; i <= 11; i++)
{
    workSheet["A" + i].Value = random.Next(1, 1000);
    workSheet["B" + i].Value = random.Next(1000, 2000);
    workSheet["C" + i].Value = random.Next(2000, 3000);
    workSheet["D" + i].Value = random.Next(3000, 4000);
    workSheet["E" + i].Value = random.Next(4000, 5000);
    workSheet["F" + i].Value = random.Next(5000, 6000);
    workSheet["G" + i].Value = random.Next(6000, 7000);
    workSheet["H" + i].Value = random.Next(7000, 8000);
    workSheet["I" + i].Value = random.Next(8000, 9000);
    workSheet["J" + i].Value = random.Next(9000, 10000);
    workSheet["K" + i].Value = random.Next(10000, 11000);
    workSheet["L" + i].Value = random.Next(11000, 12000);
}
```
This approach eliminates the need for predefined cell locations, allowing values to be inserted dynamically across a designated cell range.

```cs
// Initialize random number generator
Random randomGenerator = new Random();

// Populate cells from A2 to L11 with random values within specified ranges
for (int row = 2; row <= 11; row++)
{
    workSheet[$"A{row}"].Value = randomGenerator.Next(1, 1000);    // Values between 1 and 1000
    workSheet[$"B{row}"].Value = randomGenerator.Next(1000, 2000); // Values between 1000 and 2000
    workSheet[$"C{row}"].Value = randomGenerator.Next(2000, 3000); // Values between 2000 and 3000
    workSheet[$"D{row}"].Value = randomGenerator.Next(3000, 4000); // Values between 3000 and 4000
    workSheet[$"E{row}"].Value = randomGenerator.Next(4000, 5000); // Values between 4000 and 5000
    workSheet[$"F{row}"].Value = randomGenerator.Next(5000, 6000); // Values between 5000 and 6000
    workSheet[$"G{row}"].Value = randomGenerator.Next(6000, 7000); // Values between 6000 and 7000
    workSheet[$"H{row}"].Value = randomGenerator.Next(7000, 8000); // Values between 7000 and 8000
    workSheet[$"I{row}"].Value = randomGenerator.Next(8000, 9000); // Values between 8000 and 9000
    workSheet[$"J{row}"].Value = randomGenerator.Next(9000, 10000); // Values between 9000 and 10000
    workSheet[$"K{row}"].Value = randomGenerator.Next(10000, 11000); // Values between 10000 and 11000
    workSheet[$"L{row}"].Value = randomGenerator.Next(11000, 12000); // Values between 11000 and 12000
}
```

Each cell in the range from A2 to L11 is filled with a unique, randomly generated value.

Discussing dynamic data entry, let's explore how to automatically insert data into cells from a database. The following code example demonstrates this process, provided that your database connections are properly configured.

### 4.3. Insert Data from a Database ###

This section demonstrates how easily you can populate your Excel cells with data straight from a database. First, ensure your database connections are properly configured.

```cs
// Setting up database connection objects
string connectionString;
string query;
DataSet dataSet = new DataSet("MyDataSet");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Define the connection string
connectionString = @"Data Source=Your_Server;Initial Catalog=Your_Database;User ID=Your_Username;Password=Your_Password";

// SQL statement to fetch data
query = "SELECT Columns FROM Your_Table";

// Establish the connection and populate the DataSet
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(query, sqlConnection);
sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Iterate through DataSet and assign data to cells
foreach (DataTable dataTable in dataSet.Tables)
{
    int totalRows = dataTable.Rows.Count - 1;
    for (int index = 12; index <= 21; index++)
    {
        workSheet["A" + index].Value = dataTable.Rows[totalRows]["Column1"].ToString();
        workSheet["B" + index].Value = dataTable.Rows[totalRows]["Column2"].ToString();
        workSheet["C" + index].Value = dataTable.Rows[totalRows]["Column3"].ToString();
        workSheet["D" + index].Value = dataTable.Rows[totalRows]["Column4"].ToString();
        workSheet["E" + index].Value = dataTable.Rows[totalRows]["Column5"].ToString();
        workSheet["F" + index].Value = dataTable.Rows[totalRows]["Column6"].ToString();
        workSheet["G" + index].Value = dataTable.Rows[totalRows]["Column7"].ToString();
        workSheet["H" + index].Value = dataTable.Rows[totalRows]["Column8"].ToString();
        workSheet["I" + index].Value = dataTable.Rows[totalRows]["Column9"].ToString();
        workSheet["J" + index].Value = dataTable.Rows[totalRows]["Column10"].ToString();
        workSheet["K" + index].Value = dataTable.Rows[totalRows]["Column11"].ToString();
        workSheet["L" + index].Value = dataTable.Rows[totalRows]["Column12"].ToString();
    }
    totalRows++;
}
```

In this snippet, after establishing a connection to your database, a `DataSet` is filled using an SQL query. The data is then looped through, and values are dynamically assigned to cells in your Excel document. This process automates data entry, making it swift and error-free.

```cs
// Initializing database components to fetch data
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Establish database connection details
connectionString = @"Data Source=Your_Server_Name;Initial Catalog=Your_Database_Name;User ID=Your_User_ID;Password=Your_Password";

// Define SQL command to retrieve data
query = "SELECT Columns_Names FROM Your_Table_Name";

// Connecting and populating the DataSet
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(query, sqlConnection);
sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Iterating through DataSet to update worksheet values
foreach (DataTable table in dataSet.Tables)
{
    int rowCount = table.Rows.Count - 1;
    for (int row = 12; row <= 21; row++)
    {
        workSheet["A" + row].Value = table.Rows[rowCount]["Column1_Name"].ToString();
        workSheet["B" + row].Value = table.Rows[rowCount]["Column2_Name"].ToString();
        workSheet["C" + row].Value = table.Rows[rowCount]["Column3_Name"].ToString();
        workSheet["D" + row].Value = table.Rows[rowCount]["Column4_Name"].ToString();
        workSheet["E" + row].Value = table.Rows[rowCount]["Column5_Name"].ToString();
        workSheet["F" + row].Value = table.Rows[rowCount]["Column6_Name"].ToString();
        workSheet["G" + row].Value = table.Rows[rowCount]["Column7_Name"].ToString();
        workSheet["H" + row].Value = table.Rows[rowCount]["Column8_Name"].ToString();
        workSheet["I" + row].Value = table.Rows[rowCount]["Column9_Name"].ToString();
        workSheet["J" + row].Value = table.Rows[rowCount]["Column10_Name"].ToString();
        workSheet["K" + row].Value = table.Rows[rowCount]["Column11_Name"].ToString();
        workSheet["L" + row].Value = table.Rows[rowCount]["Column12_Name"].ToString();
    }
    rowCount++;
}
```

You just need to assign the Value property of the specific cell to the Field name that you want to enter into the cell.

<hr class="separator">

### 5. Formatting Cells ###

Properly styling cells in a spreadsheet can significantly enhance the visual appeal and readability of your data. Using IronXL, you have a comprehensive set of formatting tools at your disposal to customize your spreadsheets.

#### 5.1 Set Cell Background Colors ####

Changing the background color of cells or cell ranges in IronXL is straightforward. Here is an example of how to do it:

```cs
workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

In this line of code, we're changing the background color of cells from A1 through L1 to a gentle gray. The hexadecimal value `#d3d3d3` represents the color, where the first two characters are the red component, the next two are green, and the last two are blue.

#### 5.2 Create Cell Borders ####

Creating and customizing borders for your data cells helps to define clear boundaries and can make the data easier to read and analyze. Here's how you can set up borders with IronXL:

```cs
workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

This snippet will add a black color to the top and bottom borders of the cells from A1 to L1 and adjust the right border from L2 to L11. Additionally, it sets the border style to 'Medium' for specific ranges, ensuring that your cells are well-defined and neatly separated.

By utilizing these formatting options, you can significantly improve the presentation and functionality of your Excel documents created with IronXL.

### 5.1. Apply Cell Background Colors ###

Applying a background color to a single cell or a group of cells is straightforward. You only need to use the following code snippet:

```cs
workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This code example sets the background color to a shade of gray across the range of cells from A1 to L1. The color value is defined in a hex format, where the components of Red, Green, and Blue are specified in the string, ranging from 00 to FF (Hexadecimal representation).
```

Here's the paraphrased section:

```cs
// Apply a gray background color to cells from A1 to L1
workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This changes the background hue of a selected range of cells to a gray shade. This color specification follows the RGB (Red, Green, Blue) system, wherein the color code begins with two letters or numbers representing red, followed by two for green, and two for blue. Each pair can vary from the numerals 0 to 9 and extend through the hexadecimal alphabet from A to F.

### 5.2. Add Borders to Cells ###

Inserting borders into cells using IronXL is straightforward; here’s how it's done:
```cs
// Set the top and bottom border color to black for cells from A1 to L1
workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Set the right border color and style to medium for cells from L2 to L11
workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Similarly, set the bottom border color and type for cells from A11 to L11
workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

Here's the paraphrased section of the article:

```cs
// Set the color of the top and bottom borders of cells A1 to L1 to black
workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Configure the right border for cells from L2 to L11, setting the color to black and border style to medium
workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Establish a medium black bottom border for the range from A11 to L11
workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

This snippet demonstrates setting border colors and styles for specific cell ranges within a worksheet using IronXL.

In the provided code snippet, I have applied black top and bottom borders to the cell range from A1 to L1. Additionally, I have implemented right borders on the cells from L2 to L11 with a medium thickness. Finally, I have configured the cells from A11 to L11 with a medium thickness bottom border.

<hr class="separator">

## 6. Use Formulas in Cells ##

IronXL's user-friendly nature simplifies tasks significantly, something I cannot emphasize enough. Implementing various formulas in cells is straightforward with the following code:

```cs
// Implementing IronXL's built-in aggregation functions
decimal total = workSheet["A2:A11"].Sum();  // Calculate the sum of values in range A2:A11
decimal average = workSheet["B2:B11"].Avg();  // Compute the average of values in range B2:B11
decimal highest = workSheet["C2:C11"].Max();  // Find the maximum value in range C2:C11
decimal lowest = workSheet["D2:D11"].Min();  // Determine the minimum value in range D2:D11

// Setting the calculated values into cells
workSheet["A12"].Value = total;  // Place the total in cell A12
workSheet["B12"].Value = average;  // Place the average in cell B12
workSheet["C12"].Value = highest;  // Place the highest value in cell C12
workSheet["D12"].Value = lowest;  // Place the lowest value in cell D12
```

The appealing aspect of this functionality is that it allows you to designate the cell's data type to match the output of the formula. The code snippet provided illustrates the application of various aggregation functions: `SUM` to total values, `AVG` for averaging values, `MAX` to find the maximum value, and `MIN` to determine the minimum value.

<hr class="separator">

## 7. Configuring Worksheet and Print Settings ##

### 7.1. Worksheet Configuration ###

You can enhance your worksheet with features such as freezing panes and password protection to safeguard your data. For example, to freeze the first row in the worksheet so it remains visible while you scroll, and protect the worksheet from unwanted edits, you would use the following code:

```cs
workSheet.ProtectSheet("Password");  // Protects the worksheet with a password
workSheet.CreateFreezePane(0, 1);    // Freezes the first row
```
Below are illustrations showcasing the freezing of panes and how a protected worksheet appears:

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 7</strong> – <em>Freeze Panes</em></p>
  </div>
</center>

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 8</strong> – <em>Protected Worksheet</em></p>
  </div>
</center>

### 7.2. Print and Page Setup ###

Configuring the print settings is straightforward with IronXL. You can determine the print area, page orientation, and the size of the paper using the code snippet below:

```cs
workSheet.SetPrintArea("A1:L12");                                       // Defines the area of the worksheet to be printed
workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;  // Sets the page orientation to landscape
workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;          // Sets the paper size to A4
```

The print setup parameters ensure that everything from A1 to L12 is formatted correctly on an A4 sheet in landscape orientation. See the setup in action below:

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Print Setup</em></p>
  </div>
</center>

These enhancements and settings provide robust control over both the functional and aesthetic aspects of your worksheets, making them more efficient and secure for professional use.

### 7.1. Configuring Worksheet Options ###

Adjusting worksheet settings allows you to freeze specific rows and columns to facilitate easier navigation, as well as secure the worksheet with a password for enhanced protection. The following examples illustrate these configurations:

```cs
workSheet.ProtectSheet("Password"); // Protects the worksheet with a password
workSheet.CreateFreezePane(0, 1); // Freezes the first row so it remains visible while scrolling
```

In this setup, the first row of the worksheet will remain fixed at the top as you scroll down, while the entire worksheet is safeguarded against unauthorized modifications using a password. Detailed demonstrations of these features can be seen in Figures 7 and 8.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 7</strong> – <em>Freeze Panes</em></p>
  </div>
</center>

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 8</strong> – <em>Protected Worksheet</em></p>
  </div>
</center>

Below is the paraphrased content for the given section, with relative URL paths properly resolved:

```cs
// Protect the worksheet from unauthorized editing using a password
workSheet.ProtectSheet("SuperSecret");

// Freeze the top row to remain visible during scrolling
workSheet.CreateFreezePane(0, 1);
```

In this configuration, the first row of the worksheet remains static and won't move with the rest of the sheet during scrolling. Furthermore, to safeguard the data, the worksheet is secured with a password preventing unauthorized modifications. See this functionality demonstrated in Figures 7 and 8.

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

Adjust page configurations, including page orientation, dimensions, and designated printing areas, among other settings.

Here's a paraphrased version of the provided section, with enhanced inline code comments and slightly altered code structure:

```cs
// Define the print area as the range from A1 to L12 on the worksheet
workSheet.SetPrintArea("A1:L12");

// Configure the print orientation to be horizontal
workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;

// Specify the paper size as A4
workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

The printable area is defined as ranging from cell A1 to L12. The orientation for printing is configured to be landscape, and the selected paper size is A4.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Print Setup</em></p>
  </div>
</center>

<hr class="separator">

```cs
// Code to save the Workbook
workBook.SaveAs("Budget.xlsx");
```

```cs
// Saving the workbook to a file named "Budget.xlsx"
workBook.SaveAs("Budget.xlsx");
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
      <h3>Download this Tutorial as C&num; Source Code</h3>
      <p>The full free C&num; for Excel Source Code for this tutorial is available to download as a zipped Visual Studio 2017 project file.</p>
      <a class="btn btn-white3" href="/csharp/excel/tutorials/downloads/Using.CSharp.to.Create.Excel.Files.in.Net.zip">
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
      <a class="doc-link" href="https://github.com/iron-software/tutorials/tree/master/IronXL/Using%20C%23%20to%20Create%20Excel%20Files%20in%20.Net" target="_blank">How to Create Excel File in C&num; on GitHub<i class="fa fa-chevron-right"></i></a>
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

