# C# Excel File Creation Tutorial

***Based on <https://ironsoftware.com/tutorials/create-excel-file-net/>***


In this tutorial, we'll walk you through the process of creating an Excel Workbook on any platform compatible with either .NET Framework 4.5 or .NET Core. Creating Excel files using C# is straightforward and does not rely on the outdated **Microsoft.Office.Interop.Excel** library. Learn to utilize IronXL to manage worksheet attributes such as freezing panes and adding protection, as well as configuring printing settings and more.

<hr class="separator">

<h4 class="tutorial-segment-title">Overview</h4>




<h2>IronXL Creates C&num; Excel Files in .NET</h2>

[IronXL offers a comprehensive C# & VB Excel API](https://ironsoftware.com/csharp/excel/) designed for high-speed manipulation of Excel spreadsheets in .NET environments. This eliminates the need for installing MS Office or Excel Interop components.

The IronXL library is compatible with multiple platforms including .NET Core, .NET Framework, Xamarin, Mobile, Linux, macOS, and Azure.

<h3>IronXL Features:</h3>

- Direct assistance from our dedicated .NET engineering team

- Quick and seamless setup through Microsoft Visual Studio

- No cost for development phase. Licensing starts from `$liteLicense`.

<h4>Create and Save an Excel File: Quick Code</h4>

<a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">https://www.nuget.org/packages/IronXL.Excel/</a>

Alternatively, you can download the `IronXL.dll` directly via this [link](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and incorporate it into your project.

```cs
using IronXL;

// The default file format is XLSX, which can be modified with CreatingOptions
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet worksheet = workbook.CreateWorkSheet("example_sheet");
worksheet["A1"].Value = "Example";

// Assigning a single value to multiple cells
worksheet["A2:A4"].Value = 5;
worksheet["A5"].Style.SetBackgroundColor("#f0f0f0");

// Applying bold style to a range of cells
worksheet["A5:A6"].Style.Font.Bold = true;

// Implementing a formula and checking its result
worksheet["A6"].Value = "=SUM(A2:A4)";
if (worksheet["A6"].IntValue == worksheet["A2:A4"].IntValue)
{
    Console.WriteLine("Basic test passed");
}
workbook.SaveAs("example_workbook.xlsx");
```
This revised version maintains the original operations while enhancing readability and structure in the code comments and actions.

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Get the FREE IronXL C# Library

Start by downloading the no-cost IronXL C# Library, which is critical for your .NET projects involving Excel file manipulations.

---

<h3>Install by Using NuGet</h3>

The IronXL NuGet package can be installed using three distinct methods:

1. **Visual Studio**: Utilize the Visual Studio integrated NuGet Package Manager. This option is available via the "Project" menu or by right-clicking on your project within the Solution Explorer to navigate to the package installation settings.

2. **Developer Command Prompt**: Accessible through the Visual Studio directory, the Developer Command Prompt allows you to install packages using straightforward commands. Simply open it and enter `PM> Install-Package IronXL.Excel`, then hit Enter to commence the installation process. Refresh your project in Visual Studio to finalize the setup.

3. **Direct NuGet Package Download**: Alternatively, the IronXL package can be obtained directly from its NuGet page. Visit [https://www.nuget.org/packages/IronXL.Excel](https://www.nuget.org/packages/IronXL.Excel), select "Download Package", and upon download completion, integrate it with your project manually in Visual Studio.

<h3>Visual Studio</h3>

Visual Studio offers a built-in NuGet Package Manager, making it convenient to add NuGet packages to your projects. It is accessible from the Project Menu or by right-clicking on your project within the Solution Explorer. These methods are illustrated in Figures 3 and 4 below.

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

Once you’ve accessed Manage NuGet Packages through either of the described methods, search for the IronXL.Excel package and proceed to install it, as depicted in Figure 5.

<br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

<h3>Developer Command Prompt</h3>

Launch the Developer Command Prompt and execute the following instructions to install the IronXL.Excel NuGet package:

1. Locate the Developer Command Prompt, typically found within your Visual Studio installation directory.

2. Enter this command:

   ```
   PM> Install-Package IronXL.Excel
   ```

3. Hit Enter.

4. This will initiate the installation of the package.

5. Once the installation is complete, refresh your Visual Studio project to apply the changes.

<h3>Download the NuGet Package directly</h3>

Here's a paraphrased version of the section with URL paths resolved to ironsoftware.com:

-----
To successfully download the NuGet package, follow these steps:

1. Visit this URL: [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/)

2. Select the 'Download Package' option.

3. Once the download is complete, double-click the downloaded file.

4. Refresh or reload your Visual Studio project to apply changes.

<h3>Install IronXL by Direct Download of the Library</h3>

Here's your paraphrased section with the resolved URL path:

-----
Another method to set up IronXL is via direct download. Access the file from this link: [https://ironsoftware.com/csharp/excel/](https://ironsoftware.com/csharp/excel/).

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library</em></p>
  </div>
</center>

To include the IronXL library in your .NET project, follow these simple steps:

1. In Solution Explorer, right-click on 'Solution'.
2. Choose 'References' from the context menu.
3. Search for the `IronXL.dll` in the browse window.
4. Confirm by clicking 'OK'.

<h3>Let's Go!</h3>

With everything set up, it's time to dive into the powerful capabilities of the IronXL library!

<hr class="separator">

<h4 class="tutorial-segment-title">How to Tutorials</h4>

## 2. Building an ASP.NET Project ##

Get started by downloading the IronXL `.nuget` package through this link:
[NuGet Package for IronXL](https://www.nuget.org/packages/ironxl.excel/). Once downloaded, open the package to integrate it with your Visual Studio project.

The following steps will help you create an ASP.NET website:
1. Open Visual Studio.
2. Go to `File` and select `New Project`.
3. In the project type selection menu, choose `Web` under the Visual C# category.
4. Choose `ASP.NET Web Application` as depicted below.

<div align="center">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
</div>
<div style="margin-left: 40px;"><strong>Figure 1</strong> – <em>New Project window in Visual Studio</em></div>
    
5. Confirm with `OK`.
6. On the following screen, opt for `Web Forms` as illustrated in the figure below.
<div align="center">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
</div>
<div style="margin-left: 40px;"><strong>Figure 2</strong> – <em>Selecting Web Forms</em></div>

7. Click `OK` to continue.

You now have a base to start enriching with IronXL capabilities.

<ul>
  <li>Navigate to the following URL:</li>
  <li><a href="https://www.nuget.org/packages/ironxl.excel/" target="_blank">https://www.nuget.org/packages/ironxl.excel/</a></li>
  <li>Click on Download Package</li>
  <li>After the package has downloaded, double click it</li>
  <li>Reload your Visual Studio project</li>
</ul>

Here's a paraphrased version of the section provided, with updated relative paths for links and images:

----
Follow these instructions to start building an ASP.NET website:

1. Initiate Visual Studio.

2. Navigate to 'File' and then to 'New Project'.

3. In the project types, choose 'Web' located under Visual C#.

4. Choose 'ASP.NET Web Application' as illustrated below:

<br>
<center>
<a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>
<strong style="margin-left: 40px;">Figure 1</strong> – *New Project*

----

<br>
<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

**Figure 1** - *Starting a New Project*

5. Press OK.

6. On the following page, choose 'Web Forms' as depicted below in Figure 2.

<br>

<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 2</strong> – *Web Forms*
```

<br>

After confirming your settings by clicking OK, your project environment is ready. The next step is to integrate IronXL, allowing you to modify and enhance your files as you see fit.

<hr class="separator">

## 3. Creating an Excel Workbook ##

Creating a new Excel Workbook with IronXL is incredibly straightforward – it only takes a single line of code! Here’s how simple it is:

```cs
// Initializing a new Excel workbook with the default XLSX file format
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
```

IronXL is compatible with both the XLS (the original Excel file version) and XLSX (the modern Excel file format).

### 3.1. Creating a Default Worksheet ###

Creating a default worksheet is straightforward and quick:

Here's a paraphrased version of the section specified:

```cs
// Initialize a new worksheet named "2020 Budget" in the workbook
WorkSheet budgetSheet = workBook.CreateWorkSheet("2020 Budget");
```

In the provided code snippet, "Sheet" denotes a worksheet which can be utilized to assign values to cells and perform nearly all tasks possible with Excel.

Should you find yourself puzzled about the distinction between a Workbook and a Worksheet, allow me to clarify:

A Workbook is essentially a container for one or more Worksheets. You have the flexibility to include numerous Worksheets within a single Workbook. Details on how to manage multiple Worksheets will be covered in an upcoming article. Within each Worksheet, you are presented with Rows and Columns. The point where a Row and a Column intersect is called a Cell. Cells are the primary elements that you will engage with when operating in Excel.

<hr class="separator">

## 4. Defining Cell Values ##

### 4.1. Manually Assigning Values to Cells ###

Setting the values of cells is straightforward. Here's how you specify the content for individual cells:

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
In this example, each cell from A1 to L1 across the first row is populated with the name of a month.

### 4.2. Dynamically Populating Cell Values ###

To dynamically assign cell values, use the following approach that doesn’t require you to hard-code the cell positions:

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
This segment of code will auto-generate unique values ranging from specified minimums to maximums for cells A2 to L11.

### 4.3. Inserting Data Directly from a Database ###

You can also auto-populate data directly from a database with ease. Here’s a quick example assuming the database connections are properly configured:

```cs
// Database objects for data retrieval
string connectionString;
string query;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Establish the connection string
connectionString = @"Data Source=Server;Initial Catalog=Database;User ID=Username;Password=Password";

// SQL query to fetch data
query = "SELECT * FROM MyTable";

// Opening connection and filling dataset
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(query, sqlConnection);
sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Iterate and populate cells from dataset
foreach (DataTable table in dataSet.Tables)
{
    for (int rowIndex = 12, rowCount = table.Rows.Count - 1; rowIndex <= 21; rowIndex++, rowCount--)
    {
        workSheet["A" + rowIndex].Value = table.Rows[rowCount]["Column1"].ToString();
        workSheet["B" + rowIndex].Value = table.Rows[rowCount]["Column2"].ToString();
        workSheet["C" + rowIndex].Value = table.Rows[rowCount]["Column3"].ToString();
        workSheet["D" + rowIndex].Value = table.Rows[rowCount]["Column4"].ToString();
        workSheet["E" + rowIndex].Value = table.Rows[rowCount]["Column5"].ToString();
        workSheet["F" + rowIndex].Value = table.Rows[rowCount]["Column6"].ToString();
        workSheet["G" + rowIndex].Value = table.Rows[rowCount]["Column7"].ToString();
        workSheet["H" + rowIndex].Value = table.Rows[rowCount]["Column8"].ToString();
        workSheet["I" + rowIndex].Value = table.Rows[rowCount]["Column9"].ToString();
        workSheet["J" + rowIndex].Value = table.Rows[rowCount]["Column10"].ToString();
        workSheet["K" + rowIndex].Value = table.Rows[rowCount]["Column11"].ToString();
        workSheet["L" + rowIndex].Value = table.Rows[rowCount]["Column12"].ToString();
    }
}
```

### 4.1. Manually Assigning Values to Cells ###

When you need to input values into cells individually, you select the specific cell by its reference and assign the desired value directly. The following sample illustrates this straightforward process: 

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

Here, individual months are assigned to cells from A1 to L1, populating each cell in the first row with a different month name.

Here is a paraphrased version of the given section:

```cs
// Assigning the names of months to the first row, from column A to L
workSheet["A1"].Value = "Jan";
workSheet["B1"].Value = "Feb";
workSheet["C1"].Value = "Mar";
workSheet["D1"].Value = "Apr";
workSheet["E1"].Value = "May";
workSheet["F1"].Value = "Jun";
workSheet["G1"].Value = "Jul";
workSheet["H1"].Value = "Aug";
workSheet["I1"].Value = "Sep";
workSheet["J1"].Value = "Oct";
workSheet["K1"].Value = "Nov";
workSheet["L1"].Value = "Dec";
```

This snippet places the abbreviated names of the months into the corresponding cells from column A to L in the first row of the worksheet.

In this example, I have filled each column from A through L and assigned the first row of each column with the names of different months.

### 4.2. Dynamically Assigning Cell Values ###

The process of dynamically assigning values to cells follows closely to the previously mentioned method, but it offers the benefit of not requiring the specification of exact cell addresses. The upcoming example demonstrates how to instantiate a new `Random` object for generating random numbers. Afterwards, a `for` loop is utilized to traverse through a specified range of cells to fill them dynamically with these values.

Here's a paraphrased version of the provided C# code snippet:

```cs
// Initialize a random number generator
Random randomGenerator = new Random();

// Populate cells from A2 to L11 with random values between specified ranges
for (int i = 2; i <= 11; i++)
{
    workSheet[$"A{i}"].Value = randomGenerator.Next(1, 1000);      // Values for column A
    workSheet[$"B{i}"].Value = randomGenerator.Next(1000, 2000);   // Values for column B
    workSheet[$"C{i}"].Value = randomGenerator.Next(2000, 3000);   // Values for column C
    workSheet[$"D{i}"].Value = randomGenerator.Next(3000, 4000);   // Values for column D
    workSheet[$"E{i}"].Value = randomGenerator.Next(4000, 5000);   // Values for column E
    workSheet[$"F{i}"].Value = randomGenerator.Next(5000, 6000);   // Values for column F
    workSheet[$"G{i}"].Value = randomGenerator.Next(6000, 7000);   // Values for column G
    workSheet[$"H{i}"].Value = randomGenerator.Next(7000, 8000);   // Values for column H
    workSheet[$"I{i}"].Value = randomGenerator.Next(8000, 9000);   // Values for column I
    workSheet[$"J{i}"].Value = randomGenerator.Next(9000, 10000);  // Values for column J
    workSheet[$"K{i}"].Value = randomGenerator.Next(10000, 11000); // Values for column K
    workSheet[$"L{i}"].Value = randomGenerator.Next(11000, 12000); // Values for column L
}
```

The code retains the same functionality but varies the syntax slightly, providing fresh expression while ensuring clarity and readability for developers.

Each cell in the range from A2 to L11 is populated with a distinct, randomly generated value.

Discussing the insertion of dynamic values, let's explore how you can programmatically insert data from a database into cells. The upcoming code example demonstrates this process, provided that your database connections are appropriately established.

### 4.3. Insert Data Straight from a Database ###

Incorporating data directly from a database into your worksheet is straightforward using IronXL. Below you will find a code snippet that demonstrates how to accomplish this, assuming you already have your database connection established.

```cs
// Establish objects for database interaction
string connectionString;
string sqlQuery;
DataSet dataSet = new DataSet("ExampleDataSet");
SqlConnection sqlConnection;
SqlDataAdapter sqlDataAdapter;

// Define your database connection string
connectionString = @"Data Source=Your_Server;Initial Catalog=Your_Database;User ID=Your_Username;Password=Your_Password";

// SQL Query to retrieve data
sqlQuery = "SELECT Columns FROM Your_Table";

// Initialize the connection and fill the dataset
sqlConnection = new SqlConnection(connectionString);
sqlDataAdapter = new SqlDataAdapter(sqlQuery, sqlConnection);
sqlConnection.Open();
sqlDataAdapter.Fill(dataSet);

// Iterate over the dataset to populate the worksheet cells
foreach (DataTable table in dataSet.Tables)
{
    int rowIndex = table.Rows.Count - 1;
    for (int colIndex = 12; colIndex <= 21; colIndex++)
    {
        workSheet["A" + colIndex].Value = table.Rows[rowIndex]["Column1"].ToString();
        workSheet["B" + colIndex].Value = table.Rows[rowIndex]["Column2"].ToString();
        workSheet["C" + colIndex].Value = table.Rows[rowIndex]["Column3"].ToString();
        workSheet["D" + colIndex].Value = table.Rows[rowIndex]["Column4"].ToString();
        workSheet["E" + colIndex].Value = table.Rows[rowIndex]["Column5"].ToString();
        workSheet["F" + colIndex].Value = table.Rows[rowIndex]["Column6"].ToString();
        workSheet["G" + colIndex].Value = table.Rows[rowIndex]["Column7"].ToString();
        workSheet["H" + colIndex].Value = table.Rows[rowIndex]["Column8"].ToString();
        workSheet["I" + colIndex].Value = table.Rows[rowIndex]["Column9"].ToString();
        workSheet["J" + colIndex].Value = table.Rows[rowIndex]["Column10"].ToString();
        workSheet["K" + colIndex].Value = table.Rows[rowIndex]["Column11"].ToString();
        workSheet["L" + colIndex].Value = table.Rows[rowIndex]["Column12"].ToString();
    }
    rowIndex++;
}
```

This example utilizes a `SqlConnection` to fetch data from a specified database and fills a `DataSet`. The `WorkSheet` object then dynamically receives data, populating each cell with information directly from your database fields.

Here is the paraphrased section:

```cs
// Initialize database components to load data from the database
string connectionString;
string query;
DataSet dataSet = new DataSet("MyDataSet");
SqlConnection connection;
SqlDataAdapter adapter;

// Configure the database connection string
connectionString = @"Data Source=Your_Server_Name;Initial Catalog=Your_Database;User ID=Your_Username;Password=Your_Password";

// Define the SQL query to retrieve data
query = "SELECT ColumnNames FROM YourTable";

// Establish Connection and Populate the DataSet
connection = new SqlConnection(connectionString);
adapter = new SqlDataAdapter(query, connection);
connection.Open();
adapter.Fill(dataSet);

// Iterate over the DataSet to populate Worksheet cells
foreach (DataTable dt in dataSet.Tables)
{
    int rowIndex = dt.Rows.Count - 1; // Start from the last row
    for (int i = 12; i <= 21; i++)
    {
        workSheet["A" + i].Value = dt.Rows[rowIndex]["Column1"].ToString();
        workSheet["B" + i].Value = dt.Rows[rowIndex]["Column2"].ToString();
        workSheet["C" + i].Value = dt.Rows[rowIndex]["Column3"].ToString();
        workSheet["D" + i].Value = dt.Rows[rowIndex]["Column4"].ToString();
        workSheet["E" + i].Value = dt.Rows[rowIndex]["Column5"].ToString();
        workSheet["F" + i].Value = dt.Rows[rowIndex]["Column6"].ToString();
        workSheet["G" + i].Value = dt.Rows[rowIndex]["Column7"].ToString();
        workSheet["H" + i].Value = dt.Rows[rowIndex]["Column8"].ToString();
        workSheet["I" + i].Value = dt.Rows[rowIndex]["Column9"].ToString();
        workSheet["J" + i].Value = dt.Rows[rowIndex]["Column10"].ToString();
        workSheet["K" + i].Value = dt.Rows[rowIndex]["Column11"].ToString();
        workSheet["L" + i].Value = dt.Rows[rowIndex]["Column12"].ToString();
    }
    rowIndex++; // Next row for iteration
}
```

All you need to do is assign the Field name to the `Value` property of the specific cell you want to populate.

<hr class="separator">

## 5. Formatting Cells ##

### 5.1. Setting Background Colors ###

To change the background color of individual cells or cell ranges, you can accomplish this with a single line of code:

```cs
workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This example paints the range from A1 to L1 with a gray color. Here, the color is specified using a hexadecimal code, which is a standard color coding in HTML where each pair of digits represents the intensity of red, green, and blue components, respectively, from 00 to FF.

### 5.2. Adding Borders to Cells ###

To define borders around cells using IronXL is straightforward. Below is how you can define borders for various ranges:

```cs
// Set the top and bottom borders of A1 to L1 to black
workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Set a medium-strength right border from L2 to L11
workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Apply a medium bottom border from A11 to L11
workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

In the code snippet above, we define black top and bottom borders for the cells A1 through L1. For cells L2 through L11, a right border is applied, and similarly, a bottom border is defined for A11 through L11, both with a medium thickness. 

These examples show how versatile and easy it is to format cells using IronXL, enhancing both the functionality and aesthetics of your Excel data presentations.

### 5.1. Applying Background Colors to Cells ###

To define the background color for a single cell or a group of cells, use a simple line of code as demonstrated below:
```

```cs
// Apply gray background color to the cells from A1 to L1
workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
```

This code specifies a gray background color for a selected range of cells, utilizing the RGB (Red, Green, Blue) color format. In this notation, the first two characters signify the Red component, the subsequent two denote Green, and the final pair represents Blue. The possible values for each color component range from 0 to 9, followed by A to F, based on the hexadecimal system.

### 5.2. Establishing Cell Borders ###

Easily define borders within your Excel sheets using IronXL. Here’s how you do it:

```cs
// Set a black top border for cells A1 through L1
workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");

// Set a black bottom border for the same range
workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Right border configuration for cells from L2 to L11
workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Establishing a medium bottom border for the range A11 to L11
workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

With this straightforward code, you can apply top, bottom, and right borders of varying styles across specific cell ranges, enhancing both aesthetics and readability of your data presentation.

The provided code snippet demonstrates how to apply border styles to various cells in an Excel worksheet using IronXL:

```cs
// Setting the color of the top and bottom borders for A1 to L1 to black
workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");

// Configuring the right border for cells from L2 to L11, setting the color to black and border type to medium
workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;

// Applying a medium black bottom border from A11 to L11
workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
```

In the example provided, black borders are applied at the top and bottom of cells ranging from A1 to L1. Additionally, I have configured the right border for cells between L2 and L11, choosing a medium style for the border. Lastly, the bottom border has been applied to cells spanning from A11 to L11.

<hr class="separator">

## 6. Use Formulas in Cells ##

IronXL simplifies many programming tasks, and working with formulas in Excel cells is no exception. Here's how easily you can incorporate formulas into your spreadsheets:

```cs
// Use built-in functions of IronXL to perform calculations
decimal sumResult = workSheet["A2:A11"].Sum();
decimal averageResult = workSheet["B2:B11"].Avg();
decimal maximumValue = workSheet["C2:C11"].Max();
decimal minimumValue = workSheet["D2:D11"].Min();

// Assign calculated values to cells
workSheet["A12"].Value = sumResult;
workSheet["B12"].Value = averageResult;
workSheet["C12"].Value = maximumValue;
workSheet["D12"].Value = minimumValue;
```

This example beautifully illustrates the usability of IronXL, enabling you to apply common statistical functions like SUM (to add values), AVG (to calculate average), MAX (to find the highest value), and MIN (to determine the lowest value) directly within your .NET applications.

Here's a paraphrased version of the specified section:

```cs
// Employ IronXL's aggregation functions
decimal total = workSheet["A2:A11"].Sum();  // Calculate total of range A2:A11
decimal average = workSheet["B2:B11"].Avg();  // Compute average for range B2:B11
decimal highest = workSheet["C2:C11"].Max();  // Find maximum value in range C2:C11
decimal lowest = workSheet["D2:D11"].Min();  // Determine minimum value in range D2:D11

// Populate cells with computed values
workSheet["A12"].Value = total;
workSheet["B12"].Value = average;
workSheet["C12"].Value = highest;
workSheet["D12"].Value = lowest;
```

This revised snippet uses different variable names and comments to enhance clarity and maintain the original functions' purposes, demonstrating how to utilize IronXL to perform and apply spreadsheet calculations efficiently.

One of the great features here is the ability to specify the data type for a cell, thereby influencing the outcome of the formula. The previously mentioned block of code demonstrates how to implement basic functions like SUM (to total up values), AVG (to calculate the average), MAX (to find the maximum value), and MIN (to determine the minimum value).

<hr class="separator">

## 7. Configuring Worksheet and Printing Options ##

### 7.1. Customize Worksheet Settings ###

Using IronXL allows you to modify various worksheet properties such as freezing panes or adding password protection to your sheets, detailed below:

```cs
workSheet.ProtectSheet("YourPasswordHere");
workSheet.CreateFreezePane(0, 1);
```

This code freezes the first row and secures the worksheet, preventing unauthorized editing. Figures 7 and 8 visually demonstrate these configurations.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 7</strong> – <em>View of Freeze Panes</em></p>
  </div>
</center>

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 8</strong> – <em>Protected Worksheet Display</em></p>
  </div>
</center>

### 7.2. Adjust Page and Print Configuration ###

Tailor your document's layout and printing specifications effortlessly with IronXL with just a few lines of code:

```cs
workSheet.SetPrintArea("A1:L12");
workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

This configures the print area of A1 to L12, sets the page orientation to landscape, and specifies A4 as the paper size.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Preview of Print Setup</em></p>
  </div>
</center>

These modifications enhance both the usability and security of your Excel files, ensuring they meet both functional and presentation standards.

### 7.1. Adjusting Worksheet Settings ###

Modifying worksheet settings allows you to add security and enhance usability by freezing rows and columns, as well as securing the worksheet with a password protection. The details are as follows:

```cs
// Applying protection to the worksheet with a password
workSheet.ProtectSheet("Password");

// Freezing the top row to keep it visible during scrolling
workSheet.CreateFreezePane(0, 1);
```

The topmost row remains static and won't move vertically with the rest of the sheet content. Additionally, the worksheet is secured against modifications by using a password. This functionality is demonstrated in Figures 7 and 8.

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

### 7.2. Configure Page and Printing Options ###

Adjust various page settings including the layout orientation, page dimensions, and designated print areas, among others.

Here's a paraphrased version of the provided section, with modified code, enhanced code comments, and absolute URL paths resolved:

```cs
// Define the area of the worksheet to be printed
workSheet.SetPrintArea("A1:L12");

// Configure the print orientation to landscape
workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;

// Set the paper size to A4 for printing
workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
```

The printable section of the worksheet is defined as ranging from A1 to L12. This setting aligns the page orientation in Landscape mode and specifies the paper size as A4.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Print Setup</em></p>
  </div>
</center>

<hr class="separator">

```cs
workBook.SaveAs("Budget.xlsx");
```

Here's the paraphrased section of the article:

```cs
// Save the Excel Workbook with a new name
workBook.SaveAs("FinancialPlanning.xlsx");
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

