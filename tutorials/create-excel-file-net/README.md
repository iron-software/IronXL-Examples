# C# Create Excel File Guide

***Based on <https://ironsoftware.com/tutorials/create-excel-file-net/>***


This guide provides a detailed walkthrough on how to generate an Excel Workbook file across any platform compatible with .NET Framework 4.5 or .NET Core. Crafting Excel files in C# is straightforward and doesn't rely on the traditional **Microsoft.Office.Interop.Excel** library. With IronXL, you can easily configure worksheet attributes such as freeze panes and protection, as well as define print settings, among other features.

<hr class="separator">

<h4 class="tutorial-segment-title">Overview</h4>




<h2>IronXL Creates C&num; Excel Files in .NET</h2>

[IronXL is a powerful C# & VB Excel API](https://ironsoftware.com/csharp/excel/) that enables you to read, modify, and generate Excel spreadsheet files in .NET rapidly and efficiently, without the necessity for installing Microsoft Office or Excel Interop.

IronXL offers comprehensive support for .NET Core, .NET Framework, Xamarin, Mobile platforms, Linux, macOS, and Azure.

<h3>IronXL Features:</h3>

- Direct human assistance from our dedicated .NET developer team
- Swift integration using Microsoft Visual Studio
- Complimentary access for development phases. Licensing starts from `$liteLicense`.

<h4>Create and Save an Excel File: Quick Code</h4>

<a class="js-modal-open" href="https://www.nuget.org/packages/IronXL.Excel/" target="_blank" data-modal-id="trial-license-after-download">https://www.nuget.org/packages/IronXL.Excel/</a>

Alternatively, you can [download the IronXL.dll package here](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and include it in your project.

```cs
using IronXL;
using IronXL.Excel;

namespace ironxl.CreateExcelFileNet
{
    public class Section1
    {
        public void Execute()
        {
            // Set the default file format to XLSX, though this can be changed with CreatingOptions
            WorkBook excelWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var excelSheet = excelWorkbook.CreateWorkSheet("example_sheet");
            excelSheet["A1"].Value = "Sample Text"; // Initialize A1 cell with 'Sample Text'

            // Assigning values to a range of cells
            excelSheet["A2:A4"].Value = 5;
            // Set background color for cell A5
            excelSheet["A5"].Style.SetBackgroundColor("#f0f0f0");
            
            // Applying bold font style to cells A5 and A6
            excelSheet["A5:A6"].Style.Font.Bold = true;

            // Incorporate a formula in cell A6 and evaluate
            excelSheet["A6"].Value = "=SUM(A2:A4)";
            if (excelSheet["A6"].IntValue == excelSheet["A2:A4"].IntValue)
            {
                Console.WriteLine("Validation succeeded");
            }
            
            // Save the workbook to a file
            excelWorkbook.SaveAs("example_workbook.xlsx");
        }
    }
}
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## Get Started with the Free IronXL C# Library

### Installing IronXL via NuGet

You have multiple options to incorporate the IronXL NuGet package into your project, conveniently supporting different workflows:

1. **Visual Studio Interface**:
   - Navigate through the Project Menu or directly via the Solution Explorer.
   - Opt to manage the NuGet packages and search for IronXL.Excel to install.

   ![Project Menu Guide](https://ironsoftware.com/img/tutorials/create-excel-file-net/project-menu.png)
   *Figure 3 – Access via Project Menu*

   ![Solution Explorer Guide](https://ironsoftware.com/img/tutorials/create-excel-file-net/right-click-solution-explorer.png)
   *Figure 4 – Right Click in Solution Explorer to Install*

   After selecting ‘Manage NuGet Packages,’ use the browse feature to find `IronXL.Excel` and proceed with the installation as depicted below.

   ![NuGet Package Installation](https://ironsoftware.com/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png)
   *Figure 5 – Installing IronXL.Excel via NuGet*

2. **Developer Command Prompt**:
   - Locate the Developer Command Prompt for your installed version of Visual Studio, usually found within the Visual Studio installation directory.
   - Once open, enter the command: `PM> Install-Package IronXL.Excel` and hit Enter.
   - The package will install, after which you should reload your Visual Studio project.

### Direct Download

Alternatively, you can directly download the library if that aligns better with your setup:

1. Visit [IronXL Direct Download](https://ironsoftware.com/csharp/excel/).
2. Download the library package manually.
3. Add a reference to the downloaded `IronXL.dll` in your Visual Studio project by right-clicking on ‘References’ in Solution Explorer and browsing for the downloaded file.

Now that IronXL is integrated into your project, you can explore its vast capabilities for manipulating Excel files in .NET environments!

### Ready to Proceed?

With IronXL installed, you are now equipped to dive into creating and managing Excel files in C# with ease. Let’s start leveraging IronXL’s robust features to make your data handling tasks more efficient!

<h3>Install by Using NuGet</h3>

There are three distinct methods to integrate the IronXL NuGet package into your project:

1. **Visual Studio**: Accessible through the Project Menu or directly by right-clicking the project in the Solution Explorer.

2. **Developer Command Prompt**: Start by locating your Developer Command Prompt within the Visual Studio installation directory. Then, type the command `PM > Install-Package IronXL.Excel` and hit Enter. Once complete, ensure you refresh your project in Visual Studio.

3. **Direct NuGet Package Download**: Navigate to [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/) and download the package. After download, execute the package and revitalize your Visual Studio project to reflect the changes.

<h3>Visual Studio</h3>

Visual Studio is equipped with a NuGet Package Manager that allows you to integrate NuGet packages into your projects. Access this feature through the Project Menu or by right-clicking your project within the Solution Explorer. These methods are illustrated in Figures 3 and 4 below.

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

Once you've selected the Manage NuGet Packages option from either of the approaches, search for the `IronXL.Excel` package and proceed to install it, as demonstrated in Figure 5.

<br>
<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/install-iron-excel-nuget-package.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 5</strong> – <em>Install IronXL.Excel NuGet Package</em></p>
  </div>
</center>

<h3>Developer Command Prompt</h3>

Begin by launching the Developer Command Prompt for your particular development setup. You'll typically find this tool in your Visual Studio installation directory. Once open, perform these actions to integrate the IronXL.Excel NuGet package into your project:

1. Locate the Developer Command Prompt, usually inside your Visual Studio install folder.
2. Input the following command: `PM > Install-Package IronXL.Excel`
3. Hit the Enter key to run the command.
4. Upon pressing Enter, the IronXL.Excel package will commence installation.
5. After installation, make sure to refresh or reload your Visual Studio project to ensure the changes are applied successfully.

<h3>Download the NuGet Package directly</h3>

Follow these steps to download and install the NuGet package:

1. Visit the URL: [https://www.nuget.org/packages/ironxl.excel/](https://www.nuget.org/packages/ironxl.excel/)

2. Select "Download Package."

3. Once the download is complete, double-click the downloaded file.

4. Reopen your Visual Studio project to apply the changes.

<h3>Install IronXL by Direct Download of the Library</h3>

The alternative method for installing IronXL involves a direct download from the following URL: [https://ironsoftware.com/csharp/excel/](https://ironsoftware.com/csharp/excel/).

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/download-ironxl-library.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/download-ironxl-library.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 6</strong> – <em>Download IronXL library</em></p>
  </div>
</center>

To integrate the IronXL library into your project, simply follow these steps:

1. In the Solution Explorer, right-click on the Solution.
2. Choose 'References' from the context menu.
3. Search for and select the `IronXL.dll` library.
4. Confirm your selection by clicking 'OK'.

<h3>Let's Go!</h3>

Now that everything is configured, let's dive into the powerful capabilities of the IronXL library!

<hr class="separator">

<h4 class="tutorial-segment-title">How to Tutorials</h4>

## 2. Constructing an ASP.NET Project ##

Begin your ASP.NET project by following these steps:

1. Visit the IronXL package page at [NuGet](https://www.nuget.org/packages/ironxl.excel/) and download the package.
2. Once downloaded, double-click the file to initiate installation.
3. Reopen your Visual Studio project to complete the setup.

Next, create an ASP.NET website:

1. Launch Visual Studio.
2. Navigate to `File > New Project`.
3. In the list of project types, choose 'Web' under Visual C# options.
4. Select `ASP.NET Web Application` as shown below:

   ![New Project](https://ironsoftware.com/img/tutorials/create-excel-file-net/new-project-asp-net.png)
   **Figure 1** – *Starting a New Project*

5. Click `OK`.
6. You'll then see a new screen where you should select `Web Forms`. Refer to the illustration below for guidance:

   ![Web Forms](https://ironsoftware.com/img/tutorials/create-excel-file-net/web-form.png)
   **Figure 2** – *Selecting Web Forms*
   
7. Confirm your choice by clicking `OK`.

With this setup complete, you're ready to incorporate IronXL and start enhancing your file with its extensive features.

<ul>
  <li>Navigate to the following URL:</li>
  <li><a href="https://www.nuget.org/packages/ironxl.excel/" target="_blank">https://www.nuget.org/packages/ironxl.excel/</a></li>
  <li>Click on Download Package</li>
  <li>After the package has downloaded, double click it</li>
  <li>Reload your Visual Studio project</li>
</ul>

Follow these steps to initiate an ASP.NET website:

1. Launch Visual Studio.
2. Navigate to `File > New Project`.
3. Choose `Web` from the `Project type` dropdown under Visual C#.
4. Opt for `ASP.NET Web Application` as depicted below.

<br>
<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/new-project-asp-net.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/new-project-asp-net.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 1</strong> — *New Project*

5. Click on the OK button.

6. On the subsequent screen, choose Web Forms as depicted in Figure 2 below.
```

<br>

<center>
<a rel="nofollow" href="/img/tutorials/create-excel-file-net/web-form.png" target="_blank"><p><img src="/img/tutorials/create-excel-file-net/web-form.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></p></a>
</center>

<strong style="margin-left: 40px;">Figure 2</strong> – *Web Forms*
```

<br>

After clicking OK, you've laid the groundwork. Now, you can proceed with installing IronXL to begin tailoring your Excel file to your needs.

<hr class="separator">

Creating an Excel Workbook with IronXL couldn't be easier—it's literally just one line of code! Here's how:
```
// Simplified way to instantiate a new Excel Workbook
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
```

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section2
    {
        public void Run()
        {
            // Initiate a new workbook in XLSX format
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}
```

IronXL provides the flexibility to generate both the traditional XLS format as well as the modern XLSX format for Excel files.

```cs
using IronXL.Excel;

namespace ironxl.CreateExcelFileNet
{
    public class Section3
    {
        public void Run()
        {
            // Create a new worksheet named "2020 Budget"
            WorkSheet workSheet = workBook.CreateWorkSheet("2020 Budget");
        }
    }
}
```

"Sheet" in the provided code snippet refers to the newly created worksheet. You can apply all typical Excel operations on this "Sheet" such as setting cell values and much more.

To clarify further, a Workbook can house multiple Worksheets. Therefore, you can include as many Worksheets as you need within a single Workbook. In another section of this tutorial, the processes for adding multiple Worksheets will be detailed. For now, understand that a Worksheet is comprised of Rows and Columns, with their intersection forming a Cell. This Cell is the element you will work on while managing Excel files.

Here is a paraphrased version of the given section with improved explanations, and proper paths resolved:

```cs
using IronXL.Excel;

// Define the namespace for creating an Excel file
namespace ironxl.CreateExcelFileNet
{
    public class BudgetWorksheetCreator
    {
        public void Execute()
        {
            // Create a new worksheet within the workbook with the title "2020 Budget"
            WorkSheet budgetSheet = workBook.CreateWorkSheet("2020 Budget");
            // The 'budgetSheet' is now ready to be modified and populated with data
        }
    }
}
```

"Sheet" in the provided code sample refers to the worksheet, which affords you the capability to set cell values and perform nearly all actions available in Excel.

To clarify any possible confusion, let’s differentiate between a Workbook and a Worksheet:

A Workbook is essentially a container that holds one or more Worksheets. You have the option to include numerous Worksheets within a single Workbook. More details on this will be covered in a subsequent article. Each Worksheet is made up of Rows and Columns, with each intersecting point forming what is known as a Cell. It is these Cells that you interact with when manipulating data in Excel.

<hr class="separator">

## 4. Assigning Values to Cells ##

### 4.1. Manually Inputting Values ###

To input values into cells manually, simply identify the target cell and assign it a value, as demonstrated in the example below:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section4
    {
        public void Run()
        {
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
        }
    }
}
```
Here, Columns A through L for row one have been populated with the names of the months.

### 4.2. Dynamically Setting Cell Values ###

Dynamically setting cell values is similar to manually inputting them, with the addition of not needing to hard-code cell positions. In the following example, you'll use a `Random` object to generate random numbers and fill a range of cells:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section5
    {
        public void Run()
        {
            Random random = new Random();
            for (int i = 2 ; i <= 11 ; i++)
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
        }
    }
}
```

Each cell from A2 to L11 will contain a unique, randomly generated value.

### 4.3. Inputting Data from a Database Directly ###

To populate cells directly from a database, assuming your database connections are correctly configured, you can use the following code snippet:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section6
    {
        public void Run()
        {
            // Database objects setup
            string connectionString;
            string query;
            DataSet dataset = new DataSet("ExampleDataSet");
            SqlConnection connection;
            SqlDataAdapter adapter;
            
            // Define Database Connection
            connectionString = @"Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserID;Password=Password";
            
            // SQL Query
            query = "SELECT Fields FROM TableName";
            
            // Open Connection & Populate DataSet
            connection = new SqlConnection(connectionString);
            adapter = new SqlDataAdapter(query, connection);
            connection.Open();
            adapter.Fill(dataset);
            
            // Iterate through DataSet content
            foreach (DataTable table in dataset.Tables)
            {
                int rowCount = table.Rows.Count - 1;
                for (int j = 12; j <= 21; j++)
                {
                    workSheet["A" + j].Value = table.Rows[rowCount]["ColumnName1"].ToString();
                    workSheet["B" + j].Value = table.Rows[rowCount]["ColumnName2"].ToString();
                    workSheet["C" + j].Value = table.Rows[rowCount]["ColumnName3"].ToString();
                    workSheet["D" + j].Value = table.Rows[rowCount]["ColumnName4"].ToString();
                    workSheet["E" + j].Value = table.Rows[rowCount]["ColumnName5"].ToString();
                    workSheet["F" + j].Value = table.Rows[rowCount]["ColumnName6"].ToString();
                    workSheet["G" + j].Value = table.Rows[rowCount]["ColumnName7"].ToString();
                    workSheet["H" + j].Value = table.Rows[rowCount]["ColumnName8"].ToString();
                    workSheet["I" + j].Value = table.Rows[rowCount]["ColumnName9"].ToString();
                    workSheet["J" + j].Value = table.Rows[rowCount]["ColumnName10"].ToString();
                    workSheet["K" + j].Value = table.Rows[rowCount]["ColumnName11"].ToString();
                    workSheet["L" + j].Value = table.Rows[rowCount]["ColumnName12"].ToString();
                }
            }
        }
    }
}
```

This example demonstrates how to connect to a database, execute a query, and use the results to populate Excel cells dynamically.

### 4.1. Manually Entering Cell Data ###

Directly specifying the content of cells is straightforward; you just select the cell in question and assign the value you want, as demonstrated below:
```

Here's a paraphrased version of the provided C# code snippet that manually sets cell values in an Excel workbook using IronXL:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class MonthSetter
    {
        public void Execute()
        {
            // Assign month names to cells A1 through L1
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
        }
    }
}
```

This code sets each month of the year to corresponding cells from "A1" to "L1" in an Excel worksheet using the IronXL library.

In this example, I have filled columns A through L, with each cell in the first row assigned the name of a different month.

### 4.2. Dynamic Cell Value Assignment ###

Assigning values dynamically streamlines the process by avoiding the need to specify fixed cell locations directly. Below is an example where you will initialize a new `Random` instance for generating random numbers. Subsequently, a `for` loop is used to fill a range of cells with these values dynamically.

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section5
    {
        public void Execute()
        {
            // Create a random number generator
            Random randomGenerator = new Random();
            
            // Dynamically assign random values to cells in columns from A to L
            for (int row = 2; row <= 11; row++)
            {
                workSheet[$"A{row}"].Value = randomGenerator.Next(1, 1000);
                workSheet[$"B{row}"].Value = randomGenerator.Next(1000, 2000);
                workSheet[$"C{row}"].Value = randomGenerator.Next(2000, 3000);
                workSheet[$"D{row}"].Value = randomGenerator.Next(3000, 4000);
                workSheet[$"E{row}"].Value = randomGenerator.Next(4000, 5000);
                workSheet[$"F{row}"].Value = randomGenerator.Next(5000, 6000);
                workSheet[$"G{row}"].Value = randomGenerator.Next(6000, 7000);
                workSheet[$"H{row}"].Value = randomGenerator.Next(7000, 8000);
                workSheet[$"I{row}"].Value = randomGenerator.Next(8000, 9000);
                workSheet[$"J{row}"].Value = randomGenerator.Next(9000, 10000);
                workSheet[$"K{row}"].Value = randomGenerator.Next(10000, 11000);
                workSheet[$"L{row}"].Value = randomGenerator.Next(11000, 12000);
            }
        }
    }
}
```

Every cell between A2 and L11 is populated with a randomly generated, distinct value.

On the subject of dynamic data, let's explore how to populate cells directly from a database. The following code snippet illustrates this process, provided your database connections are properly configured.

### 4.3. Inserting Data from a Database ###

Using IronXL, you can easily populate your Excel worksheet with data from a database. This process leverages a simple connection setup and SQL queries. Below, we demonstrate how to seamlessly retrieve data from a database and write it into an Excel sheet.

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section6
    {
        public void Run()
        {
            // Initialize database connection objects
            string connectionString;
            string query;
            DataSet dataSet = new DataSet("MyDataSet");
            SqlConnection connection;
            SqlDataAdapter dataAdapter;
            
            // Define the connection string
            connectionString = "Data Source=Your_Server;Initial Catalog=Your_Database;User ID=Your_UserID;Password=Your_Password";
            
            // Specify the SQL query for data retrieval
            query = "SELECT Columns FROM Your_Table";
            
            // Set up and execute the connection and command
            connection = new SqlConnection(connectionString);
            dataAdapter = new SqlDataAdapter(query, connection);
            connection.Open();
            dataAdapter.Fill(dataSet);
            
            // Iterate through the dataset and assign values to the worksheet
            foreach (DataTable table in dataSet.Tables)
            {
                int rowCounter = table.Rows.Count - 1;
                for (int i = 12; i <= 21; i++)
                {
                    workSheet["A" + i].Value = table.Rows[rowCounter]["Column1"].ToString();
                    workSheet["B" + i].Value = table.Rows[rowCounter]["Column2"].ToString();
                    workSheet["C" + i].Value = table.Rows[rowCounter]["Column3"].ToString();
                    workSheet["D" + i].Value = table.Rows[rowCounter]["Column4"].ToString();
                    workSheet["E" + i].Value = table.Rows[rowCounter]["Column5"].ToString();
                    workSheet["F" + i].Value = table.Rows[rowCounter]["Column6"].ToString();
                    workSheet["G" + i].Value = table.Rows[rowCounter]["Column7"].ToString();
                    workSheet["H" + i].Value = table.Rows[rowCounter]["Column8"].ToString();
                    workSheet["I" + i].Value = table.Rows[rowCounter]["Column9"].ToString();
                    workSheet["J" + i].Value = table.Rows[rowCounter]["Column10"].ToString();
                    workSheet["K" + i].Value = table.Rows[rowCounter]["Column11"].ToString();
                    workSheet["L" + i].Value = table.Rows[rowCounter]["Column12"].ToString();
                }
                rowCounter++;
            }
        }
    }
}
```

The above example uses a straightforward approach to fetch data from a database and write it into various cell positions in an Excel document using IronXL. This method ensures your data-driven documents are always current and accurate, directly reflecting your database status.

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section6
    {
        public void Run()
        {
            // Initialize database connection objects to load data
            string connectionString;
            string query;
            DataSet data = new DataSet("DataFromDB");
            SqlConnection connection;
            SqlDataAdapter adapter;

            // Configure the connection string for the database
            connectionString = @"Data Source=Server_Name;Initial Catalog=Database_Name;User ID=User_ID;Password=Password";

            // Prepare the SQL query for data retrieval
            query = "SELECT Field_Names FROM Table_Name";

            // Establish database connection and fill dataset
            connection = new SqlConnection(connectionString);
            adapter = new SqlDataAdapter(query, connection);
            connection.Open();
            adapter.Fill(data);

            // Iterate through dataset tables and populate worksheet cells with data
            foreach (DataTable dt in data.Tables)
            {
                int index = dt.Rows.Count - 1;
                for (int rowIndex = 12; rowIndex <= 21; rowIndex++)
                {
                    workSheet["A" + rowIndex].Value = dt.Rows[index]["Field_Name_1"].ToString();
                    workSheet["B" + rowIndex].Value = dt.Rows[index]["Field_Name_2"].ToString();
                    workSheet["C" + rowIndex].Value = dt.Rows[index]["Field_Name_3"].ToString();
                    workSheet["D" + rowIndex].Value = dt.Rows[index]["Field_Name_4"].ToString();
                    workSheet["E" + rowIndex].Value = dt.Rows[index]["Field_Name_5"].ToString();
                    workSheet["F" + rowIndex].Value = dt.Rows[index]["Field_Name_6"].ToString();
                    workSheet["G" + rowIndex].Value = dt.Rows[index]["Field_Name_7"].ToString();
                    workSheet["H" + rowIndex].Value = dt.Rows[index]["Field_Name_8"].ToString();
                    workSheet["I" + rowIndex].Value = dt.Rows[index]["Field_Name_9"].ToString();
                    workSheet["J" + rowIndex].Value = dt.Rows[index]["Field_Name_10"].ToString();
                    workSheet["K" + rowIndex].Value = dt.Rows[index]["Field_Name_11"].ToString();
                    workSheet["L" + rowIndex].Value = dt.Rows[index]["Field_Name_12"].ToString();
                }
                index++;
            }
        }
    }
}
```

All that's required is to assign the Field name to the Value property of the specific cell.

<hr class="separator">

### 5. Apply Formatting ###

Applying formatting in IronXL is straightforward and versatile, allowing for customization of cell appearances in numerous ways.

#### 5.1 Set Cell Background Colors ####

To designate a color for the background of either a single cell or a group of cells, use the following code:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section7
    {
        public void Run()
        {
            // Apply a gray background color (#d3d3d3) to cells ranging from A1 to L1.
            workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}
```
This example sets a gray color for the background across the cells from A1 to L1. The color value used is a hexadecimal code representing shades of red, green, and blue.

#### 5.2 Configure Cell Borders ####

Enhance the visual structure of your spreadsheet by adding borders to specific cells as shown below:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section8
    {
        public void Run()
        {
            // Set a black (#000000) colored border for the top and bottom of cells A1 to L1.
            workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
            workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");
            
            // Define a right-side border for cells L2 to L11 and set its style to medium.
            workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
            workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
            
            // Set a medium black bottom border for the range from A11 to L11.
            workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
            workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
        }
    }
}
```

The code provided adds top and bottom black borders to cells from A1 to L1, and a right border to cells from L2 to L11 with a medium thickness. Additionally, it sets a medium bottom border to the cells from A11 to L11. Each border color is defined by a hexadecimal color code (`#000000` for black).

### 5.1. Assign Background Colors to Cells ###

Adjusting the background color for a single cell or multiple cells in a spreadsheet is straightforward with just one line of code, as demonstrated below:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section7
    {
        public void Run()
        {
            workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}
```

This snippet will alter the background to a light gray shade for the specified range of cells. The RGB color code "#d3d3d3" specifies the color intensity for red, green, and blue components respectively.

Here's the paraphrased section of the article, with relative URL paths resolved to `ironsoftware.com`:

```cs
using IronXL.Excel;

// Declaring a namespace specific to our Excel file creation task
namespace ironxl.CreateExcelFileNet
{
    public class Section7
    {
        public void Run()
        {
            // Setting a gray background color for the first row across columns A to L
            workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}
```

This code assigns a gray background color to a specific range of cells. The color specification is in RGB format—Red, Green, Blue—where the first two hexadecimal digits represent the red component, the middle two the green component, and the last two the blue component, with each digit ranging from '0' to '9' and then 'A' to 'F'.

### 5.2 Set Cell Borders ###

Setting up cell borders in IronXL is an effortless process. The example below demonstrates how to do it:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section8
    {
        public void Run()
        {
            // Apply a top and bottom black border to cells A1 to L1
            workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
            workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");
            
            // Apply a right black border of medium thickness to cells from L2 to L11
            workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
            workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
            
            // Apply a medium thickness bottom border to cells from A11 to L11
            workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
            workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
        }
    }
}
```

In this code snippet, black borders are set for certain ranges of cells both at the top and bottom of the range "A1 to L1", and on the right for the range "L2 to L11". Furthermore, a medium-thick border is applied at the bottom of the range "A11 to L11" to ensure clear separation of data sections.

Here's the paraphrased content for the specified section:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section8
    {
        public void Run()
        {
            // Set black color for the top and bottom borders of the first row
            workSheet["A1:L1"].Style.TopBorder.SetColor("#000000"); // Black color
            workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000"); // Black color
            
            // Apply black color and set the border type to medium for the right border from row 2 to 11 in column L
            workSheet["L2:L11"].Style.RightBorder.SetColor("#000000"); // Black color
            workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
            
            // Set a medium black bottom border for the last row (row 11) across columns A to L
            workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000"); // Black color
            workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
        }
    }
}
```

In this version, I have added comments to clarify each line's purpose, enhancing readability and maintaining code integrity while adjusting the structure and phrasing.

In the provided code snippet, I have configured both the top and bottom borders for cells A1 through L1 to be black. Additionally, I applied a medium-strength right border to the range from L2 to L11. Lastly, the bottom border was also applied to the range from A11 to L11.

<hr class="separator">

## 6. Applying Formulas with IronXL ##

The simplicity of IronXL cannot be overstated, and it's worth mentioning repeatedly. Below, find out how effortlessly you can incorporate formulas into your spreadsheets:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section9
    {
        public void Run()
        {
            // Use IronXL's built-in methods to calculate aggregates
            decimal sum = workSheet["A2:A11"].Sum();
            decimal avg = workSheet["B2:B11"].Avg();
            decimal max = workSheet["C2:C11"].Max();
            decimal min = workSheet["D2:D11"].Min();
            
            // Populate cells with the calculated values
            workSheet["A12"].Value = sum;
            workSheet["B12"].Value = avg;
            workSheet["C12"].Value = max;
            workSheet["D12"].Value = min;
        }
    }
}
```

This example illustrates how seamlessly one can apply various Excel formulas such as SUM, AVG, MAX, and MIN using IronXL, enriching your spreadsheets with dynamic calculations.
```

Here's the paraphrased section of the article, with resolved relative URL paths:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section9
    {
        public void Run()
        {
            // Utilize IronXL's predefined aggregation functions
            decimal total = workSheet["A2:A11"].Sum();
            decimal average = workSheet["B2:B11"].Avg();
            decimal highest = workSheet["C2:C11"].Max();
            decimal lowest = workSheet["D2:D11"].Min();

            // Populate cells with computed values
            workSheet["A12"].Value = total;
            workSheet["B12"].Value = average;
            workSheet["C12"].Value = highest;
            workSheet["D12"].Value = lowest;
        }
    }
}
```

A convenient feature of this procedure is your ability to determine the cell's data type, thereby influencing the outcome of the computations. The provided example demonstrates the utilization of various formulas such as SUM (to calculate the total), AVG (to compute the average), MAX (to find the maximum value), and MIN (to ascertain the minimum value).

<hr class="separator">

## 7. Configuring Sheet and Printing Options ##

### 7.1. Worksheet Configuration ###

Enhancing worksheet functionality involves two primary settings: securing it with a password and setting up frozen panes. Here's how to do it:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section10
    {
        public void Run()
        {
            workSheet.ProtectSheet("Password");  // Protect worksheet with a password
            workSheet.CreateFreezePane(0, 1);    // Freeze the first row
        }
    }
}
```

This implementation ensures the first row stays visible while you scroll through the rest of your data. It also secures your worksheet from unauthorized modifications.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/freeze-panes.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 7</strong> – <em>Illustration of Frozen Panes</em></p>
  </div>
</center>

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/protected-worksheet.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 8</strong> – <em>Protected Worksheet</em></p>
  </div>
</center>

### 7.2. Adjusting Page and Print Settings ###

You can tailor page settings such as orientation, paper size, and the print area with ease:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section11
    {
        public void Run()
        {
            workSheet.SetPrintArea("A1:L12");
            workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape; // Set orientation to landscape
            workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4; // Set paper size to A4
        }
    }
}
```

This code snippet defines the print area along with setting the page to landscape orientation and specifying the paper size as A4.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="https://ironsoftware.com/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Print Setup Screen</em></p>
  </div>
</center>

The adjustments made to the worksheet and print properties enhance the usability and accessibility of your Excel documents, making it simpler to manage and share professional-looking reports.

### 7.1. Configuring Worksheet Settings ###

You can enhance the functionality of your worksheet by employing features like row and column freezing and sheet protection with a password. Here's how to do it:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section10
    {
        public void Run()
        {
            // Protect the worksheet with a password
            workSheet.ProtectSheet("Password");
            // Freeze the first row
            workSheet.CreateFreezePane(0, 1);
        }
    }
}
```

These settings ensure that the first row remains static and visible while you scroll through the worksheet, and the content is secured with a specified password.

Here's the paraphrased version of the provided C# section, with relative URL paths resolved against ironsoftware.com:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class WorksheetProtectionExample
    {
        public void Execute()
        {
            // Protect the worksheet with a password
            workSheet.ProtectSheet("Password");

            // Freeze the top row of the worksheet for easy navigation
            workSheet.CreateFreezePane(0, 1);
        }
    }
}
```

The initial row is set to remain static and will not move when you scroll through the rest of the Worksheet. Additionally, the worksheet is secured, preventing any unauthorized edits by requiring a password. This functionality is demonstrated in Figures 7 and 8.

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

### 7.2. Configuring Page and Print Settings ###

It's possible to adjust various page properties, including its orientation, dimensions, and the designated print area, among others.

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section11
    {
        public void Run()
        {
            // Define the area of the sheet to be included in print
            workSheet.SetPrintArea("A1:L12");
            
            // Set the print orientation to landscape mode
            workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
            
            // Specify the size of the print paper
            workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
        }
    }
}
```

The print area is configured to cover from `A1` to `L12`. The orientation of the page is adjusted to landscape, while the size of the paper is established as A4.

<center>
  <div style="display: inline-block; text-align: left;">
    <a rel="nofollow" href="/img/tutorials/create-excel-file-net/print-setup.png" target="_blank"><img src="/img/tutorials/create-excel-file-net/print-setup.png" alt="" class="img-responsive add-shadow img-margin" style="max-width:100%; margin: 0;"></a>
    <p><strong>Figure 9</strong> – <em>Print Setup</em></p>
  </div>
</center>

<hr class="separator">

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section12
    {
        public void Run()
        {
            workBook.SaveAs("Budget.xlsx");
        }
    }
}
```

To persist the Workbook to a file, implement the following code snippet:

```cs
using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class SaveWorkbook
    {
        public void Execute()
        {
            // Save the workbook with a specific file name
            workBook.SaveAs("Budget_Planner.xlsx");
        }
    }
}
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

