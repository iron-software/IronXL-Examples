# Blazor Read Excel File in C# Using IronXL (Example Tutorial)

***Based on <https://ironsoftware.com/get-started/blazor-read-excel-file-tutorial/>***


## Introduction

Blazor, a .NET Web framework developed by Microsoft, allows applications to run C# code on the web by compiling it into JavaScript and HTML that the browser can execute. This tutorial will guide you through an effective method for reading Excel files within a Blazor server-side application using the IronXL C# library.

![Demonstration of IronXL Viewing Excel in Blazor](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/demo.gif "IronXL Excel Display in Blazor")

### Start Using IronXL
!!!--LIBRARY_START_TRIAL_BLOCK--!!!

----------------------------------

## Step 1 - Create a Blazor Project in Visual Studio

Below is the data from an XLSX file that we'll load and display using a Blazor Server Application:

<div>
<table style="margin: 0 auto;">
<tr>
    <th style="border: 1px solid black; padding: 5px;">Input XLSX Excel Sheet</th>
    <th style="border: 1px solid black; padding: 5px;">Result in Blazor Server Browser</th>
</tr>
<tr>
    <td style="border: 1px solid black; padding: 5px;">
        <table style="border: 1px solid black; margin: 0 auto;">
            <tr>
                <th style="border: 2px solid black; padding: 5px;">First name</th>
                <th style="border: 2px solid black; padding: 5px;">Last name</th>
                <th style="border: 2px solid black; padding: 5px;">ID</th>
            </tr>
            <tr>
                <td>John</td>
                <td>Applesmith</td>
                <td>1</td>
            </tr>
            <tr>
                <td>Richard</td>
                <td>Smith</td>
                <td>2</td>
            </tr>
            <tr>
                <td>Sherry</td>
                <td>Robins</td>
                <td>3</td>
            </tr>
        </table>
    </td>
    <td>
        ![Browser view of Excel data](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/browser-view.webp)
    </td>
</tr>
</table>
</div>

Begin by initiating a new Blazor Project in Visual Studio:

![New Project Setup](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/new-project.webp)

Select **Blazor Server App** as the project type:

![Choosing Blazor Project Type](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/choose-blazor-project-type.webp)

Run the application using the `F5` key and navigate to the `Fetch data` tab in the application.

Our objective is to incorporate an Excel file uploading functionality and subsequently display the Excel content on this web page.

## Step 2 - Integrate IronXL into Your Solution

### IronXL: .NET Excel Library (Installation Steps):

IronXL is a robust .NET library that enables developers to interact with Excel spreadsheets as if they are objects. This feature-rich library facilitates easy manipulation of rows, columns, and cell data in Excel files, making it superior to NPOI in both functionality and user support.

IronXL is compatible with the latest .NET (8, 7, and 6) and .NET Core 4.6.2+ versions.

To add IronXL to your solution, use one of the following methods and then compile your project.

#### Option 2A - Via NuGet Package Manager

```shell
Install-Package IronXL.Excel
```

#### Option 2B - Direct PackageReference in csproj File

Include IronXL in your project by adding the following line to any `<ItemGroup>` in your solution's `.csproj` file:

```xml
<PackageReference Include="IronXL.Excel" Version="*" />
```

As demonstrated in Visual Studio:

![Adding IronXL to csproj](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/add-ironxl-csproj.webp)

## Step 3 - Code the File Upload and Display Functionality

Navigate to the `Pages/` folder in your Visual Studio Solution Explorer and open the `FetchData.razor` file, or any other suitable Razor file provided by the Blazor Server App Template.

Replace the file contents with the following implementation:

```csharp
@using IronXL;
@using System.Data;

@page "/fetchdata"

<PageTitle>Excel File Viewer</PageTitle>

<h1>Open Excel File to View</h1>

<InputFile OnChange="@OpenExcelFileFromDisk" />

<table>
    <thead>
        <tr>
            @foreach (DataColumn column in displayDataTable.Columns)
            {
                <th>@column.ColumnName</th>
            }
        </tr>
    </thead>
    <tbody>
        @foreach (DataRow row in displayDataTable.Rows)
        {
            <tr>
                @foreach (DataColumn column in displayDataTable.Columns)
                {
                    <td>@row[column.ColumnName].ToString()</td>
                }
            }
tr>
        }
    </tbody>
</table>

@code {
    private DataTable displayDataTable = new DataTable();

    async Task OpenExcelFileFromDisk(InputFileChangeEventArgs e)
    {
        IronXL.License.LicenseKey = "PASTE TRIAL OR LICENSE KEY";

        MemoryStream ms = new MemoryStream();
        await e.File.OpenReadStream().CopyToAsync(ms);
        ms.Position = 0;

        WorkBook loadedWorkBook = WorkBook.FromStream(ms);
        WorkSheet loadedWorkSheet = loadedWorkBook.DefaultWorkSheet; // Or use .GetWorkSheet()

        RangeRow headerRow = loadedWorkSheet.GetRow(0);
        for (int col = 0; col < loadedWorkSheet.ColumnCount; col++)
        {
            displayDataTable.Columns.Add(headerRow.ElementAt(col).ToString());
        }

        for (int row = 1; row < loadedWorkSheet.RowCount; row++)
        {
            IEnumerable<string> excelRow = loadedWorkSheet.GetRow(row).ToArray().Select(c => c.ToString());
            displayDataTable.Rows.Add(excelRow.ToArray());
        }
    }
}
```

## Summary

The `<InputFile>` component allows file uploading directly on the web page, invoking `OpenExcelFileFromDisk`. This async callback loads and displays the Excel data into a dynamically created HTML table.

IronXL.Excel operates independently of Microsoft Excel, providing broad support for reading and manipulating Excel files without needing the Excel software installed.

<hr class="separator">

### Further Exploration

<div class="row">
    <div class="col-sm-4">
      <img src="https://ironsoftware.com/img/svgs/documentation.svg" alt="" style="max-width: 110px;"/>
    </div>
    <div class="col-sm-8">
        <h3>Detailed API Reference</h3>
        <p>Explore the comprehensive API Reference for IronXL, detailing all its features, namespaces, classes, methods, fields, and enums.</p>
        <a href="https://ironsoftware.com/csharp/excel/object-reference/api/" class="doc-link" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
    </div>
</div>

*[Download](https://ironsoftware.com/csharp/excel/get-started/blazor-read-excel-file-tutorial/) the IronXL software.*