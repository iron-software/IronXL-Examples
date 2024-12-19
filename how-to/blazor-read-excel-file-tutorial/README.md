# Reading Excel Files in Blazor with IronXL: A Comprehensive Tutorial

***Based on <https://ironsoftware.com/how-to/blazor-read-excel-file-tutorial/>***


## Introduction

Blazor is a .NET Web framework developed by Microsoft, allowing developers to use C# and compile it into browser-compatible JavaScript and HTML. This guide outlines a straightforward approach to read Excel files in a Blazor serverside application using the IronXL library.

![Demonstration of IronXL Viewing Excel in Blazor](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/demo.gif)

### Getting Started with IronXL


----------------------------------

## Step 1 - Create a Blazor Project in Visual Studio

Let's take an XLSX file as an example that we will read and display in a Blazor Server App:

<div>
<table style="margin: 0 auto 0 auto">
<tr>
    <th style="border: 1px solid black ; padding: 5px;">Input XLSX Excel Sheet</th>
    <th style="border: 1px solid black ; padding: 5px;">Result in Blazor Server Browser</th>
</tr>
<tr>
    <td style="border: 1px solid black ; padding: 5px;">
        <table style="border: 1px solid black ; margin: 0 auto 0 auto">
            <tr>
                <th style="border: 2px solid black ; padding: 5px;">First name</th>
                <th style="border: 2px solid black ; padding: 5px;">Last name</th>
                <th style="border: 2px solid black ; padding: 5px;">ID</th>
            </tr>
            <tr>
                <td style="border: 1px solid black ; padding: 5px;">John</td>
                <td style="border: 1px solid black ; padding: 5px;">Applesmith</td>
                <td style="border: 1px solid black ; padding: 5px;">1</td>
            </tr>
            <tr>
                <td style="border: 1px solid black ; padding: 5px;">Richard</td>
                <td style="border: 1px solid black ; padding: 5px;">Smith</td>
                <td style="border: 1px solid black ; padding: 5px;">2</td>
            </tr>
            <tr>
                <td style="border: 1px solid black ; padding: 5px;">Sherry</td>
                <td style="border: 1px solid black ; padding: 5px;">Robins</td>
                <td style="border: 1px solid black ; padding: 5px;">3</td>
            </tr>
        </table>
    </td>
    <td style="border: 1px solid black ; padding: 5px;">
        ![Browser view of the spreadsheet](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/browser-view.webp)
    </td>
</tr>
</table>
</div>

1. Begin by initiating a new Blazor Project from the Visual Studio IDE:
    ![Setting up a new project](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/new-project.webp)

2. Select the **`Blazor Server App`** Project type:
    ![Choosing the project type](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/choose-blazor-project-type.webp)

3. Execute the application with the `F5` key and navigate to the `Fetch data` tab:
    ![First run of the app](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/first-run.webp)

The aim is to incorporate an upload mechanism for Excel files and to display them within the application.

## Step 2 - Incorporate IronXL into Your Solution

### IronXL: .NET Excel Library

IronXL is a robust .NET library that treats Excel spreadsheets as objects. This enables deep integration with C# to manipulate spreadsheet data effectively. IronXL offers comprehensive functionalities over alternatives like NPOI, providing easier complex operations, more licensing options, and better support.

IronXL is compatible with the latest .NET versions (8, 7, 6) and .NET Core Framework 4.6.2+.

To add IronXL to your solution, you can use either of the following methods:

#### Method 1 - NuGet Package Manager:

```shell
Install-Package IronXL.Excel
```

#### Method 2 - csproj File:

Add IronXL directly to your project by including this line in an `<ItemGroup>` section in your solution's `.csproj` file:

```xml
<PackageReference Include="IronXL.Excel" Version="*" />
```

Here is how it appears in Visual Studio:
![Add IronXL to project via csproj](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/add-ironxl-csproj.webp)

## Step 3 - Code the File Upload and Display Functionality

Inside the Visual Studio Solution Explorer, navigate to the `Pages/` directory and open `FetchData.razor`. This file is a good starting point, as it's part of the Blazor Server App template.

Replace the content of `FetchData.razor` with the following code:

```cs
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
            </tr>
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

The `<InputFile>` component aids in uploading files to the webpage, invoking the `OpenExcelFileFromDisk` asynchronous method to process the uploaded Excel file. The HTML renders the Excel sheet as a table on the web tab.

IronXL.Excel stands out in the .NET ecosystem as a versatile library for reading an array of spreadsheet formats without needing Microsoft Excel installed.

<hr class="separator">

### Further Reading

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>Explore the API Reference</h3>
      <p>Dive deeper into the IronXL API, detailing all its features, namespaces, classes, methods, fields, and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

*[Download IronXL for Blazor](https://ironsoftware.com/csharp/excel/how-to/blazor-read-excel-file-tutorial/)*