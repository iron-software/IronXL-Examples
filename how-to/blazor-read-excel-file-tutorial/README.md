# Blazor Read Excel File in C# Using IronXL (Example Tutorial)

## Introduction

Developed by Microsoft, Blazor is an open-source .NET Web framework that enables a C# application to run in the browser as JavaScript and HTML. In this guide, we will delve into a straightforward method for reading Excel files in a server-side Blazor application using the IronXL C# library.

![Demonstration of IronXL Viewing Excel in Blazor](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/demo.gif)

## Step 1 - Create a Blazor Project in Visual Studio

To demonstrate, I will use an XLSX file with the following data, which we will read and display in a Blazor Server application:

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
                <td style="border: 1px solid black; padding: 5px;">John</td>
                <td style="border: 1px solid black; padding: 5px;">Applesmith</td>
                <td style="border: 1px solid black; padding: 5px;">1</td>
            </tr>
            <tr>
                <td style="border: 1px solid black; padding: 5px;">Richard</td>
                <td style="border: 1px solid black; padding: 5px;">Smith</td>
                <td style="border: 1px solid black; padding: 5px;">2</td>
            </tr>
            <tr>
                <td style="border: 1px solid black; padding: 5px;">Sherry</td>
                <td style="border: 1px solid black; padding: 5px;">Robins</td>
                <td style="border: 1px solid black; padding: 5px;">3</td>
            </tr>
        </table>
    </td>
    <td style="border: 1px solid black; padding: 5px;">
        ![Browser View of Data](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/browser-view.webp)
    </td>
</tr>
</table>
</div>

Begin by initiating a Blazor project through Visual Studio:

![Create a New Project](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/new-project.webp)

Select the **`Blazor Server App`** template:

![Choose Blazor Project Type](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/choose-blazor-project-type.webp)

Run the application using the `F5` key and navigate to the `Fetch data` tab:

![First Application Run](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/first-run.webp)

The objective here is to introduce an upload button in the application to load the Excel file, which we will then display.

## Step 2 - Add IronXL to your Solution

### IronXL: .NET Excel Library (Installation Guide):

IronXL is a .NET library that treats Excel spreadsheets as objects, allowing developers to leverage C# and .NET to manipulate and process data. It offers detailed functionalities for extracting cell values, contents, images, references, and formats, presenting advantages over NPOI such as better functionality, easier license management, and excellent support.

IronXL is compatible with the newest versions of .NET (8, 7, and 6) and .NET Core Framework 4.6.2+.

Add IronXL to your solution following one of the methods outlined below:

### Option 2A - Use NuGet Package Manager

### Option 2B - Add PackageReference in the .csproj file

Integrate IronXL directly into your project by inserting the line below into any `<ItemGroup>` in your `.csproj` file:

```xml
<PackageReference Include="IronXL.Excel" Version="*" />
```

Here’s how it looks in Visual Studio:

![Add IronXL to csproj](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/add-ironxl-csproj.webp)

## Step 3 - Coding the File Upload and View

Navigate to the `Pages/` directory and open the `FetchData.razor` file, although any razor file would work. Here's how to replace the content:

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
                <th>
                    @column.ColumnName
                </th>
            }
        </tr>
    </thead>
    <tbody>
        @foreach (DataRow row in displayDataTable.Rows)
        {
            <tr>
                @foreach (DataColumn column in displayDataTable.Columns)
                {
                    <td>
                        @row[column.ColumnName].ToString()
                    </td>
                }
            </tr>
        }
    </tbody>
</table>

@code {
    private DataTable displayDataTable = new DataTable();

    async Task OpenExcelFileFromDisk(InputFileChangeEventArgs e)
    {
        IronXL.License.LicenseKey = "INSERT LICENSE KEY HERE";

        MemoryStream ms = new MemoryStream();
        await e.File.OpenReadStream().CopyToAsync(ms);
        ms.Position = 0;

        WorkBook loadedWorkBook = WorkBook.FromStream(ms);
        WorkSheet loadedWorkSheet = loadedWorkBook.DefaultWorkSheet;

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

The `<InputFile>` component facilitates the uploading of files. The `OpenExcelFileFromDisk` method, an asynchronous function in the bottom `@code` block, handles the upload and display of the Excel file as an HTML table.

IronXL.Excel is a versatile .NET library for reading a wide array of spreadsheet formats without requiring Microsoft Excel or relying on Interop.

---

<h4 class="tutorial-segment-title">Further Reading</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img src="https://ironsoftware.com/img/svgs/documentation.svg" alt="" class="img-responsive add-shadow" style="max-width: 110px; width: 100px; height: 140px;">
      </div>
    </div
    <div class="col-sm-8">
      <h3>Explore the API Reference for IronXL</h3>
      <p>Delve into the comprehensive API Reference which details all features, namespaces, classes, methods, fields, and enums of IronXL.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

*[Download IronXL](https://ironsoftware.com/csharp/excel/how-to/blazor-read-excel-file-tutorial/)*