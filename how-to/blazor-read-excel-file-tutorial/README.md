# Blazor Read Excel File in C# Using IronXL (Example Guide)

***Based on <https://ironsoftware.com/how-to/blazor-read-excel-file-tutorial/>***


## Introduction

Blazor, Microsoft's open-source .NET Web framework, allows developers to compile C# code into JavaScript and HTML that run in the browser. This guide provides insights on how to efficiently read Excel files in a Blazor serverside application using the IronXL C# library.

![Demonstration of IronXL Viewing Excel in Blazor](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/demo.gif)

## Step 1 - Setting Up a Blazor Project in Visual Studio

To demonstrate reading data from an XLSX file, I will use the following data set, which will then be displayed in the Blazor Server App:


| Input XLSX Excel Sheet | Result in Blazor Server Browser |
| ----------------------- | ------------------------------- |
| First name  | Last name   | ID | (Image of Web Browser)       |
| John        | Applesmith  | 1  | ![Browser View](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/browser-view.webp)  |
| Richard     | Smith       | 2  |                            |
| Sherry      | Robins      | 3  |                            |
```

Begin by initializing a new Blazor Project using the Visual Studio IDE:
- ![Create New Project](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/new-project.webp)
- Select the **Blazor Server App** as the project type:
  ![Choose Blazor Project Type](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/choose-blazor-project-type.webp)
- Launch the application using the `F5` key and head over to the `Fetch data` tab:
  ![First Application Run](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/first-run.webp)

Our objective is to level up the Blazor app by adding functionality to upload and display Excel files.

## Step 2 - Integrating IronXL in Your Solution

### IronXL: .NET Excel Library (Installation Guide):

IronXL serves as a robust .NET library that transforms Excel spreadsheets into manipulatable objects. This leveraging allows developers to utilize C# and the .NET Framework to engage with data efficiently. IronXL provides enhanced functionality over alternatives like NPOI, offering more features, easier complex logic implementation, better licenses, and superior support.

IronXL is compatible with the latest releases of .NET including versions 8, 7, and 6, as well as .NET Core Framework 4.6.2+.

IronXL can be added to your project either via NuGet Package Manager or by directly adding a package reference in your solution's `.csproj` file:

### Option 2A - Using NuGet Package Manager

### Option 2B - Include PackageReference in the .csproj

Enter the following line in any `<ItemGroup>` in your `.csproj` file:

```xml
<PackageReference Include="IronXL.Excel" Version="*" />
```

Visualized here in Visual Studio:
![Add IronXL to Project](https://ironsoftware.com/static-assets/excel/how-to/blazor-read-excel-file-tutorial/add-ironxl-csproj.webp)

## Step 3 - Implementing File Upload and Display

In the Visual Studio's Solution Explorer, navigate to the `Pages/` directory and open the `FetchData.razor` file. Although any Razor file can be modified, `FetchData.razor` is recommended since it's included in the Blazor Server App Template.

Replace the content with the following code:

```cs
@using IronXL;
@using System.Data;

@page "/fetchdata"

<PageTitle>View Excel File</PageTitle>

<h1>Load and Display Excel File</h1>

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
        IronXL.License.LicenseKey = "YOUR LICENSE KEY HERE";
        MemoryStream ms = new MemoryStream();
        await e.File.OpenReadStream().CopyToAsync(ms);
        ms.Position = 0;
        WorkBook workbook = WorkBook.Load(ms);
        WorkSheet worksheet = workbook.DefaultWorkSheet;
        
        // Setup columns from headers
        RangeRow headerRow = worksheet.Rows.First();
        foreach (var cell in headerRow)
        {
            displayDataTable.Columns.Add(cell.StringValue);
        }
        
        // Populate rows
        foreach (var row in worksheet.Rows.Skip(1))
        {
            var rowValues = row.Select(cell => cell.StringValue).ToArray();
            displayDataTable.Rows.Add(rowValues);
        }
    }
}
```

## Summary

Utilize the `<InputFile>` component for file uploads on the webpage. The uploaded Excel files can be viewed directly by triggering the `OpenExcelFileFromDisk` event.

IronXL stands out as a top-notch .NET library for reading diverse spreadsheet formats seamlessly without the prerequisite of having Microsoft Excel installed.

---

### Additional Resources

<div class="tutorial-section">
  <img src="https://ironsoftware.com/img/svgs/documentation.svg" alt="" style="width: 100px; height: 140px;">
  <h3>Dive Deeper with API Documentation</h3>
  <p>Delve into the comprehensive API reference for IronXL to explore its extensive functionalities, namespaces, and coding components.</p>
  <a href="https://ironsoftware.com/csharp/excel/object-reference/api/">Discover More <i class="fa fa-chevron-right"></i></a>
</div>

*[Download IronXL](https://ironsoftware.com/csharp/excel/how-to/blazor-read-excel-file-tutorial/)*