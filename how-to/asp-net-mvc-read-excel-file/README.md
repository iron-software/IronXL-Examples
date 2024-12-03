# Parsing Excel Files in ASP.NET MVC with IronXL

***Based on <https://ironsoftware.com/how-to/asp-net-mvc-read-excel-file/>***


This guide describes how to incorporate Excel file reading capabilities into an ASP.NET MVC application using the IronXL library.

## Initiate an ASP.NET MVC Project

Start by creating a new ASP.NET MVC project in Visual Studio 2022, or a comparable version. Ensure you include all necessary NuGet packages and additional source code required.

## Integration of IronXL Library

Once the project is set up, proceed to integrate the IronXL library. Utilize NuGet Package Manager Console to execute the following command:

```shell
Install-Package IronXL.Excel
```

## Implement Excel File Reading

Navigate to the default controller in your project, typically `HomeController`, and modify the `Index` method by adding the following code:

```cs
public ActionResult Index()
{
    WorkBook workbook = WorkBook.Load(@"C:\Files\Customer Data.xlsx");
    WorkSheet sheet = workbook.WorkSheets.First();

    var dataTable = sheet.ToDataTable();

    return View(dataTable);
}
```

In this updated `Index` action, we use `WorkBook.Load` from IronXL to open the specified Excel file. The first worksheet in the workbook is accessed and transformed into a `DataTable` which is then passed to the view.

## Display Excel Data in a Web Browser

To show how the data can be displayed, the Excel file referenced is illustrated as follows:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/asp-net-mvc-read-excel-file/asp-net-mvc-read-excel-file-1.webp" alt="Read Excel Files in ASP.NET MVC Using IronXL, Figure 1: Excel file" class="img-responsive add-shadow">
        <p><em>Excel file</em></p>
    </div>
</div>

Adjust the `index.cshtml` (index view) file by replacing its content with the following HTML code:

```cshtml
@{
    ViewData["Title"] = "Home Page";
}

@using System.Data
@model DataTable

<div class="text-center">
    <h1 class="display-4">Welcome to IronXL Read Excel MVC</h1>
</div>
<table class="table table-dark">
    <tbody>
        @foreach (DataRow row in Model.Rows)
        {
            <tr>
                @for (int i = 0; i < Model.Columns.Count; i++)
                {
                    <td>@row[i]</td>
                }
            </tr>
        }
    </tbody>
</table>
```

This updated script uses the `DataTable` received from the `Index` method as its model. It formats the data in a Bootstrap-themed table, iterating through each row and displaying each cell.

Upon executing the project, the output will appear as follows:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/asp-net-mvc-read-excel-file/asp-net-mvc-read-excel-file-2.webp" alt="Read Excel Files in ASP.NET MVC Using IronXL, Figure 2: Bootstrap Table" class="img-responsive add-shadow">
        <p><em>Bootstrap Table</em></p>
    </div>
</div>