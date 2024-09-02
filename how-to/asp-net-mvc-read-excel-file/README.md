# Parsing Excel Files in ASP.NET MVC with IronXL

This tutorial explains how to parse Excel files within an ASP.NET MVC application by utilizing the IronXL library.

## Setting Up an ASP.NET Project

Begin by creating a new ASP.NET project in Visual Studio 2022 or a compatible version. Integrate any necessary NuGet packages and additional source code required for your specific project.

## Installing the IronXL Library

Once your project is set up, the next step is to install the IronXL library. Open the NuGet Package Manager Console and execute the following command:

```shell
Install-Package IronXL.Excel
```

## Implementing Excel File Reading

Navigate to the default controller in your ASP.NET project, typically the `HomeController`. Modify the `Index` method with the following code snippet:

```cs
public ActionResult Index()
{
    WorkBook workbook = WorkBook.Load(@"C:\Files\Customer Data.xlsx");
    WorkSheet sheet = workbook.GetFirstSheet();

    var dataTable = sheet.ToDataTable();

    return View(dataTable);
}
```

In this updated `Index` action method, the application begins by loading an Excel file using `WorkBook.Load`, specifying the full path to the file. It then accesses the first worksheet and converts its contents into a `DataTable` object which is then passed to the view.

## Displaying Excel Data in a Web Page

To display the `DataTable` in a web browser, use the following example:

The image below shows the Excel file used in this tutorial:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/asp-net-mvc-read-excel-file/asp-net-mvc-read-excel-file-1.webp" alt="Read Excel Files in ASP.NET MVC Using IronXL, Figure 1: Excel file" class="img-responsive add-shadow">
        <p><em>Excel file</em></p>
    </div>
</div>

Replace the content in `index.cshtml` (the index view) with the HTML code provided below:

```cs
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

This HTML code snippet utilizes the `DataTable` from the `Index` method as a model, iterating over each row and printing the content into a web page which benefits from Bootstrap styling.

Executing this project will display data as demonstrated in the following image:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/asp-net-mvc-read-excel-file/asp-net-mvc-read-excel-file-2.webp" alt="Read Excel Files in ASP.NET MVC Using IronXL, Figure 2: Bootstrap Table" class="img-responsive add-shadow">
        <p><em>Bootstrap Table</em></p>
    </div>
</div>
This structured approach allows for efficient processing and visualization of Excel data within an ASP.NET MVC application using IronXL.