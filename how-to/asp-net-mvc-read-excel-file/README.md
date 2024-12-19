# Read Excel Files in ASP.NET MVC Using IronXL

***Based on <https://ironsoftware.com/how-to/asp-net-mvc-read-excel-file/>***


This guide will walk developers through the steps required to incorporate Excel file reading functionality within ASP.NET MVC applications using IronXL.

## Create an ASP.NET Project

First, launch Visual Studio 2022 or a similar IDE and create a new ASP.NET project. Integrate any necessary NuGet packages and add the source code required for your specific application.

## Install IronXL Library

Once your ASP.NET project is set up, the next step is to install the IronXL library for working with Excel files. You can easily install it by executing the following command in the NuGet Package Manager Console:

```shell
Install-Package IronXL.Excel
```

## Reading Excel file

Navigate to the default controller in your ASP.NET project, typically `HomeController`, and update the `Index` method as shown below:

```cs
public ActionResult Index()
{
    WorkBook workbook = WorkBook.Load(@"C:\Files\Customer Data.xlsx");
    WorkSheet sheet = workbook.WorkSheets.First();

    var dataTable = sheet.ToDataTable();

    return View(dataTable);
}
```

In this code snippet, the `Index` method starts by utilizing IronXL's `Load` method to open an Excel file. The path to the Excel document is furnished as an argument to this method. After loading, the first worksheet of the workbook is designated as the active worksheet and its contents are loaded into a `DataTable` object, which is then passed to the view.

## Display Excel Data on a Web Page

This segment illustrates the method of displaying the `DataTable` obtained from the previous section in a web browser.

Here's the Excel file we'll be showcasing:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://www.ironsoftware.com/img/tutorials/asp-net-mvc-read-excel-file/asp-net-mvc-read-excel-file-1.webp" alt="Read Excel Files in ASP.NET MVC Using IronXL, Figure 1: Excel file" class="img-responsive add-shadow">
        <p><em>Excel file</em></p>
    </div>
</div>

Replace the existing code in the `index.cshtml` (index view) with the following HTML markup:

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

This HTML code utilizes the aforementioned `DataTable` as a model, rendering each of its rows into a web page using a loop, while Bootstrap attributes enhance the visual formatting.

Upon running your project, the output will appear as follows.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://www.ironsoftware.com/img/tutorials/asp-net-mvc-read-excel-file/asp-net-mvc-read-excel-file-2.webp"  alt="Read Excel Files in ASP.NET MVC Using IronXL, Figure 2: Bootstrap Table" class="img-responsive add-shadow">
        <p><em>Bootstrap Table</em></p>
    </div>
</div>