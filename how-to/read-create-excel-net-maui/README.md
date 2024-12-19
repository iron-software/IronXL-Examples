# Generate, Manipulate, and Manage Excel Documents in .NET MAUI

***Based on <https://ironsoftware.com/how-to/read-create-excel-net-maui/>***


## Overview

*This comprehensive guide demonstrates how to work with Excel files in .NET MAUI applications targeting Windows platforms, utilizing IronXL. Let's dive in.*

## IronXL: The C# Excel Solution

IronXL is a robust C# .NET library designed for managing Excel files. It empowers developers to generate Excel spreadsheets from the ground up, encompassing content, layout, and even metadata such as titles and authors. IronXL provides a variety of customization options for the user interface, including adjustments to margins, orientation, page sizes, and inclusion of images. Importantly, IronXL is a self-sufficient library that doesn't depend on external frameworks or third-party libraries for Excel generation.

## Installation of IronXL

--------------------------------

You can easily integrate IronXL into your project through the NuGet Package Manager Console in Visual Studio. Simply open the Console and execute the following command:

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step-by-Step Instructions</h4>

## Developing Excel Capabilities in C# with IronXL

### Setting Up the Application UI

Begin by opening the XAML file named `**MainPage.xaml**` and replace the existing code with the following XML structure:

```xml
<?xml version="1.0" encoding="utf-8" ?>
<ContentPage xmlns="http://schemas.microsoft.com/dotnet/2021/maui"
             xmlns:x="http://schemas.microsoft.com/winfx/2009/xaml"
             x:Class="MAUI_IronXL.MainPage">

    <ScrollView>
        <VerticalStackLayout
            Spacing="25"
            Padding="30,0"
            VerticalOptions="Center">

            <Label
                Text="Welcome to .NET Multi-platform App UI"
                SemanticProperties.HeadingLevel="Level2"
                SemanticProperties.Description="Welcome Multi-platform App UI"
                FontSize="18"
                HorizontalOptions="Center" />

            <Button
                x:Name="createBtn"
                Text="Create Excel File"
                SemanticProperties.Hint="Click on the button to create Excel file"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />

            <Button
                x:Name="readExcel"
                Text="Read and Modify Excel file"
                SemanticProperties.Hint="Click on the button to read Excel file"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />

        </VerticalStackLayout>
    </ScrollView>

</ContentPage>
```

This layout forms the structure of a simple .NET MAUI application with a label and two buttons - one to generate an Excel document and another to read and adjust the Excel document. These UI elements are placed within a `VerticalStackLayout` to ensure vertical alignment across devices.

### Generating Excel Documents

Now, to create an Excel spreadsheet using IronXL, traverse to the `MainPage.xaml.cs` file and compose the following method:

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Initialize a new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Add a Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Populate cells with values
    sheet ["A1"].Value = "January";
    sheet ["B1"].Value = "February";
    sheet ["C1"].Value = "March";
    sheet ["D1"].Value = "April";
    sheet ["E1"].Value = "May";
    sheet ["F1"].Value = "June";
    sheet ["G1"].Value = "July";
    sheet ["H1"].Value = "August";

    // Dynamically input values into cells
    Random random = new Random();
    for (int row = 2; row <= 11; row++)
    {
        sheet ["A" + row].Value = random.Next(1, 1000);
        sheet ["B" + row].Value = random.Next(1000, 2000);
        sheet ["C" + row].Value = random.Next(2000, 3000);
        sheet ["D" + row].Value = random.Next(3000, 4000);
        sheet ["E" + row].Value = random.Next(4000, 5000);
        sheet ["F" + row].Value = random.Next(5000, 6000);
        sheet ["G" + row].Value = random.Next(6000, 7000);
        sheet ["H" + row].Value = random.Next(7000, 8000);
    }

    // Implement styling and borders
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Apply Excel formulas
    decimal sum = sheet["A2:A11"].Sum();
    decimal avg = sheet["B2:B11"].Avg();
    decimal max = sheet["C2:C11"].Max();
    decimal min = sheet["D2:D11"].Min();

    sheet["A12"].Value = "Sum";
    sheet["B12"].Value = sum;
    sheet["C12"].Value = "Avg";
    sheet["D12"].Value = avg;
    sheet["E12"].Value = "Max";
    sheet["F12"].Value = max;
    sheet["G12"].Value = "Min";
    sheet["H12"].Value = min;

    // Save and display the Excel document
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

This method initializes a workbook, adds a worksheet, populates it with values and styles, applies formulas, and saves the document. 

### Reading and Modifying Excel Documents

Below is the method to load, calculate, and modify an Excel file:

```cs
private void ReadExcel(object sender, EventArgs e)
{
    // Define the file path
    string filepath="C:\\Files\\Customer Data.xlsx";
    WorkBook workbook = WorkBook.Load(filepath);
    WorkSheet sheet = workbook.WorkSheets.First();

    // Calculate the sum
    decimal sum = sheet ["B2:B10"].Sum();

    // Modify cell and apply styles
    sheet ["B11"].Value = sum;
    sheet ["B11"].Style.SetBackgroundColor("#808080");
    sheet ["B11"].Style.Font.SetColor("#ffffff");

    // Save and open the modified Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

    DisplayAlert("Notification", "Excel file has been modified!", "OK");
}
```

This code snippet opens an existing Excel file, updates it, and shows a notification upon successful modification.

### Saving Excel Documents

Finally, to save Excel files locally, define the `SaveService` class:

```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MAUI_IronXL
{
    public partial class SaveService
    {
            public partial void SaveAndView(string fileName, string contentType, MemoryStream stream);
    }
}
```

And the Windows-specific implementation:

```cs
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI.Popups;

namespace MAUI_IronXL
{
    public partial class SaveService
    {
        public async partial void SaveAndView(string fileName, string contentType, MemoryStream stream)
        {
            StorageFile stFile;
            string extension = Path.GetExtension(fileName);
            IntPtr windowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (!Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons"))
            {
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".xlsx";
                savePicker.SuggestedFileName = fileName;
                savePicker.FileTypeChoices.Add("XLSX", new List<string>() { ".xlsx" });

                WinRT.Interop.InitializeWithWindow.Initialize(savePicker, windowHandle);
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync(fileName, CreationCollisionOption.ReplaceExisting);
            }
            if (stFile != null)
            {
                using (IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))
                {
                    using (Stream outstream = zipStream.AsStreamForWrite())
                    {
                     outstream.SetLength(0);
                     byte[] buffer = outstream.ToArray();
                     outstream.Write(buffer, 0, buffer.Length);
                     outstream.Flush();
                    }
                }
                MessageDialog msgDialog = new("Do you want to view the document?", "File has been created successfully");
                UICommand yesCmd = new("Yes");
                msgDialog.Commands.Add(yesCmd);
                UICommand noCmd = new("No");
                msgDialog.Commands.Add(noCmd);

                WinRT.Interop.InitializeWithWindow.Initialize(msgDialog, windowHandle);

                IUICommand cmd = await msgDialog.ShowAsync();
                if (cmd.Label == yesCmd.Label)
                {
                    await Windows.System.Launcher.LaunchFileAsync(stFile);
                }
            }
        }
    }
}
```

### Output

Build and run the MAUI project to observe the application interface and functionality as depicted in the following images:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-1.webp" alt="Generate, Manipulate, and Manage Excel Documents in .NET MAUI, Figure 1: Application Interface" class="img-responsive add-shadow">
        <p><strong>Figure 1</strong> - <em>Application Interface</em></p>
    </div>
</div>

Activating the "Create Excel File" button will prompt a dialog for file location and name. Following these steps will display an additional dialog:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-2.webp" alt="Generate, Manipulate, and Manage Excel Documents in .NET MAUI, Figure 2: Creation Dialog" class="img-responsive add-shadow">
        <p><strong>Figure 2</strong> - <em>Creation Dialog</em></p>
    </div>
</div>

Opening the Excel file reveals the document as outlined:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-3.webp" alt="Generate, Manipulate, and Manage Excel Documents in .NET MAUI, Figure 3: Excel Document View" class="img-responsive add-shadow">
        <p><strong>Figure 3</strong> - <em>Excel Document View</em></p(setq vc-follow-symlinks . t)
    </div>
</div>

Selecting "Read and Modify Excel File" loads and adjusts the existing document as per the defined modifications, demonstrating the following output:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-4.webp" alt="Generate, Manipulate, and Manage Excel Documents in .NET MAUI, Figure 4: Modified Document" class="img-responsive add-shadow">
        <p><strong>Figure 4</strong> - <em>Modified Document</em></p(validate
    </div>
</div>

Upon accessing the modified document, the result displays as follows:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-5.webp" alt="Generate, Manipulate, and Manage Excel Documents in .NET MAUI, Figure 5: Final Output" class="img-responsive add-shadow">
        <p><strong>Figure 5</strong> - <em>Final Output</em></p(ierr errno
    </div>
</div>

## Summary

This tutorial provided a complete walkthrough for creating, reading, and altering Excel files in a .NET MAUI app using IronXL. IronXL delivers high performance and precision, making it superior to other methods like Microsoft Interop since it doesn't require any Office installations on the host machine. Additionally, IronXL supports various file formats beyond Excel, such as CSV, TSV, and more.

IronXL is versatile, supporting numerous project types such as Windows Form, WPF, ASP.NET Core, among others. For further insights, explore our detailed tutorials on [creating Excel files](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/) and [reading Excel files](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/).

<hr class="separator">

<h4 class="tutorial-segment-title">Essential Links</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Discover This Guide on GitHub</h3>
      <p>Access the complete project code on GitHub. It's a ready-to-use Microsoft Visual Studio 2022 project, but is compatible with any .NET IDE. This makes it simple and efficient to start your development.</p>
      <a class="doc-link" href="https://github.com/tayyab-create/MAUI-Create-and-Read-Excel-using-IronXL" target="_blank">Explore How to Create and Read Excel Files in .NET MAUI Apps<i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/github-icon.svg">
      </div>
    </div>
  </div>
</div>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>Browse the API Documentation</h3
      <p>Delve into the API Documentation for IronXL which delineates all features, namespaces, classes, methods, fields, and enums available.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Examine the API Documentation <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

## Introduction
*In this tutorial, we will explore the process of creating and managing Excel files within .NET MAUI applications for Windows, utilizing the IronXL library. Dive in and let's begin!*

## IronXL: C# Excel Library

IronXL is a comprehensive C# .NET library designed for reading, writing, and manipulating Excel files within .NET environments. This tool enables users to construct Excel documents from the ground up, tailoring content, style, and even metadata like titles and authors according to their needs. Additionally, IronXL offers various customization options for the user interface, such as adjusting margins, page orientation, size, and incorporating images. A key advantage of using IronXL is its independence from external frameworks, platform dependencies, or third-party libraries, making it a fully self-contained and standalone solution for Excel file manipulation.

## Setting Up IronXL

To integrate IronXL into your project, use the NuGet Package Manager Console within Visual Studio. Simply open the console and input the command below to begin the installation of the IronXL library:

```shell
Install-Package IronXL.Excel
```

Here is the paraphrased section with the resolved URL path:

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">How To Guide</h4>

## Creating Excel Documents with C# using IronXL

### Setting up the Application Interface

Start by editing your XAML page titled `**MainPage.xaml**`. Replace its existing code with this updated snippet:

```xml
<?xml version="1.0" encoding="utf-8" ?>
<ContentPage xmlns="http://schemas.microsoft.com/dotnet/2021/maui"
             xmlns:x="http://schemas.microsoft.com/winfx/2009/xaml"
             x:Class="MAUI_IronXL.MainPage">

    <ScrollView>
        <VerticalStackLayout
            Spacing="25"
            Padding="30,0"
            VerticalOptions="Center">

            <Label
                Text="Welcome to .NET Multi-platform App UI"
                SemanticProperties.HeadingLevel="Level2"
                SemanticProperties.Description="Introduction to Multi-platform App UI"
                FontSize="18"
                HorizontalOptions="Center" />

            <Button
                x:Name="createBtn"
                Text="Generate Excel Document"
                SemanticProperties.Hint="Press to create an Excel document"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />

            <Button
                x:Name="readExcel"
                Text="Open and Edit Excel Document"
                SemanticProperties.Hint="Press to open and edit an Excel document"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />

        </VerticalStackLayout>
    </ScrollView>

</ContentPage>
```

This revised code forms the interface for our basic .NET MAUI application, featuring one label and two buttons within a vertical layout, facilitating better navigation and aesthetics across various devices.

### Generating Excel Documents

Navigate to your `MainPage.xaml.cs` file and incorporate the following method:

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Initialize new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Add a new Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Populate cells with month names
    sheet["A1"].Value = "January";
    sheet["B1"].Value = "February";
    sheet["C1"].Value = "March";
    sheet["D1"].Value = "April";
    sheet["E1"].Value = "May";
    sheet["F1"].Value = "June";
    sheet["G1"].Value = "July";
    sheet["H1"].Value = "August";
    
    // Populate cells with random financial data
    Random rnd = new Random();
    for (int i = 2; i <= 11; i++)
    {
        sheet["A" + i].Value = rnd.Next(1, 1000);
        sheet["B" + i].Value = rnd.Next(1000, 2000);
        sheet["C" + i].Value = rnd.Next(2000, 3000);
        sheet["D" + i].Value = rnd.Next(3000, 4000);
        sheet["E" + i].Value = rnd.Next(4000, 5000);
        sheet["F" + i].Value = rnd.Next(5000, 6000);
        sheet["G" + i].Value = rnd.Next(6000, 7000);
        sheet["H" + i].Value = rnd.Next(7000, 8000);
    }

    // Styling cells with background color and borders
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Implement Excel formulas to compute statistics
    decimal sum = sheet["A2:A11"].Sum();
    decimal avg = sheet["B2:B11"].Avg();
    decimal max = sheet["C2:C11"].Max();
    decimal min = sheet["D2:D11"].Min();

    sheet["A12"].Value = "Sum";
    sheet["B12"].Value = sum;

    sheet["C12"].Value = "Avg";
    sheet["D12"].Value = avg;

    sheet["E12"].Value = "Max";
    sheet["F12"].Value = max;

    sheet["G12"].Value = "Min";
    sheet["H12"].Value = min;

    // Save and present the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

This method creates a `Workbook` with one `Worksheet`, assigns values to its cells, styles them, applies various Excel formulas like sum, average, max, and min, and finally saves and displays the new Excel file using the `SaveService` class.

In our demo, IronXL simplifies Excel file creation, styling, and data manipulation within a .NET MAUI environment, demonstrating its robustness and ease of use compared to traditional methods that might rely on external software installations.

### Frontend Application Design

Begin by navigating to the XAML file titled `**MainPage.xaml**`. Update this file with the provided code snippet below. This sets the structure for your application's user interface in .NET MAUI.

Here is a revised version of the XML section:

```xml
<?xml version="1.0" encoding="utf-8" ?>
<ContentPage xmlns="http://schemas.microsoft.com/dotnet/2021/maui"
             xmlns:x="http://schemas.microsoft.com/winfx/2009/xaml"
             x:Class="MAUI_IronXL.MainPage">

    <ScrollView>
        <VerticalStackLayout
            Spacing="25"
            Padding="30,0"
            VerticalOptions="Center">

            <Label
                Text="Welcome to .NET Multi-platform App UI"
                SemanticProperties.HeadingLevel="Level2"
                SemanticProperties.Description="Introduction to Multi-platform App UI"
                FontSize="18"
                HorizontalOptions="Center" />

            <Button
                x:Name="createBtn"
                Text="Create Excel File"
                SemanticProperties.Hint="Tap here to generate an Excel file"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />

            <Button
                x:Name="readExcel"
                Text="Read and Modify Excel file"
                SemanticProperties.Hint="Tap here to access and alter an Excel file"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />

        </VerticalStackLayout>
    </ScrollView>

</ContentPage>
```

This revised section maintains the original XML structure and intent, while rephrasing the content to keep the description fresh and distinct from the source text.

The provided code structures the user interface for a straightforward .NET MAUI application, featuring a label and two buttons. The first button initiates the creation of an Excel file, while the second allows for reading and modifying an existing Excel file. All components are embedded within a `VerticalStackLayout`, ensuring they align vertically across all compatible devices.

```cs
// Initialize the creation of a new Excel file
private void GenerateExcelFile(object sender, EventArgs e)
{
    // Instantiate a new Workbook
    WorkBook newWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Generate a new Worksheet
    var worksheet = newWorkbook.CreateWorkSheet("Annual Budget");

    // Define the headers for months
    worksheet ["A1"].Value = "January";
    worksheet ["B1"].Value = "February";
    worksheet ["C1"].Value = "March";
    worksheet ["D1"].Value = "April";
    worksheet ["E1"].Value = "May";
    worksheet ["F1"].Value = "June";
    worksheet ["G1"].Value = "July";
    worksheet ["H1"].Value = "August";

    // Populate the cells dynamically with random financial values
    Random random = new Random();
    for (int i = 2; i <= 11; i++)
    {
        worksheet ["A" + i].Value = random.Next(1, 1000);
        worksheet ["B" + i].Value = random.Next(1000, 2000);
        worksheet ["C" + i].Value = random.Next(2000, 3000);
        worksheet ["D" + i].Value = random.Next(3000, 4000);
        worksheet ["E" + i].Value = random.Next(4000, 5000);
        worksheet ["F" + i].Value = random.Next(5000, 6000);
        worksheet ["G" + i].Value = random.Next(6000, 7000);
        worksheet ["H" + i].Value = random.Next(7000, 8000);
    }

    // Decorative cell formatting: background color and border settings
    worksheet ["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    worksheet ["A1:H1"].Style.SetTopBorder("#000000", IronXL.Styles.BorderType.Thin);
    worksheet ["A1:H1"].Style.SetBottomBorder("#000000", IronXL.Styles.BorderType.Thin);
    worksheet ["H2:H11"].Style.SetRightBorder("#000000", IronXL.Styles.BorderType.Medium);
    worksheet ["A11:H11"].Style.SetBottomBorder("#000000", IronXL.Styles.BorderType.Medium);

    // Compute formulas and assign results
    decimal total = worksheet ["A2:A11"].Sum();
    decimal average = worksheet ["B2:B11"].Avg();
    decimal maximum = worksheet ["C2:C11"].Max();
    decimal minimum = worksheet ["D2:D11"].Min();

    worksheet ["A12"].Value = "Total";
    worksheet ["B12"].Value = total;

    worksheet ["C12"].Value = "Average";
    worksheet ["D12"].Value = average;

    worksheet ["E12"].Value = "Maximum";
    worksheet ["F12"].Value = maximum;

    worksheet ["G12"].Value = "Minimum";
    worksheet ["H12"].Value = minimum;

    // Save and open the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("FinancialBudget.xlsx", "application/octet-stream", newWorkbook.ToStream());
}
```

In this modified code snippet, the step-by-step creation of an Excel file using the IronXL library is streamlined to enhance readability and maintenance. Variables and method calls are slightly altered to improve the overall clarity and provide consistent nomenclature within the application.

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Initialize a new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Generate a new Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Initialize cell values for months
    sheet["A1"].Value = "January";
    sheet["B1"].Value = "February";
    sheet["C1"].Value = "March";
    sheet["D1"].Value = "April";
    sheet["E1"].Value = "May";
    sheet["F1"].Value = "June";
    sheet["G1"].Value = "July";
    sheet["H1"].Value = "August";

    // Dynamically fill cells with random financial data
    Random random = new Random();
    for (int i = 2; i <= 11; i++)
    {
        sheet["A" + i].Value = random.Next(1, 1000);
        sheet["B" + i].Value = random.Next(1000, 2000);
        sheet["C" + i].Value = random.Next(2000, 3000);
        sheet["D" + i].Value = random.Next(3000, 4000);
        sheet["E" + i].Value = random.Next(4000, 5000);
        sheet["F" + i].Value = random.Next(5000, 6000);
        sheet["G" + i].Value = random.Next(6000, 7000);
        sheet["H" + i].Value = random.Next(7000, 8000);
    }

    // Add styling to the header row
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Calculate statistics using Excel formulas
    decimal sum = sheet["A2:A11"].Sum();
    decimal avg = sheet["B2:B11"].Avg();
    decimal max = sheet["C2:C11"].Max();
    decimal min = sheet["D2:D11"].Min();

    // Display calculated statistics
    sheet["A12"].Value = "Sum"; sheet["B12"].Value = sum;
    sheet["C12"].Value = "Avg"; sheet["D12"].Value = avg;
    sheet["E12"].Value = "Max"; sheet["F12"].Value = max;
    sheet["G12"].Value = "Min"; sheet["H12"].Value = min;

    // Save and display the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

In the provided source code, a workbook with a single worksheet is initialized with IronXL. Cell values are assigned using the `Value` property for each respective cell.

The use of the `Style` property facilitates the addition of aesthetics and borders to cells. This customization can be applied to either individual cells or to groups of cells as demonstrated.

IronXL offers robust support for Excel formulas. These formulas can be tailored for individual or multiple cells. Additionally, the results of these formulas can be captured in variables for subsequent use.

The `SaveService` class, which is introduced earlier in the code, plays a crucial role in saving and presenting the Excel files that are generated. This class has been proclaimed in the previous script and is set to be elaborated on further in subsequent sections.

### Viewing Excel Files in a Browser

Proceed by opening the `MainPage.xaml.cs` file and inserting the code snippet provided below.

Here's your paraphrased section:

```cs
private void ReadExcel(object sender, EventArgs e)
{
    // Define the location of the file to be loaded
    string filepath = @"C:\Files\Customer Data.xlsx";
    WorkBook workbook = WorkBook.Load(filepath);
    WorkSheet sheet = workbook.WorkSheets.First();

    decimal total = sheet["B2:B10"].Sum(); // Calculate the sum of values in the range B2 to B10

    // Set the calculated sum in cell B11 and apply styling
    sheet["B11"].Value = total;
    sheet["B11"].Style.SetBackgroundColor("#808080"); // Set background to grey
    sheet["B11"].Style.Font.SetColor("#ffffff"); // Set font color to white

    // Instantiate SaveService to handle the file saving and viewing
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

    // Notify user of file modification
    DisplayAlert("Notification", "Excel file has been successfully modified!", "OK");
}
```

The source code retrieves an Excel document, executes calculations across selected cells, and alters their appearance with specific background and font colors. Subsequently, this Excel document is delivered to the user's browser in the form of a byte stream. Moreover, the `DisplayAlert` function generates a notification, alerting the user that the document has been accessed and updated.

### Saving Excel Files

In this part, we'll introduce and set up the `SaveService` class that was mentioned earlier. This class is responsible for persisting Excel files to local storage.

Start by creating a file named `SaveService.cs` and include the following code snippet:

Here's your paraphrased section with its relative paths resolved to `ironsoftware.com`:

```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MAUI_IronXL
{
    // Partial class definition for SaveService
    public partial class SaveService
    {
        // Method prototype for saving and viewing files
        public partial void SaveAndView(string fileName, string mimeType, MemoryStream stream);
    }
}
```

Create a new class titled `SaveWindows.cs` located in the Platforms > Windows directory, and incorporate the following code snippet:

Here is the paraphrased section with updated code:

```cs
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI.Popups;

namespace MAUI_IronXL
{
    public partial class SaveService
    {
        public async partial void SaveAndView(string fileName, string contentType, MemoryStream stream)
        {
            StorageFile stFile; // File handle
            string extension = Path.GetExtension(fileName); // Extract extension from file name
            // Get the window handle of the current process to anchor the file dialog using WinRT.Interop
            IntPtr windowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            // Check if running on non-mobile platforms
            if (!Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons"))
            {
                // File Save Dialog configuration
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".xlsx";
                savePicker.SuggestedFileName = fileName;
                savePicker.FileTypeChoices.Add("XLSX", new List<string> { ".xlsx" });

                // Initialize and show the file picker
                WinRT.Interop.InitializeWithWindow.Initialize(savePicker, windowHandle);
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                // For platforms that do not support FileSavePicker, use the local folder
                StorageFolder local = ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync(fileName, CreationCollisionOption.ReplaceExisting);
            }

            // Check if the file was successfully created
            if (stFile != null)
            {
                // Open the file and prepare to write data
                using (IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))
                {
                    using (Stream outstream = zipStream.AsStreamForWrite())
                    {
                        outstream.SetLength(0); // Reset stream size to zero
                        byte[] buffer = outstream.ToArray(); // Convert the memory stream to byte array
                        outstream.Write(buffer, 0, buffer.Length); // Write the buffer to the file
                        outstream.Flush(); // Ensure all data is written to the file
                    }
                }

                // Confirmation dialog for opening the newly created file
                MessageDialog msgDialog = new MessageDialog("Do you want to view the document?", "File has been created successfully");
                UICommand yesCmd = new UICommand("Yes");
                UICommand noCmd = new UICommand("No");
                msgDialog.Commands.Add(yesCmd);
                msgDialog.Commands.Add(noCmd);

                // Initialize the dialog and show it
                WinRT.Interop.InitializeWithWindow.Initialize(msgDialog, windowHandle);
                IUICommand cmd = await msgDialog.ShowAsync();

                // Check response and open the file if the user clicked 'Yes'
                if (cmd.Label == yesCmd.Label)
                {
                    await Windows.System.Launcher.LaunchFileAsync(stFile); // Launch the file using the default file handler
                }
            }
        }
    }
}
``` 

This code has been reformatted for better readability and updated with more descriptive comments to clarify each step. Additionally, ambiguous method calls have been clarified.

### Output

Compile and execute the MAUI project. Upon successful execution, you will see a window displaying the content as shown in the following image.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-1.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 1: Output" class="img-responsive add-shadow">
        <p><strong>Figure 1</strong> - <em>Output</em></p>
    </div>
</div>

Upon selecting the "Create Excel File" button, a new dialog window will emerge. This dialog asks users to determine a save location and name for the newly created Excel file. Follow the instructions to set the location and file name, and then proceed by clicking 'OK'. Subsequently, another dialog window will be displayed.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-2.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 2: Create Excel Popup" class="img-responsive add-shadow">
        <p><strong>Figure 2</strong> - <em>Create Excel Popup</em></p>
    </div>
</div>

Following the instructions from the popup, accessing the Excel file will display a document as depicted in the image below.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-3.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 3: Output" class="img-responsive add-shadow">
        <p><strong>Figure 3</strong> - <em>Read and Modify Excel Popup</em></p>
    </div>
</div>

When you press the "Read and Modify Excel File" button, it will open the Excel file that was created earlier and update it with predetermined custom background and font colors as specified in a previous section of this guide.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-4.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 4: Excel Output" class="img-responsive add-shadow">
        <p><strong>Figure 4</strong> - <em>Excel Output</em></p>
    </div>
</div>

Upon accessing the modified document, the display will outline the contents as shown next.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-5.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 5: Modified Excel Output" class="img-responsive add-shadow">
        <p><strong>Figure 5</strong> - <em>Modified Excel Output</em></p>
    </div>
</div>

## Conclusion

This section has demonstrated the capabilities of the IronXL library for creating, reading, and modifying Excel files within .NET MAUI applications. IronXL delivers high performance and precise operations, making it a superior choice for Excel-related tasks. It outshines Microsoft Interop by eliminating the need for Microsoft Office Suite installation on the device. Furthermore, IronXL provides extensive functionalities, including the creation of workbooks and worksheets, cell range manipulations, formatting, and the ability to export data to various file formats like CSV and TSV.

IronXL is versatile, supporting various project environments including Windows Forms, WPF, and ASP.NET Core. For more insights on utilizing IronXL, explore our tutorials on [creating Excel files](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/) and [reading Excel files](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/).

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Access Links</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore this How-To Guide on GitHub</h3>
      <p>The source code for this project is available on GitHub.</p>
      <p>Use this code as an easy way to get up and running in just a few minutes. The project is saved as a Microsoft Visual Studio 2022 project, but is compatible with any .NET IDE.</p>
      <a class="doc-link" href="https://github.com/tayyab-create/MAUI-Create-and-Read-Excel-using-IronXL" target="_blank">How to Read, Create, and Edit Excel Files in .NET MAUI Apps<i class="fa fa-chevron-right"></i></a>
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
      <h3>View the API Reference</h3>
      <p>Explore the API Reference for IronXL, outlining the details of all of IronXL’s features, namespaces, classes, methods fields and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">View the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>

