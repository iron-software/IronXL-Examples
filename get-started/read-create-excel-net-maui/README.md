# Create, Read and Edit Excel Files in .NET MAUI

> Full guide: [Create, Read and Edit Excel Files in .NET MAUI](https://ironsoftware.com/csharp/excel/get-started/read-create-excel-net-maui/)

## Introduction

*Welcome to our How-To Guide on using IronXL to create and read Excel files within .NET MAUI applications on Windows.*

## IronXL: C# Excel Library

IronXL is a C# .NET library designed for the manipulation, creation, and reading of Excel files. This library allows users to build Excel documents from the ground up, manage content and styles, and set document properties like title and author. IronXL offers features for customizing the user interface, such as adjusting margins, setting page orientation, adding images, and more without the need for any third-party libraries or external frameworks. It operates independently as a self-contained unit.

Incorporating IronXL into your project is straightforward using the NuGet Package Manager Console within Visual Studio. Simply open up the Console and execute the given command:

```shell
Install-Package IronXL.Excel
```

## Creating Excel Files in C# with IronXL

### Designing the Application Frontend

Begin by opening the XAML page named `MainPage.xaml` and replace its content with the following code block:

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
                SemanticProperties.Hint="Click on the button to create an Excel file"
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

This snippet configures the UI for our basic .NET MAUI app, including labels and buttons to create and read Excel files, arranged vertically.

### Generating Excel Files

Now, let's generate an Excel file. Open the `MainPage.xaml.cs` file and implement the following method:

```csharp
private void CreateExcel(object sender, EventArgs e)
{
    // Instantiate a new Workbook with an Excel format
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Generate a Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Define headers for the cells
    string[] headers = { "January", "February", "March", "April", "May", "June", "July", "August" };
    for (int i = 0; i < headers.Length; i++)
    {
        sheet[$"{(char)('A' + i)}1"].Value = headers[i];
    }

    // Populate cells with random data
    Random random = new Random();
    for (int row = 2; row <= 11; row++)
    {
        for (int col = 0; col < 8; col++)
        {
            sheet[$"{(char)('A' + col)}{row}"].Value = random.Next(1, 8001);
        }
    }

    // Formatting: Set background, borders and apply formulas
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3")
           .TopBorder.SetColor("#000000")
           .BottomBorder.SetColor("#000000")
           .BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    decimal[] results = new decimal[4];
    results[0] = sheet["A2:A11"].Sum();
    results[1] = sheet["B2:B11"].Avg();
    results[2] = sheet["C2:C11"].Max();
    results[3] = sheet["D2:D11"].Min();

    string[] markers = { "Sum", "Avg", "Max", "Min" };
    for (int i = 0; i < results.Length; i++)
    {
        sheet[$"{(char)('A' + i)}12"].Value = markers[i];
        sheet[$"{(char)('B' + i)}12"].Value = results[i];
    }

    // Save and display the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

This updated method creates a workbook, initializes cells with headers, populates them with random numbers, applies formatting, and integrates Excel formulas to aggregate data. The modified Excel document is then saved and could be viewed immediately after creation.

### Displaying Excel Files in the Browser

Let's proceed with reading and displaying Excel files. Modify the `MainPage.xaml.cs` file with these additions:

```csharp
private void ReadExcel(object sender, EventArgs e)
{
    // Define the Excel file location
    string filepath = @"C:\Files\Customer Data.xlsx";
    WorkBook workbook = WorkBook.Load(filepath);
    WorkSheet sheet = workbook.WorkSheets.First();

    // Calculate sums from a cell range
    decimal totalSum = sheet["B2:B10"].Sum();

    // Alter a specific cell and apply styling
    sheet["B11"].Value = totalSum;
    sheet["B11"].Style.SetBackgroundColor("#808080")
                      .Font.SetColor("#ffffff");

    // Re-save and display updated Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

    DisplayAlert("Notification", "The Excel file has been updated successfully!", "OK");
}
```

This code snippet adds the capability to load an existing Excel file, compute a sum of specific cells, modify cell's value and style, and then save the updated file. A notification is displayed after these operations.

### Saving Excel Files

We need a service to handle the saving of our Excel files locally. Establish the `SaveService` class with its corresponding method in `SaveService.cs`:

```csharp
using System;
using System.IO;

namespace MAUI_IronXL
{
    public partial class SaveService
    {
        public partial void SaveAndView(string fileName, string contentType, MemoryStream stream);
    }
}

// Implement platform-specific functionality in appropriate files like SaveWindows.cs etc.
```

### Final Output

Compile and run the MAUI project. Upon successful operation, a window will display which guides you through the process of creating and modifying Excel files through the designed UI.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-1.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 1: Output" class="img-responsive add-shadow">
        <p><strong>Figure 1</strong> - <em>Output</em></p>
    </div>
</div>

#### Excel Creation and Reading Dialogs

When the "Create Excel File" button is clicked, a dialog prompts you to save the newly created file. After saving, another dialog confirms that the file is ready to be opened.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-2.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 2: Create Excel Popup" class="img-responsive add-shadow">
        <p><strong>Figure 2</strong> - <em>Create Excel Popup</em></p>
    </div>
</div>

When you follow the directions from the popup to open the created Excel file, the displayed document looks like the screenshot below.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-3.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 3: Output" class="img-responsive add-shadow">
        <p><strong>Figure 3</strong> - <em>Read and Modify Excel Popup</em></p>
    </div>
</div>

Clicking the "Read and Modify Excel File" button retrieves the previously generated file and updates it per the defined styles.

<div class="content-img-align-center">
    <div="