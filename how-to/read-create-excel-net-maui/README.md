# Working with Excel Documents in .NET MAUI

***Based on <https://ironsoftware.com/how-to/read-create-excel-net-maui/>***


## Overview

*This step-by-step guide demonstrates how to craft and read Excel documents in .NET MAUI applications for Windows utilizing IronXL. Let’s dive in.*

## IronXL: Excel Handling in C#

IronXL is a robust C# library for .NET that facilitates the reading, writing, and manipulation of Excel files. It enables the creation of Excel sheets from the ground up, inclusive of content and visual styling, along with metadata like document titles and author details. It presents options for tweaking user interface elements such as margins, page sizes, orientations, and image embedding, all without the need for external frameworks, platforms, or third-party libraries. This library functions independently, ensuring seamless Excel document manipulation.

## Setting Up IronXL

Install the IronXL library through the NuGet Package Manager Console in Visual Studio. Simply open the Console and run the below command:

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">Step-by-Step Tutorial</h4>

## Developing Excel Applications in C# using IronXL

### Configuring the Application's Frontend

Begin by opening the XAML page titled `**MainPage.xaml**`. Replace the current code with the following lines:

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

The above XML layout defines the user interface for a simple .NET MAUI app, where a single label and two buttons are integrated. These controls allow for the creation and modification of Excel files and are organized vertically due to their placement within a `VerticalStackLayout`.

### Generating Excel Documents

Next, let's generate an Excel document using IronXL. Open `MainPage.xaml.cs` and implement the following method:

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Initialize a new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Add a Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Populate Cells
    sheet ["A1"].Value = "January";
    sheet ["B1"].Value = "February";
    sheet ["C1"].Value = "March";
    sheet ["D1"].Value = "April";
    sheet ["E1"].Value = "May";
    sheet ["F1"].Value = "June";
    sheet ["G1"].Value = "July";
    sheet ["H1"].Value = "August";

    // Dynamically set cell values
    Random r = new Random();
    for (int i = 2; i <= 11; i++)
    {
        sheet ["A" + i].Value = r.Next(1, 1000);
        sheet ["B" + i].Value = r.Next(1000, 2000);
        sheet ["C" + i].Value = r.Next(2000, 3000);
        sheet ["D" + i].Value = r.Next(3000, 4000);
        sheet ["E" + i].Value = r.Next(4000, 5000);
        sheet ["F" + i].Value = r.Next(5000, 6000);
        sheet ["G" + i].Value = r.Next(6000, 7000);
        sheet ["H" + i].Value = r.Next(7000, 8000);
    }

    // Style cells
    sheet ["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet ["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet ["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet ["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet ["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet ["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet ["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Implement Formulas
    decimal sum = sheet ["A2:A11"].Sum();
    decimal avg = sheet ["B2:B11"].Avg();
    decimal max = sheet ["C2:C11"].Max();
    decimal min = sheet ["D2:D11"].Min();

    sheet ["A12"].Value = "Sum";
    sheet ["B12"].Value = sum;

    sheet ["C12"].Value = "Avg";
    sheet ["D12"].Value = avg;

    sheet ["E12"].Value = "Max";
    sheet ["F12"].Value = max;

    sheet ["G12"].Value = "Min";
    sheet ["H12"].Value = min;

    // Save and Display Excel File
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

This method initializes a workbook, fills it with initial data and random values, applies styles, integrates calculation formulas, and finally uses a `SaveService` to save and display the file.

## Introduction

*Welcome to this instructional tutorial where we will learn how to generate and access Excel files in Windows-based .NET MAUI applications using IronXL. Let's dive in.*

## IronXL: The C# Excel Library for .NET

IronXL stands out as a comprehensive .NET library specifically designed for creating, reading, and editing Excel documents using C#. This library provides the flexibility to craft Excel files from the ground up, allowing full control over content, aesthetics, and document properties like titles and authors. Moreover, IronXL is equipped with extensive UI customization options, including the ability to adjust margins, orientation, and page size, and also to embed images. Remarkably, IronXL operates independently without the need for external frameworks, platform-specific integrations, or reliance on third-party libraries to produce Excel documents. Its standalone nature ensures simplicity and ease of integration into .NET applications.

## Installing IronXL

To integrate IronXL into your project, utilize the NuGet Package Manager Console within Visual Studio. Simply launch the console and execute the command below to add the IronXL library to your application.

```shell
Install-Package IronXL.Excel
```

Here's the paraphrased section with relative URL paths resolved:

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">How To Guide</h4>

## Generating Excel Documents with C# using IronXL

### Constructing the Interface

Begin by navigating to the XAML page named `**MainPage.xaml**` and update the markup with the following code:

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

This markup configures the user interface of your .NET MAUI application, integrating a label for a greeting message and two buttons designed for generating and modifying Excel documents. These components are conveniently arranged vertically using a `VerticalStackLayout`.

### Generating Excel Documents

Next, let's draft the Excel document using IronXL. Navigate to the `MainPage.xaml.cs` file and inject the following method:

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Initialize a new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Generate a new Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Initialize cells with values
    sheet["A1"].Value = "January";
    sheet["B1"].Value = "February";
    sheet["C1"].Value = "March";
    sheet["D1"].Value = "April";
    sheet["E1"].Value = "May";
    sheet["F1"].Value = "June";
    sheet["G1"].Value = "July";
    sheet["H1"].Value = "August";

    // Dynamically input values
    Random randomGenerator = new Random();
    for (int i = 2; i <= 11; i++)
    {
        sheet["A" + i].Value = randomGenerator.Next(1, 1000);
        sheet["B" + i].Value = randomGenerator.Next(1000, 2000);
        sheet["C" + i].Value = randomGenerator.Next(2000, 3000);
        sheet["D" + i].Value = randomGenerator.Next(3000, 4000);
        sheet["E" + i].Value = randomGenerator.Next(4000, 5000);
        sheet["F" + i].Value = randomGenerator.Next(5000, 6000);
        sheet["G" + i].Value = randomGenerator.Next(6000, 7000);
        sheet["H" + i].Value = randomGenerator.Next(7000, 8000);
    }

    // Styling cells
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Applying formulas
    decimal total = sheet["A2:A11"].Sum();
    decimal average = sheet["B2:B11"].Avg();
    decimal maximum = sheet["C2:C11"].Max();
    decimal minimum = sheet["D2:D11"].Min();

    // Displaying results in Excel
    sheet["A12"].Value = "Sum";
    sheet["B12"].Value = total;
    sheet["C12"].Value = "Avg";
    sheet["D12"].Value = average;
    sheet["E12"].Value = "Max";
    sheet["F12"].Value = maximum;
    sheet["G12"].Value = "Min";
    sheet["H12"].Value = minimum;

    // Save and Display the Excel File
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

In this method, a new workbook and a worksheet titled "2022 Budget" are devised, populated with dynamic values and styled. IronXL is leveraged to apply formulas directly in the worksheet, simplifying operations like sum, average, maximum, and minimum. The result is visualized immediately in Excel through custom setup dialogs, enhancing both functionality and user interaction.

### Configure the Application's Frontend Interface

Start by launching the XAML file titled `**MainPage.xaml**`. Proceed by substituting its existing code with the snippet provided below.

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
                SemanticProperties.Description="Introduction to Multi-platform Application Interface"
                FontSize="18"
                HorizontalOptions="Center" />

            <Button
                x:Name="createBtn"
                Text="Create Excel File"
                SemanticProperties.Hint="Activate this button to generate an Excel file"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />

            <Button
                x:Name="readExcel"
                Text="Read and Modify Excel File"
                SemanticProperties.Hint="Activate this button to access and modify an Excel file"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />

        </VerticalStackLayout>
    </ScrollView>

</ContentPage>
```

The provided code snippet is used to design the interface of a basic .NET MAUI application. It arranges a label along with two buttons within the interface. The first button is designated for generating an Excel file, while the second button is tasked with reading and adjusting an existing Excel file. These interface components are neatly organized into a `VerticalStackLayout`, ensuring that they are displayed in a vertical order across all compatible devices.

### Creating Excel Documents

Let's now proceed to generate an Excel document leveraging IronXL. Start by navigating to the `MainPage.xaml.cs` file and implement the method outlined below.

Here's the paraphrased section:

```cs
private void GenerateExcelFile(object sender, EventArgs e)
{
    // Initialize a new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Create a new Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Initialize Cell values for the first row
    sheet["A1"].Value = "January";
    sheet["B1"].Value = "February";
    sheet["C1"].Value = "March";
    sheet["D1"].Value = "April";
    sheet["E1"].Value = "May";
    sheet["F1"].Value = "June";
    sheet["G1"].Value = "July";
    sheet["H1"].Value = "August";

    // Dynamically generate values for cells using random numbers
    Random random = new Random();
    for (int rowIndex = 2; rowIndex <= 11; rowIndex++)
    {
        sheet["A" + rowIndex].Value = random.Next(1, 1000);
        sheet["B" + rowIndex].Value = random.Next(1000, 2000);
        sheet["C" + rowIndex].Value = random.Next(2000, 3000);
        sheet["D" + rowIndex].Value = random.Next(3000, 4000);
        sheet["E" + rowIndex].Value = random.Next(4000, 5000);
        sheet["F" + rowIndex].Value = random.Next(5000, 6000);
        sheet["G" + rowIndex].Value = random.Next(6000, 7000);
        sheet["H" + rowIndex].Value = random.Next(7000, 8000);
    }

    // Format the cells with backgrounds and borders
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Calculate and apply Excel formulas for statistical results
    decimal total = sheet["A2:A11"].Sum();
    decimal average = sheet["B2:B11"].Avg();
    decimal maximum = sheet["C2:C11"].Max();
    decimal minimum = sheet["D2:D11"].Min();

    sheet["A12"].Value = "Sum";
    sheet["B12"].Value = total;
    sheet["C12"].Value = "Avg";
    sheet["D12"].Value = average;
    sheet["E12"].Value = "Max";
    sheet["F12"].Value = maximum;
    sheet["G12"].Value = "Min";
    sheet["H12"].Value = minimum;

    // Save and Display the Excel File
    SaveService savingService = new SaveService();
    savingService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

The provided code snippet utilizes IronXL to generate a workbook that includes a single worksheet. It utilizes the property `Value` to set the values for individual cells. 

Styling and borders can effectively enhance the appearance of cells. The style attribute in the code facilitates the addition of such enhancements either to individual cells or to a collection of cells concurrently.

IronXL offers robust support for deploying Excel formulas. You can integrate custom formulas into one or several cells. Importantly, the outcomes of these formulas can be captured in variables for subsequent use.

To manage the storage and display of created Excel files, the `SaveService` class is employed. This class, mentioned earlier in the text, will be elaborately defined in subsequent sections of the document.

### Presenting Excel Files in the Browser

Navigate to the `MainPage.xaml.cs` file and insert the code provided below:

```cs
private void ReadExcel(object sender, EventArgs e)
{
    // Define the file path
    string filepath="C:\\Files\\Customer Data.xlsx";
    WorkBook workbook = WorkBook.Load(filepath);
    WorkSheet sheet = workbook.WorkSheets.First();

    // Execute formula
    decimal sum = sheet["B2:B10"].Sum();

    // Update cell value and style
    sheet["B11"].Value = sum;
    sheet["B11"].Style.SetBackgroundColor("#808080");
    sheet["B11"].Style.Font.SetColor("#ffffff");

    // Save and view the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

    // Display a notification alert
    DisplayAlert("Notification", "Excel file has been modified!", "OK");
}
```

This code snippet loads an existing Excel file, calculates the sum of a range of cells using a formula, and styles the resulting cell. The styled Excel file is then saved and streamed to the user's browser. Additionally, a notification alert informs the user once the file has been opened and modified.

Below is the paraphrased section of the article with resolved relative URL paths:

```cs
private void ReadExcel(object sender, EventArgs e)
{
    // Define the path where the file is located
    string pathToFile = @"C:\Files\Customer Data.xlsx";
    WorkBook workbook = WorkBook.Load(pathToFile);
    WorkSheet sheet = workbook.WorkSheets.First();

    // Calculate the sum of values in a specified range
    decimal total = sheet ["B2:B10"].Sum();

    // Update the cell with the calculated sum and apply styling
    sheet ["B11"].Value = total;
    sheet ["B11"].Style.SetBackgroundColor("#808080"); // Set a grey background
    sheet ["B11"].Style.Font.SetColor("#ffffff"); // Set font color to white

    // Initialize the service to manage file saving and viewing
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

    // Display a notification after modification
    DisplayAlert("Alert", "The Excel file has been successfully updated!", "OK");
}
```

The source code reads an Excel file, performs computations on a specified range of cells, and customizes their appearance with specified background and font colors. Subsequently, the modified Excel file is converted into a byte stream which is then sent to the user's browser for download. Moreover, the `DisplayAlert` method is utilized to show a notification that confirms the modifications have been applied to the file and it is ready to be viewed.

### Saving Excel Files

This segment outlines the setup of the `SaveService` class, which was mentioned previously and assists in saving Excel files locally.

Proceed by creating a class file named `SaveService.cs` and input the following code:

```csharp
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

Here's the paraphrased section, with updated markdown formatting and any relative URL paths resolved to `ironsoftware.com`:

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
        public partial void SaveAndDisplay(string fileName, string mimeType, MemoryStream stream);
    }
}
```

After that, proceed to establish a class titled `SaveWindows.cs` within the Platforms > Windows directory, and insert the following code as illustrated:

Here is the paraphrased section of your article, with relative URL paths resolved to ironsoftware.com:

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
            // Retrieves the handle of the current process window for dialogue initialization.
            IntPtr windowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (!Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons"))
            {
                // Initiates a file saver dialogue to store files.
                FileSavePicker savePicker = new FileSavePicker
                {
                    DefaultFileExtension = ".xlsx",
                    SuggestedFileName = fileName,
                    FileTypeChoices = { { "XLSX", new List<string> { ".xlsx" } } }
                };

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
                    // Directly writes data from memory to file stream.
                    using (Stream outstream = zipStream.AsStreamForWrite())
                    {
                        outstream.SetLength(0);  // Clear existing contents.
                        byte[] buffer = stream.ToArray(); // Temporarily holds data.
                        outstream.Write(buffer, 0, buffer.Length);
                        outstream.Flush(); // Ensure all data is written to the file.
                    }
                }
                // Prompt to check if the user wants to view the newly created file.
                MessageDialog msgDialog = new MessageDialog("Do you want to view the document?", "File has been created successfully");
                UICommand yesCmd = new UICommand("Yes");
                UICommand noCmd = new UICommand("No");
                msgDialog.Commands.Add(yesCmd);
                msgDialog.Commands.Add(noCmd);

                WinRT.Interop.InitializeWithWindow.Initialize(msgDialog, windowHandle);

                // Display and handle the response from the dialog.
                IUICommand cmd = await msgDialog.ShowAsync();
                if (cmd == yesCmd)
                {
                    // Opens the saved file.
                    await Windows.System.Launcher.LaunchFileAsync(stFile);
                }
            }
        }
    }
}
```

### Result Display

After compiling and executing the MAUI project, a window will appear showcasing the output as illustrated in the following image.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-1.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 1: Output" class="img-responsive add-shadow">
        <p><strong>Figure 1</strong> - <em>Output</em></p>
    </div>
</div>

When you press the "Create Excel File" button, a new dialog window will appear. This dialog will prompt you to select a location and a filename for saving the newly created Excel file. Follow the instructions to choose the appropriate location and filename, then click OK. Subsequently, a second dialog window will be displayed.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-2.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 2: Create Excel Popup" class="img-responsive add-shadow">
        <p><strong>Figure 2</strong> - <em>Create Excel Popup</em></p>
    </div>
</div>

Following the instructions from the popup to open the Excel file will display the document as depicted in the screenshot below.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-3.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 3: Output" class="img-responsive add-shadow">
        <p><strong>Figure 3</strong> - <em>Read and Modify Excel Popup</em></p>
    </div>
</div>

Upon selecting the "Read and Modify Excel File" button, the application will access the previously generated Excel document and update it with the predefined custom background and text coloring specified before.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-4.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 4: Excel Output" class="img-responsive add-shadow">
        <p><strong>Figure 4</strong> - <em>Excel Output</em></p>
    </div>
</div>

Upon opening the adjusted file, the display will present the subsequent output, complete with a table of contents.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-5.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 5: Modified Excel Output" class="img-responsive add-shadow">
        <p><strong>Figure 5</strong> - <em>Modified Excel Output</em></p>
    </div>
</div>

## Conclusion

This section has illustrated how the IronXL library can be effectively used to create, read, and modify Excel files within .NET MAUI applications. Notably, IronXL is recognized for its rapid and precise execution of tasks. It is a superior choice for managing Excel data operations compared to Microsoft Interop, mainly because it does not necessitate the installation of Microsoft Office on your device. Moreover, IronXL facilitates a broad range of functionalities, including but not limited to, the creation of workbooks and worksheets, managing cell data and formats, and exporting contents to various file formats like CSV and TSV.

Additionally, IronXL is compatible with diverse project formats including Windows Form, WPF, ASP.NET Core, among others. For further guidance on utilizing IronXL, explore our detailed tutorials on [creating Excel files](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/) and [reading Excel files](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/).

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

