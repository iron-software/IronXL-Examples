# Creating, Reading, and Modifying Excel Files in .NET MAUI Applications

## Overview

*In this tutorial, we'll explore how to effectively handle Excel files within .NET MAUI applications for Windows, utilizing the IronXL library. Let’s dive in!*

## IronXL: The C# Excel Solution

IronXL serves as a robust C# .NET class library engineered for the manipulation, generation, and reading of Excel files. This library allows developers to fabricate Excel documents from the ground up, controlling aspects from content to aesthetics, and includes metadata capabilities like setting a document's title and author. Moreover, IronXL delivers extensive customization options for the user interface, including adjustments to margins, orientation, page size, and incorporation of images, among others. Importantly, IronXL operates independently without dependence on additional frameworks, platform integrations, or third-party libraries.

## Getting Started with IronXL

### Installation

To integrate IronXL into your project, utilize the NuGet Package Manager Console within Visual Studio. Simply execute the following command in the Console to install the IronXL library:

```shell
Install-Package IronXL.Excel
```

---

<h4>Step-by-Step Guide</h4>

## Building Excel Documents with IronXL in C#

### Designing the Application User Interface

Start by opening the `MainPage.xaml` file. Replace the existing code with the following XAML code to design the user interface.

```xml
<?xml version="1.0" encoding="utf-8"?>
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
                SemanticProperties.Description="Explore Multi-platform App UI"
                FontSize="18"
                HorizontalOptions="Center" />
            <Button
                x:Name="createBtn"
                Text="Create Excel File"
                SemanticProperties.Hint="Tap to create an Excel file"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />
            <Button
                x:Name="readExcel"
                Text="Read and Modify Excel File"
                SemanticProperties.Hint="Tap to read and modify an Excel file"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />
        </VerticalStackLayout>
    </ScrollView>
</ContentPage>
```

This XAML script sets up a straightforward interface for our .NET MAUI application, featuring a welcoming label and two interactive buttons—each dedicated to either creating or reading and modifying Excel files. These components are organized vertically ensuring a coherent display across different devices.

### Creating Excel Documents

Now, let’s create an Excel file. Open `MainPage.xaml.cs` and add the following C# method:

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Instantiate a new Workbook
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Add a Worksheet
    var sheet = workbook.CreateWorkSheet("2022 Budget");

    // Define Cell values for months
    string[] months = { "January", "February", "March", "April", "May", "June", 
                        "July", "August" };
    for (int i = 0; i < months.Length; i++)
    {
        sheet["A" + (i + 1)].Value = months[i];
    }

    // Dynamically set cell values
    Random rng = new Random();
    for (int row = 2; row <= 11; row++)
    {
        for (int col = 0; col < months.Length; col++)
        {
            sheet[$"{(char)('A' + col)}{row}"].Value = rng.Next(1, 8000);
        }
    }

    // Formatting cells with styles and borders
    var headerRange = sheet["A1:H1"];
    headerRange.Style.SetBackgroundColor("#d3d3d3").TopBorder.SetColor("#000000").BottomBorder.SetColor("#000000");
    var rowRange = sheet["H2:H11"];
    rowRange.Style.RightBorder.SetColor("#000000").Type = IronXL.Styles.BorderType.Medium;
    var bottomRange = sheet["A11:H11"];
    bottomRange.Style.BottomBorder.SetColor("#000000").Type = IronXL.Styles.BorderType.Medium;

    // Calculating and displaying results using formulas
    var calculations = new
    {
        Sum = sheet["A2:A11"].Sum(),
        Average = sheet["B2:B11"].Avg(),
        Max = sheet["C2:C11"].Max(),
        Min = sheet["D2:D11"].Min()
    };

    sheet["A12"].Value = "Sum";
    sheet["B12"].Value = calculations.Sum;
    sheet["C12"].Value = "Avg";
    sheet["D12"].Value = calculations.Average;
    sheet["E12"].Value = "Max";
    sheet["F12"].Value = calculations.Max;
    sheet["G12"].Value = "Min";
    sheet["H12"].Value = calculations.Min;

    // Save and view the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

This method demonstrates how to create a workbook, add and format a worksheet, populate it with data, apply styles, execute formulas, and ultimately, save and open the Excel file using a custom `SaveService` class.

## Introduction

*In this How-To Guide, we'll cover the steps to generate and manage Excel files within .NET MAUI applications for Windows, utilizing the capabilities of IronXL. Let’s dive in.*

## IronXL: The C# Library for Excel Operations

IronXL is a robust C# .NET library designed to handle Excel file operations. With this library, users can effortlessly generate Excel documents from the ground up, personalizing not only the content and aesthetic aspects but also the metadata like title and author. It offers a range of customization options for the user interface, including the ability to adjust margins, page orientation, size, and embed images. Importantly, IronXL operates independently without the need for any external frameworks or third-party libraries. This makes it a fully self-contained solution ideal for managing Excel files within .NET environments.

## Setting Up IronXL

To incorporate IronXL into your project, utilize the NuGet Package Manager Console within Visual Studio. Simply launch the Console and execute the command below to add the IronXL library to your application.

```shell
Install-Package IronXL.Excel
```

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">How To Guide</h4>

## Generating Excel Documents with C# Using IronXL

### Establishing the Application's Frontend

Begin by opening the `MainPage.xaml` file within your project and swap out the existing XML with the snippet provided below.

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
                SemanticProperties.Description="This is your starting point for a multi-platform app."
                FontSize="18"
                HorizontalOptions="Center" />

            <Button
                x:Name="createBtn"
                Text="Create Excel File"
                SemanticProperties.Hint="Tap here to start creating an Excel file"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />

            <Button
                x:Name="readExcel"
                Text="Read and Modify Excel file"
                SemanticProperties.Hint="Tap here to open and edit an Excel file"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />

        </VerticalStackLayout>
    </ScrollView>

</ContentPage>
```

This XML defines the interface for a simple .NET MAUI application. It includes a label greeting the user and two buttons—one for creating and another for reading and editing an Excel document, arranged in a vertical stack to maintain proper alignment across various devices.

### Creating Excel Documents

Now, let's construct an Excel document using IronXL. Navigate to your `MainPage.xaml.cs` file and implement the following code in it.

```cs
private void CreateExcel(object sender, EventArgs e)
{
    // Initialization of a new Workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Adding a new Worksheet
    var sheet = workbook.CreateWorkSheet("Annual Budget");

    // Assigning values to cells
    sheet["A1"].Value = "January";
    sheet["B1"].Value = "February";
    sheet["C1"].Value = "March";
    sheet["D1"].Value = "April";
    sheet["E1"].Value = "May";
    sheet["F1"].Value = "June";
    sheet["G1"].Value = "July";
    sheet["H1"].Value = "August";

    // Dynamically setting cell values
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

    // Applying formatting options
    sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Calculating and Displaying Results
    decimal sum = sheet["A2:A11"].Sum();
    decimal avg = sheet["B2:B11"].Avg();
    decimal max = sheet["C2:C11"].Max();
    decimal min = sheet["D2:A11"].Min();

    sheet["A12"].Value = "Sum";
    sheet["B12"].Value = sum;
    sheet["C12"].Value = "Avg";
    sheet["D12"].Value = avg;
    sheet["E12"].Value = "Max";
    sheet["F12"].Value = max;
    sheet["G12"].Value = "Min";
    sheet["H12"].Value = min;

    // Saving and Viewing the Excel File
    SaveService saveService = new SaveService();
    saveService.SaveAndView("AnnualBudget.xlsx", "application/octet-stream", workbook.ToStream());
}
```

The method above incrementally builds an Excel workbook. It sets cell values for each column header, fills subsequent cell rows with random financial data for illustrative purposes, applies specific styling to enhance readability, utilizes formulas to compute statistical values, and eventually saves and opens the Excel file through a predefined `SaveService` class.

### Constructing the Application User Interface

Begin by accessing the XAML page titled `**MainPage.xaml**`. Replace its existing content with the code provided below. This will structure the interface of your application.

Here's the paraphrased section of your article, with relative URL paths resolved to `ironsoftware.com`:

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
                SemanticProperties.Description="Greet .NET Multi-platform App UI"
                FontSize="18"
                HorizontalOptions="Center" />

            <Button
                x:Name="createBtn"
                Text="Create Excel File"
                SemanticProperties.Hint="Press this button to initiate Excel file creation"
                Clicked="CreateExcel"
                HorizontalOptions="Center" />

            <Button
                x:Name="readExcel"
                Text="Read and Modify Excel file"
                SemanticProperties.Hint="Press this button to access and alter an Excel file"
                Clicked="ReadExcel"
                HorizontalOptions="Center" />

        </VerticalStackLayout>
    </ScrollView>

</ContentPage>
```

The provided code sets up the foundational user interface of our simple .NET MAUI application. It includes a label and two buttons within a `VerticalStackLayout` container. The first button is designed to initiate the creation of an Excel file, while the second button is intended for opening and modifying an existing Excel file. This arrangement ensures that the elements are vertically aligned and consistent across all devices that support .NET MAUI.

### Excel File Creation

Now, let’s dive into generating an Excel file using IronXL. Navigate to the `MainPage.xaml.cs` file and implement the method displayed below.

Here's a paraphrased version of the provided C# function for creating an Excel file using IronXL in .NET MAUI:

```cs
private void CreateExcelFile(object sender, EventArgs args)
{
    // Initialize a new workbook
    WorkBook newWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);

    // Add a worksheet named '2022 Budget'
    var budgetSheet = newWorkbook.CreateWorkSheet("2022 Budget");

    // Define initial cell values for months
    budgetSheet["A1"].Value = "January";
    budgetSheet["B1"].Value = "February";
    budgetSheet["C1"].Value = "March";
    budgetSheet["D1"].Value = "April";
    budgetSheet["E1"].Value = "May";
    budgetSheet["F1"].Value = "June";
    budgetSheet["G1"].Value = "July";
    budgetSheet["H1"].Value = "August";

    // Populate cells dynamically with random financial data
    Random randomGenerator = new Random();
    for (int rowIndex = 2; rowIndex <= 11; rowIndex++)
    {
        budgetSheet["A" + rowIndex].Value = randomGenerator.Next(1, 1000);
        budgetSheet["B" + rowIndex].Value = randomGenerator.Next(1000, 2000);
        budgetSheet["C" + rowIndex].Value = randomGenerator.Next(2000, 3000);
        budgetSheet["D" + rowIndex].Value = randomGenerator.Next(3000, 4000);
        budgetSheet["E" + rowIndex].Value = randomGenerator.Next(4000, 5000);
        budgetSheet["F" + rowIndex].Value = randomGenerator.Next(5000, 6000);
        budgetSheet["G" + rowIndex].Value = randomGenerator.Next(6000, 7000);
        budgetSheet["H" + rowIndex].Value = randomGenerator.Next(7000, 8000);
    }

    // Formatting cells with styles and borders
    budgetSheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
    budgetSheet["A1:H1"].Style.TopBorder.SetColor("#000000");
    budgetSheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
    budgetSheet["H2:H11"].Style.RightBorder.SetColor("#000000");
    budgetSheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
    budgetSheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
    budgetSheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

    // Calculations and formulas
    decimal totalSum = budgetSheet["A2:A11"].Sum();
    decimal average = budgetSheet["B2:B11"].Avg();
    decimal maximum = budgetSheet["C2:C11"].Max();
    decimal minimum = budgetSheet["D2:D11"].Min();

    // Set formula results in the worksheet
    budgetSheet["A12"].Value = "Sum";
    budgetSheet["B12"].Value = totalSum;
    budgetSheet["C12"].Value = "Avg";
    budgetSheet["D12"].Value = average;
    budgetSheet["E12"].Value = "Max";
    budgetSheet["F12"].Value = maximum;
    budgetSheet["G12"].Value = "Min";
    budgetSheet["H12"].Value = minimum;

    // Save the Excel file and prompt for viewing
    SaveService fileSaver = new SaveService();
    fileSaver.SaveAndView("Budget.xlsx", "application/octet-stream", newWorkbook.ToStream());
}
```

This version retains the original function's logic while varying the structure and terminology to provide a fresh perspective on how to perform these operations with IronXL in a .NET MAUI application.

The provided source code utilizes IronXL to initialize a workbook that contains a single worksheet. It assigns values to individual cells through the `Value` property.

Through the style property, you can enhance the appearance of the cells with various styling options and borders. These enhancements can be applied to individual cells or collectively to a range of cells.

IronXL features support for Excel formulas, allowing the creation of customized formulas across one or several cells. The outcomes of these formulas can be captured in variables for subsequent use.

To facilitate the saving and viewing of the Excel files created, the `SaveService` class is employed. This class is introduced in the preceding sections and is detailed more extensively later in the document.

### View Excel Files in the Browser

To begin, access the `MainPage.xaml.cs` file and implement the code below.

```cs
private void ReadExcel(object sender, EventArgs e)
{
    // Define the file path
    string filepath="C:\\Files\\Customer Data.xlsx";
    WorkBook workbook = WorkBook.Load(filepath);
    WorkSheet sheet = workbook.WorkSheets.First();

    decimal sum = sheet ["B2:B10"].Sum();

    sheet ["B11"].Value = sum;
    sheet ["B11"].Style.SetBackgroundColor("#808080");
    sheet ["B11"].Style.Font.SetColor("#ffffff");

    // Save and display the Excel file
    SaveService saveService = new SaveService();
    saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

    DisplayAlert("Notification", "Excel file has been modified!", "OK");
}
```

This snippet demonstrates how to load an Excel file from a specified path, calculate a sum from a specified range of cells, and then both format and display the modified Excel file. It also presents a notification indicating that the file modifications are complete.

Here's the paraphrased section:

```cs
private void ModifyExcel(object sender, EventArgs e)
{
    // Define the file path
    string path = @"C:\Files\Customer Data.xlsx";
    WorkBook loadedWorkbook = WorkBook.Load(path);
    WorkSheet firstSheet = loadedWorkbook.WorkSheets.First();

    // Calculate the sum of values in a specific range
    decimal total = firstSheet["B2:B10"].Sum();

    // Update cell with the computed sum and adjust style
    firstSheet["B11"].Value = total;
    firstSheet["B11"].Style.SetBackgroundColor("#808080"); // Set the background color of the cell
    firstSheet["B11"].Style.Font.SetColor("#ffffff"); // Set the font color

    // Initialize the service to save and display the Excel file
    SaveService fileService = new SaveService();
    fileService.SaveAndView("Modified Data.xlsx", "application/octet-stream", loadedWorkbook.ToStream());

    // Show a notification indicating the modification
    DisplayAlert("Notification", "The Excel file has been successfully updated.", "OK");
}
```

This paraphrased code segment maintains the original logic and functionality while using slightly different wording and structure to describe the steps and actions performed.

The provided source code opens an Excel file and executes a formula across specific cells, additionally applying custom formatting to the background and text. Subsequently, it sends the Excel file as a byte stream to the user's browser. Moreover, it utilizes the `DisplayAlert` method to display a notification that confirms the file has been successfully opened and altered.

### Saving Excel Files

This section outlines the implementation of the `SaveService` class, initially mentioned earlier, which is responsible for saving our Excel files locally.

Begin by creating a new class file named `SaveService.cs` and populate it with the following code:

Here is the paraphrased section of your article with resolved paths from `ironsoftware.com`:

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
        // Method to save and view documents with filename, content type and data stream parameters
        public partial void SaveAndDisplay(string fileName, string contentType, MemoryStream stream);
    }
}
```

Subsequently, you will need to establish a class named `SaveWindows.cs`, which should be located within the `Platforms > Windows` directory. Please insert the following code into this class.

Here's the paraphrased section with resolved URL paths from links and images to ironsoftware.com:

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
            StorageFile savedFile;
            string fileExtension = Path.GetExtension(fileName);
            // Obtains the handle of the current process to display the dialog.
            IntPtr windowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (!Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons"))
            {
                // Initialize a file save picker dialog.
                FileSavePicker filePicker = new FileSavePicker();
                filePicker.DefaultFileExtension = ".xlsx";
                filePicker.SuggestedFileName = fileName;
                // Set the file type to Excel for saving.
                filePicker.FileTypeChoices.Add("XLSX", new List<string>() { ".xlsx" });

                // Associate the file picker with the current app window.
                WinRT.Interop.InitializeWithWindow.Initialize(filePicker, windowHandle);
                savedFile = await filePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder localFolder = ApplicationData.Current.LocalFolder;
                savedFile = await localFolder.CreateFileAsync(fileName, CreationCollisionOption.ReplaceExisting);
            }
            if (savedFile != null)
            {
                using (IRandomAccessStream fileStream = await savedFile.OpenAsync(FileAccessMode.ReadWrite))
                {
                    // Write the data from memory to the file.
                    using (Stream outputStream = fileStream.AsStreamForWrite())
                    {
                        outputStream.SetLength(0);
                        // Copy the stream data to a buffer, then save it to the file.
                        byte[] buffer = outputStream.ToArray();
                        outputStream.Write(buffer, 0, buffer.Length);
                        outputStream.Flush();
                    }
                }
                // Construct the message dialog for viewing the document.
                MessageDialog confirmationDialog = new("Do you want to view the document?", "File has been created successfully");
                UICommand yesCommand = new("Yes");
                confirmationDialog.Commands.Add(yesCommand);
                UICommand noCommand = new("No");
                confirmationDialog.Commands.Add(noCommand);

                // Attach the dialog to the main application window for proper UI interaction.
                WinRT.Interop.InitializeWithWindow.Initialize(confirmationDialog, windowHandle);

                // Display the dialog to the user.
                IUICommand selectedAction = await confirmationDialog.ShowAsync();
                if (selectedAction.Label == yesCommand.Label)
                {
                    // If the user chooses 'Yes', open the saved file.
                    await Windows.System.Launcher.LaunchFileAsync(savedFile);
                }
            }
        }
    }
}
```

### Output

Compile and execute the .NET MAUI project. Upon successful execution, you will see a window displaying the content shown in the image below.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-1.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 1: Output" class="img-responsive add-shadow">
        <p><strong>Figure 1</strong> - <em>Output</em></p>
    </div>
</div>

Pressing the "Create Excel File" button will trigger the display of a distinct dialog window, where users are asked to specify a location and a filename to save a newly created Excel file. Follow the prompts to enter the desired details and proceed by clicking OK. Subsequently, another dialog window will be presented.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-2.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 2: Create Excel Popup" class="img-responsive add-shadow">
        <p><strong>Figure 2</strong> - <em>Create Excel Popup</em></p>
    </div>
</div>

Opening the Excel file as instructed in the popup will display a document as depicted in the following screenshot.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-3.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 3: Output" class="img-responsive add-shadow">
        <p><strong>Figure 3</strong> - <em>Read and Modify Excel Popup</em></p>
    </div>
</div>

When you click the "Read and Modify Excel File" button, the application will open the Excel file created earlier and update it with the custom background and text colors specified previously.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-4.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 4: Excel Output" class="img-responsive add-shadow">
        <p><strong>Figure 4</strong> - <em>Excel Output</em></p>
    </div>
</div>

Upon opening the modified Excel file, the output displayed will include a table of contents as illustrated below.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="/img/tutorials/read-create-excel-net-maui/read-create-excel-net-maui-5.webp" alt="Read, Create, and Edit Excel Files in .NET MAUI, Figure 5: Modified Excel Output" class="img-responsive add-shadow">
        <p><strong>Figure 5</strong> - <em>Modified Excel Output</em></p>
    </div>
</div>

## Conclusion

This section has demonstrated how to use the IronXL library to create, read, and modify Excel files within a .NET MAUI application. IronXL excels in performing these tasks efficiently and accurately. It is a superior alternative to Microsoft Interop because it operates independently without needing Microsoft Office installed on the system. Furthermore, IronXL offers extensive functionality including the creation of workbooks and worksheets, managing cell ranges, applying formatting, and exporting to various file formats such as CSV and TSV.

IronXL is compatible with various project templates including Windows Forms, WPF, and ASP.NET Core. For more detailed guidance on utilizing IronXL, visit our tutorials on [creating Excel files](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/) and [reading Excel files](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/).

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

