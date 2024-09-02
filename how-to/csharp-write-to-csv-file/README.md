# How to Generate CSV Files in .NET Using C#

Curious about how to swiftly generate CSV files using C#? IronXL has made this task both fast and simple for .NET developers.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Steps to Write CSV in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-add-ironxl-to-your-project">Integrate IronXL into Your Solution</a></li>
        <li><a href="#anchor-2-create-an-excel-workbook">Craft a Workbook Using C#</a></li>
        <li><a href="#anchor-3-save-workbook-to-csv">Convert and Save the Workbook as a CSV File</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Integrate IronXL into Your Solution

If IronXL is not yet part of your toolkit, here is how you can add it:

* Launch Visual Studio and go to the Project menu
* Choose Manage NuGet Packages
* Look for IronXL.Excel
* Click on Install

Alternatively, you can execute this command in the Developer Command Prompt:

```shell
Install-Package IronXL.Excel
```

For additional support, view our guides at [IronXL Documentation](https://ironsoftware.com/csharp/excel/docs/).

Download the sample project [here](https://ironsoftware.com/csharp/excel/downloads/csharp-write-to-csv.zip).

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Instructions</h4>

## 2. Craft a Workbook Using C#

Begin by creating an Excel workbook with this data:

<center>
  <div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/normal-excel-data-to-be-exported-to-csv.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/normal-excel-data-to-be-exported-to-csv.png" alt="Normal Excel data to be exported to CSV">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">
        Figure 1
      </span>
      <span class="image-description-text_regular">
        -
      </span>
      <span class="image-description-text_italic">
        Typical Excel data ready for CSV conversion
      </span>
    </div>
  </div>
</center>

Next, import the IronXL namespace to enable writing to CSV files using C#.

```cs
using IronXL;
```

<hr class="separator">

## 3. Convert and Save the Workbook as a CSV File

This section introduces a neat way to load and save your workbook using the IronXL library; it also appends the sheet name to the filename for clarity.

```cs
/**
Save as CSV File
**/
private void ExportToCSVButton_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.Load("Sample_Excel_File.xlsx");
    workbook.SaveAs("Exported_Excel_to_CSV.csv"); // The filename reflects the data source: Exported_Excel_to_CSV.Sheet1.csv
}
```

Here's how the exported CSV appears when viewed in a text editor like Notepad:

<center>
  <div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/output-csv-file.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/output-csv-file.png" alt="Output CSV file">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">
        Figure 2
      </span>
      <span class="image-description-text_regular">
        -
      </span>
      <span class="image-description-text_italic">
        Final CSV Output
      </span>
    </div>
  </div>
</center>

<hr class="separator">

<h4 class="tutorial-segment-title">Library Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Access the IronXL API Reference</h3>
      <p>Dive deeper into the capabilities of merging, unmerging, and managing cells through our comprehensive IronXL API Reference Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Visit IronXL API Reference <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>