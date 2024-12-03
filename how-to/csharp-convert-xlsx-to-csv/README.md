# Converting Excel File Formats with IronXL in C&num;

***Based on <https://ironsoftware.com/how-to/csharp-convert-xlsx-to-csv/>***


IronXL is versatile in converting Excel files into several popular formats including JSON, CSV, XML, and older formats like XLS. This article walks you through the conversion process using IronXL, offering examples in XML, CSV, JSON, and also demonstrates how to export an Excel worksheet as a dataset.


---

### Step 1: Setting Up IronXL

Before you can begin converting Excel files, you need to install the IronXL library into your project. IronXL can be installed in two primary ways:

You can directly download it from [Iron Software's official Excel documentation](https://ironsoftware.com/csharp/excel/docs/).

Alternatively, you can install it using the NuGet Package Manager:

- Right-click on the solution name in Solution Explorer.
- Select 'Manage NuGet Packages'.
- Search for `IronXL.Excel`.
- Click 'Install'.

```shell
Install-Package IronXL.Excel
```

---

### How to Tutorial: Convert XLSX to Different Formats

With IronXL installed, you’re ready to transform Excel files into multiple formats.

Here’s how you can do it:

```cs
// Convert Excel document into multiple formats
using IronXL;

private void ExportExcelFile(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.Load("Sample_Excel_File.xlsx");
    workbook.Metadata.Title = "Sample_Excel_File";

    // Save in different formats
    workbook.SaveAs("Old_Excel_Format.xls");
    workbook.SaveAs("New_Excel_Format.xlsx");
    workbook.SaveAsCsv("Exported_File.csv");
    workbook.SaveAsJson("Exported_File.json");
    workbook.SaveAsXml("Exported_File.xml");

    // Export to DataSet for use in a DataGridView
    var dataSet = workbook.ToDataSet();
    dataGridView1.DataSource = dataSet;
    dataGridView1.DataMember = "Sheet1";
}
```

This sample code loads an ordinary XLSX file, sets its title, exports it in various formats, and finally populates a `DataGridView` from the resulting dataset.

### Visual Demonstrations of Exports

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/csv-file-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/csv-file-export.png" alt="CSV File Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 1</span> - <span class="image-description-text_italic">CSV File Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xml-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xml-export.png" alt="XML Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 2</span> - <span class="image-description-text_italic">XML Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/json-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/json-export.png" alt="JSON Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 3</span> - <span class="image-description-text_italic">JSON Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xls-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xls-export.png" alt="XLS Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 4</span> - <span class="image-description-text_italic">XLS Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
  <div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/excel-input-for-all-exports.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/excel-input-for-all-exports.png" alt="Excel Input for all exports">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 5</span> - <span class="image-description-text_italic">Excel Input for all exports</span>
    </div>
  </div>
</div>

---

### Library Quick Access

Explore more about merging, unmerging, and managing cells in Excel spreadsheets through the comprehensive IronXL API Reference Documentation. Navigate through the resource-rich documentation and expand your development skills!

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Documentation</h3>
      <p>Gain insights and extend your capabilities in handling Excel files with IronXL.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Explore IronXL API Docs <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>