# Transforming XLSX to CSV, JSON, XML, and More Using C&num;

IronXL enables the conversion of Excel files into various file formats.

These formats encompass JSON, CSV, XML, and even the older Excel file type XLS.

The following guide illustrates how you can leverage IronXL to convert files into XML, CSV, JSON; additionally, it includes a demonstration of exporting an Excel worksheet as a dataset.

---

### Step 1: IronXL Library Installation

To utilize IronXL in your projects, you must first install it. You can install IronXL via two methods:

Download from: [IronXL Documentation](https://ironsoftware.com/csharp/excel/docs/)

Or through Nuget Package Manager:

- Right-click the Solution name in Solution Explorer.
- Select Manage NuGet Packages.
- Search for `IronXL.Excel`.
- Click Install.

```shell
Install-Package IronXL.Excel
```

---

### Tutorial on Converting File Formats

With IronXL installed, you're set to begin converting your files.

Incorporate the following code snippet:

```cs
/**
Convert various formats
anchor-convert-to-xml-json-csv-xls
**/
using IronXL;

private void ConvertButton_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.Load("Standard_Excel_File.xlsx");

    workbook.Metadata.Title = "Standard_Excel_File.xlsx";

    workbook.SaveAs("Converted_XLS.xls");
    workbook.SaveAs("Converted_XLSX.xlsx");
    workbook.SaveAsCsv("Converted_CSV.csv");
    workbook.SaveAsJson("Converted_JSON.json");
    workbook.SaveAsXml("Converted_XML.xml");

    System.Data.DataSet dataSet = workbook.ToDataSet();

    dataGridView.DataSource = dataSet;
    dataGridView.DataMember = "Sheet1";
}
```

This script loads a standard XLSX file, assigns a Title, converts it to multiple formats, and finally exports the worksheet as a DataSet used by a DataGridView component.

Below are the exported files rendered in various formats.

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/csv-file-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/csv-file-export.png" alt="CSV File Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 1 - </span>
      <span class="image-description-text_italic">CSV File Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xml-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xml-export.png" alt="XML Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 2 - </span>
      <span class="image-description-text_italic">XML Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
	<div the="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/json-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/json-export.png" alt="JSON Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 3 - </span>
      <span class="image-description-text_italic">JSON Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xls-export.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xls-export.png" alt="XLS Export">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 4 - </span>
      <span class="image-description-text_italic">XLS Export</span>
    </div>
  </div>
</div>

<div class="content-img-align-center">
  <div class="center-image-wrapper">
    <a href="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/excel-input-for-all-exports.png" target="_blank">
      <img class="img-responsive" src="https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/excel-input-for-all-exports.png" alt="Excel Input for all exports">
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">Figure 5 - </span>
      <span class="image-description-text_italic">Excel Input for all exports</span>
    </div>
  </div>
</div>

---

### Library Quick Access

Explore more and share your insights on manipulating and interacting with Excel spreadsheets at the IronXL API Reference Documentation.

[Learn More at IronXL API Reference Documentation](https://ironsoftware.com/csharp/excel/object-reference/api/)
<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Documentation</h3>
      <p>Dive deep into merging, unmerging, and working with cells in Excel sheets using the comprehensive IronXL API Reference.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">IronXL API Reference Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>