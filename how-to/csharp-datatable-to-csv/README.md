# C# Tutorial: Convert DataTable to CSV with IronXL

***Based on <https://ironsoftware.com/how-to/csharp-datatable-to-csv/>***


This guide will illustrate how to transform a DataTable into a CSV file using IronXL, streamlining the process into simple, easy-to-follow steps.

---

### Step 1: Installing IronXL

To begin, install IronXL in your project. You can easily add IronXL in several ways:

- Download it directly from [IronXL's documentation page](https://ironsoftware.com/csharp/excel/docs/)
- Or from within Visual Studio:
  - Open the `Project` menu
  - Choose `Manage NuGet Packages`
  - Search for `IronXL.Excel` and then hit `Install`

```shell
Install-Package IronXL.Excel
```

<div align="center">
  <a href="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/ironxl-excel-nuget-package.png" target="_blank">
    <img src="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/ironxl-excel-nuget-package.png" alt="IronXL.Excel NuGet Package" class="img-responsive">
  </a>
  <div>
    <strong>Figure 1: IronXL.Excel NuGet Package</strong>
  </div>
</div>

---

### How to Tutorial: Exporting DataTable to CSV

Once IronXL is set up, you are ready to proceed.

First, include the IronXL namespace in your application:

```csharp
using IronXL;
```

Next, use the following code snippet to export a `DataTable` to CSV:

```csharp
// Method to convert DataTable to CSV
private void ConvertDataTableToCSV(object sender, EventArgs e)
{
    DataTable myTable = new DataTable();
    myTable.Columns.Add("Sample_Column", typeof(string));

    // Populating the DataTable
    for (int i = 0; i < 7; i++) {
        myTable.Rows.Add(i % 4);
    }

    // Creating a new Excel workbook
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLS);
    workbook.Metadata.Author = "Your Name";
    WorkSheet worksheet = workbook.DefaultWorkSheet;

    // Inserting data into Excel worksheet
    for (int rowIndex = 1; rowIndex <= myTable.Rows.Count; rowIndex++) {
        worksheet["A" + rowIndex].Value = myTable.Rows[rowIndex - 1][0].ToString();
    }

    // Exporting to CSV
    workbook.SaveAsCsv("Exported_DataTable.csv", ";"); // The file is saved as 'Exported_DataTable_Sheet1.csv'
}
```

This code initializes a `DataTable` and fills it with example data. A `WorkBook` instance is created, and each row from the `DataTable` is written into the workbook. Finally, the workbook is saved as a CSV file using the `SaveAsCsv` method.

<div align="center">
  <a href="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/datatable-output-to-csv.png" target="_blank">
    <img src="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/datatable-output-to-csv.png" alt="Datatable output to CSV" class="img-responsive">
  </a>
  <div>
    <strong>Figure 2: Datatable output to CSV</strong>
  </div>
</div>

---

#### Library Quick Access

<div class="row">
  <div class="col-sm-8">
    <h3>IronXL API Reference Documentation</h3>
    <p>Explore further and learn how to manage cells in Excel spreadsheets proficiently using IronXL's comprehensive API reference documentation.</p>
    <a href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Access IronXL API Reference Documentation <i class="fa fa-chevron-right"></i></a>
  </div>
  <div class="col-sm-4">
    <img style="max-width: 110px; width: 100%; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg">
  </div>
</div>