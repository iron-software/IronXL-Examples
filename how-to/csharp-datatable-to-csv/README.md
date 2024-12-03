# Converting DataTable to CSV with C# and IronXL

***Based on <https://ironsoftware.com/how-to/csharp-datatable-to-csv/>***


This tutorial walks you through the steps for transforming a `DataTable` into a CSV file using IronXL. This is a straightforward process that we'll break down into easy-to-follow steps.

---

### Step 1: Install IronXL

To begin, you'll need to have IronXL installed in your project. IronXL offers multiple installation options to get started:

Visit the official site to download IronXL using this link: [IronXL Documentation](https://ironsoftware.com/csharp/excel/docs/)

Or, within Visual Studio:

- Navigate to the `Project` menu
- Choose `Manage NuGet Packages`
- Search for `IronXL.Excel`
- Select `Install`

```shell
Install-Package IronXL.Excel
```

<div align="center">
  <a
    href="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/ironxl-excel-nuget-package.png"
    target="_blank"
  >
    <img
      src="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/ironxl-excel-nuget-package.png"
      alt="IronXL.Excel NuGet Package"
      style="width:auto; max-width:100%; height:auto;"
    >
  </a>
  <div>
    <strong>Figure 1</strong> - IronXL.Excel NuGet Package
  </div>
</div>

---

### Step 2: Create and Export DataTable to CSV

Let's move on to coding.

Start by including the IronXL namespace in your project:

```csharp
using IronXL;
```

Proceed by applying the below C# code:

```csharp
// This method demonstrates how to save a DataTable to a CSV file.
private void button6_Click(object sender, EventArgs e)
{
    DataTable table = new DataTable();
    table.Columns.Add("Example_DataSet", typeof(string));
    
    // Adding dummy data to the DataTable
    for(int i = 0; i < 7; i++)
    {
        table.Rows.Add(i % 4); // Cycle through 0, 1, 2, 3, 1, 2, 3
    }

    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLS);
    workbook.Metadata.Author = "OJ"; // Setting author metadata
    WorkSheet worksheet = workbook.DefaultWorkSheet;

    int rowIndex = 1;
    foreach (DataRow row in table.Rows)
    {
        worksheet[$"A{rowIndex}"].Value = row[0].ToString();
        rowIndex++;
    }

    workbook.SaveAsCsv("Exported_DataTable_CSV.csv", ";"); // Specify delimiter
}
```

In this script, a `DataTable` is populated and subsequently exported into a CSV file through IronXL. The `SaveAsCsv` method streamlines this process.

### Visual Output

<div align="center">
  <a
    href="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/datatable-output-to-csv.png"
    target="_blank"
  >
    <img
      src="https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/datatable-output-to-csv.png"
      alt="Datatable output to CSV"
      style="width:auto; max-width:100%; height:auto;"
    >
  </a>
  <div>
    <strong>Figure 2</strong> - Datatable output to CSV
  </div>
</div>

---

### Library Quick Access

IronXL's API Reference Documentation provides extensive guides and samples for managing Excel interactions:

[Explore IronXL API Reference](https://ironsoftware.com/csharp/excel/object-reference/api/)

<div style="display:flex; align-items:center;">
  <img
    src="https://ironsoftware.com/img/svgs/documentation.svg"
    alt="Documentation"
    style="max-width: 110px; width: 100px; height: 140px; margin-right:20px;"
  >
  <strong>IronXL API Reference Documentation</strong>
</div>