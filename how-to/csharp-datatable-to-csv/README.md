# Converting Datatable to CSV in C#

Easily transform your datatables into CSV format using IronXL with a few simple steps. This guide demonstrates a straightforward approach to accomplish this.

---

### Step 1

#### First, Install IronXL for Free

To begin, you must have IronXL installed. IronXL offers multiple installation options tailored to your project needs.

You can download it directly from their website through this link: [IronXL Installation Guide](https://ironsoftware.com/csharp/excel/docs/)

Alternatively, you can install it via NuGet in Visual Studio:

- Navigate to the Project menu
- Choose Manage NuGet Packages
- Look for `IronXL.Excel`
- Select Install

```shell
Install-Package IronXL.Excel
```

<center>

![IronXL.Excel NuGet Package](https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/ironxl-excel-nuget-package.png "IronXL.Excel NuGet Package")

</center>

---

### How-to Tutorial

#### Step 2: Create a Datatable and Export it as CSV

With IronXL installed, you're all set to proceed.

Begin by importing the IronXL namespace:

```cs
using IronXL;
```

Then, use this code to perform the conversion:

```cs
// Creates a datatable and exports it to a CSV file
private void button6_Click(object sender, EventArgs e)
{
    DataTable table = new DataTable();
    table.Columns.Add("Example_DataSet", typeof(string));
    // Adding sample data to the datatable
    table.Rows.Add("0");
    table.Rows.Add("1");
    table.Rows.Add("2");
    table.Rows.Add("3");
    table.Rows.Add("1");
    table.Rows.Add("2");
    table.Rows.Add("3");

    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
    wb.Metadata.Author = "OJ";
    WorkSheet ws = wb.DefaultWorkSheet;

    // Filling Excel worksheet with data from datatable
    int rowCount = 1;
    foreach (DataRow row in table.Rows)
    {
        ws["A" + rowCount].Value = row[0].ToString();
        rowCount++;
    }

    wb.SaveAsCsv("Save_DataTable_CSV.csv", ";"); // Specify delimiter if needed
}
```

This script first initializes a new datatable and populates it. It then creates a workbook and binds the datatable content to an Excel worksheet. Finally, it exports this data to a CSV file using the `SaveAsCsv` method.

#### Output CSV file:

<center>

![Datatable output to CSV](https://ironsoftware.com/img/faq/excel/csharp-datatable-to-csv/datatable-output-to-csv.png "Datatable output to CSV")

</center>

---

### Library Quick Access

#### Learn More Through IronXL API Reference Documentation

Explore more functionalities like merging, unmerging, and manipulating Excel cells by visiting the [IronXL API Reference Documentation](https://ironsoftware.com/csharp/excel/object-reference/api/).

![Documentation](https://ironsoftware.com/img/svgs/documentation.svg)

---