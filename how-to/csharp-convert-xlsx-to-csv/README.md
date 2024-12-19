# Converting XLSX to CSV, JSON, XML, and more using C&num; and IronXL

***Based on <https://ironsoftware.com/how-to/csharp-convert-xlsx-to-csv/>***


IronXL provides a versatile way to convert Excel files into a variety of formats including JSON, CSV, XML, and even the older Excel format such as XLS. This article will guide you on how to utilize IronXL for converting Excel documents to XML, CSV, JSON, and will also demonstrate how to convert an Excel worksheet into a dataset.

---

## Step 1: Installing the IronXL Library

To begin using IronXL in your projects, it must first be installed. You can install IronXL via two primary methods:

- **Download Directly:** 
  - Visit [IronXL Excel Documentation](https://ironsoftware.com/csharp/excel/docs/)

- **Using NuGet Package Manager:** 
  - Right-click on the Solution name in the Solution Explorer.
  - Select **Manage NuGet Packages**.
  - Search for `IronXL.Excel` and click **Install**.

```shell
Install-Package IronXL.Excel
```

---

## How-to Guide: Conversion to XML, JSON, CSV, XLS

Once IronXL is installed, you can start converting files. Below is an example demonstrating how to convert an Excel file to various formats and ultimately, how to represent an Excel sheet as a `DataSet` which can be used in a `DataGridView`.

```cs
// Convert Excel file to multiple formats including XML, JSON, CSV, and XLS
using IronXL;

private void button7_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.Load("Normal_Excel_File.xlsx");

    // Assign a title to the workbook metadata
    workbook.Metadata.Title = "Normal Excel File Conversion";

    // Save in different formats
    workbook.SaveAs("Old_Excel_Format.xls");
    workbook.SaveAs("Updated_Excel_Format.xlsx");
    workbook.SaveAsCsv("Exported_CSV_File.csv");
    workbook.SaveAsJson("Exported_JSON_File.json");
    workbook.SaveAsXml("Exported_XML_File.xml");

    // Convert to DataSet and bind it to DataGridView
    System.Data.DataSet dataSet = workbook.ToDataSet();
    dataGridView1.DataSource = dataSet;
    dataGridView1.DataMember = "Sheet1";
}
```

The above code loads an ordinary XLSX file, modifies its title metadata, then saves it in various formats. It also demonstrates how to export the worksheet data into a `DataSet`.

### Visual Representations of Export Files

Enjoy a visual representation of each file format exported:

- ![CSV File Export](https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/csv-file-export.png "CSV File Export")
- ![XML Export](https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xml-export.png "XML Export")
- ![JSON Export](https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/json-export.png "JSON Export")
- ![XLS Export](https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/xls-export.png "XLS Export")
- ![Excel Input Used For All Exports](https://ironsoftware.com/img/faq/excel/csharp-convert-xlsx-to-csv/excel-input-for-all-exports.png "Excel Input for all exports")

---

## Quick Access to Library Documentation

To learn more about working with Excel functionalities like merging, unmerging, and manipulating cells using IronXL, refer to:

- [IronXL API Reference Documentation](https://ironsoftware.com/csharp/excel/object-reference/api/ "IronXL API Reference Documentation")
  ![IronXL Document](https://ironsoftware.com/img/svgs/documentation.svg "IronXL API Documentation")