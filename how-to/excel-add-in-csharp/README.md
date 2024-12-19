# C# Excel Integration (Code Example Tutorial)

***Based on <https://ironsoftware.com/how-to/excel-add-in-csharp/>***


Developing applications often requires the ability to manipulate Excel spreadsheets without the use of Excel itself. For instance, you might find it necessary to programmatically insert new rows or columns into an existing Excel spreadsheet. The C# "Excel: Add" functionality in IronXL enables you to do exactly this and much more. Below are detailed examples of how to implement these functions.

---

### Step 1: Install the IronXL Excel Library

To utilize the functionalities for adding rows and columns in Excel, you must initially download the IronXL Excel Library. It is available at no cost for development within your projects. You can [get the DLL directly here](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Add.Excel.Csharp.zip) or utilize the [NuGet package manager](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```

---

### Tutorial: Adding Rows and Columns in Excel

After the installation of IronXL, you can easily add new rows and columns to existing Excel spreadsheets through C#.

#### Adding a Row at the End of the Spreadsheet

Consider an Excel file named `sample.xlsx`, which contains 5 columns labeled from `A` to `E`. Below is how you can append a new row at the end:

```csharp
// Example: Adding a Row at the Last Position
using IronXL;

class Program
{
    static void Main(string[] args)
    {
        WorkBook wb = WorkBook.Load("sample.xlsx");
        WorkSheet ws = wb.GetWorkSheet("Sheet1");
        int rowIndex = ws.Rows.Count() + 1;  // Calculate the new row index
        string[] columns = {"A", "B", "C", "D", "E"};

        foreach (var col in columns)
        {
            ws[col + rowIndex].Value = "New Row";  // Set value for each column in the new row
        }

        wb.SaveAs("sample.xlsx");  // Save the changes
    }
}
```

This code adds a new row with the value `New Row` in each column to the `sample.xlsx` at the bottom.

#### Adding a Row at the Beginning of the Spreadsheet

Here's how to prepend a new row at the beginning of the spreadsheet:

```csharp
// Example: Adding a Row at the First Position
using IronXL;

class Program
{
    static void Main(string[] args)
    {
        WorkBook wb = WorkBook.Load("sample.xlsx");
        WorkSheet ws = wb.GetWorkSheet("Sheet1");
        ws.Rows.First().InsertRowsAbove(1);  // Insert a new row at the very top
        ws["A1:E1"].Value = "new row";  // Set values for the entire row

        wb.SaveAs("sample.xlsx");  // Save the changes
    }
}
```

This operation shifts existing rows down and sets all columns of the newly added top row to `new row`.

|Before|After|
|:---:|:---:|
|![Before](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/before2.png)|![After](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/after2.png)|

#### Adding a Column in Excel

Adding a new column is equally simple:

```csharp
// Example: Adding a Column
using IronXL;

class Program
{
    static void Main(string[] args)
    {
        WorkBook wb = WorkBook.Load("sample.xlsx");
        WorkSheet ws = wb.GetWorkSheet("Sheet1");
        ws.Columns.First().InsertColumnsBefore(1);  // Insert a new column before the first
        ws["A1:A" + ws.Rows.Count()].Value = "New Column Added";  // Fill the new column

        wb.SaveAs("sample.xlsx");  // Save workbook
    }
}
```

This shifts existing columns to the right and adds a new column at position `A` filled with `New Column Added`.

|Before|After|
|:---:|:---:|
|![Before](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/before1.png)|![After](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/after1.png)|

---

### Quick Access to Library Documentation

Explore more functions and further documentation on adding and manipulating rows, columns, and other Excel functionalities with C# through IronXL's extensive documentation.

[Read the IronXL Documentation](https://ironsoftware.com/csharp/excel/object-reference/api/) on detailed API references and guides.

![IronXL Documentation](https://ironsoftware.com/img/svgs/documentation.svg)