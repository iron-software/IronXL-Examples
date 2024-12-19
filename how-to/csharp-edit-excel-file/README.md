# C# Edit Excel File

***Based on <https://ironsoftware.com/how-to/csharp-edit-excel-file/>***


When working with Excel files in C#, developers should proceed with caution to avoid unwanted modifications to the document. Using streamlined and robust lines of code not only minimizes the risk of error but also simplifies the task of programmatically editing or deleting Excel files. In this guide, we'll demonstrate how to effectively edit Excel files in C# by leveraging trusted functions.

---

### Step 1

## Editing Excel Files in C# with the IronXL Library

For this tutorial, we're using the functionalities provided by IronXL, a comprehensive C# library for handling Excel files. First, you need to install IronXL into your project, which is free for development purposes.

You can [download IronXL.zip](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Edit.Excel.Csharp.zip) directly or learn more and install it via the [NuGet package page](https://www.nuget.org/packages/IronXL.Excel).

After installation, let's set up your environment:

```shell
Install-Package IronXL.Excel
```

---

### How to Tutorial

## 2. Modify Specific Cell Values

To begin, import the Excel spreadsheet you wish to modify and access its worksheet as shown below:

```cs
// Import and Modify Spreadsheet
// anchor-edit-specific-cell-values
using IronXL;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx"); // Import the spreadsheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1"); // Access the desired worksheet
    ws.Rows[3].Columns[1].Value = "New Value"; // Modify a specific cell
    wb.SaveAs("sample.xlsx"); // Save the modifications
}
```

Observe the modifications below, illustrated by before-and-after screenshots of the `sample.xlsx` file:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before1.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after1.png)|

Alternatively, access and modify cells directly by their address:

```cs
ws["B4"].Value = "New Value"; // Alternate method to modify a specific cell
```

---

## 3. Edit Entire Row Values

Editing entire rows in an Excel spreadsheet is straightforward:

```cs
// Edit Full Row Values
// anchor-edit-full-row-values
using IronXL;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Rows[3].Value = "New Value"; // Modify the entire row
    wb.SaveAs("sample.xlsx");
}
```

View the changes in `sample.xlsx` below:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before2.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after2.png)|

To change specific ranges within a row:

```cs
ws["A3:E3"].Value = "New Value";
```

---

## 4. Amend Full Column Values

Similarly to editing rows, columns can be modified with ease:

```cs
// Adjust Full Column Values
// anchor-edit-full-column-values
using IronXL;

static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Columns[1].Value = "New Value"; // Modify the entire column
    wb.SaveAs("sample.xlsx");
}
```

The effect on `sample.xlsx` after the change:

|Before|After|
|:---:|:---:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before4.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_After4.png)|

---

## 5. Modify Rows with Dynamic Values

Use IronXL to edit specific rows dynamically, assigning unique values to each cell:

```cs
// Modify Row with Dynamic Values
// anchor-edit-full-row-with-dynamic-values
using IronXL;

static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    for (int i = 0; i < ws.Columns.Count(); i++)
    {
        ws.Rows[3].Columns[i].Value = "New Value " + i;
    }
    wb.SaveAs("sample.xlsx");
}
```

Below are the visual results for this operation on `sample.xlsx`:

|Before|After|
|:---:|:---:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before3.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after3.png)|

---

## 6. Edit Columns with Dynamic Values

Editing specific columns dynamically is also straightforward:

```cs
// Modify Column with Dynamic Values
// anchor-edit-full-column-with-dynamic-values
using IronXL;

static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    for (int i = 0; i < ws.Rows.Count(); i++)
    {
        if (i == 0) // Skip if the first column is a header
            continue;
        ws.Rows[i].Columns[1].Value = "New Value " + i;
    }
    wb.SaveAs("sample.xlsx");
}
```

See the resulting changes in the table:

|Before|After|
|:---:|:---:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before5.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after5.png)|

---

## 7. Update Spreadsheet Values

IronXL allows for versatile replacements across worksheets. For example, updating specific or all values with new entries:

### 7.1. Replace Particular Value Throughout a Worksheet

To replace a specific worksheet value:

```cs
/**
Replace Cell Values
anchor-replace-specific-value-of-complete-worksheet
**/
ws.Replace("old value", "new value"); // Replace old value with new value across the worksheet
```

Remember to save your file after each modification.

### 7.2. Update Specific Row Values

To modify only a certain row:

```cs
ws.Rows[2].Replace("old value", "new value"); // Update specific row
```

### 7.3. Adjust Values in Row Ranges

Replace values in a row range:

```cs
ws["B4:E4"].Replace("old value", "new value"); // Update values within a defined range
```

### 7.4. Update Specific Column Values

Replace column-specific values:

```cs
ws.Columns[1].Replace("old value", "new Value"); // Update specific column
```

### 7.5. Modify Values in Column Ranges

Update values within a range in a column:

```cs
ws["B5:B10"].Replace("old value", "new value");
```

---

## 8. Remove Row from an Excel Worksheet

Removing a specific row is effortlessly handled by IronXL:

```cs
// Remove Specified Row
// anchor-remove-row-from-excel-worksheet
using IronXL;
static void Main(string [] args)
{ 
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Rows[3].RemoveRow(); // Remove a particular row
    wb.SaveAs("sample.xlsx");
}
```

The following table illustrates the row removal:

|Before|After|
|:---:|:---:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before6.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after6.png)|

---

## 9. Exclude a Worksheet from an Excel File

To remove a complete worksheet from an Excel file, utilize the following method:

```cs
// Remove Worksheet from Excel File
wb.RemoveWorkSheet(1); // Remove by worksheet index
wb.RemoveWorkSheet("Sheet1"); // Remove by worksheet name
```

IronXL provides a multitude of functions for efficiently performing any sort of modification or deletion in Excel files. If you need further assistance or have any questions, don't hesitate to contact our development team.

---

### Library Quick Access

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore IronXL Library Documentation</h3>
      <p>Dive into the extensive features of the IronXL C# Library with various functions for editing, deleting, styling, and enhancing your Excel workbooks.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> IronXL Library Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100%;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </>
    </div>
  </div>
</div>

---