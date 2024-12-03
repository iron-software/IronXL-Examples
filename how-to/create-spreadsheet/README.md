# How to Generate New Excel Files

***Based on <https://ironsoftware.com/how-to/create-spreadsheet/>***


XLSX is a contemporary file format for storing Microsoft Excel spreadsheets. This format, which adheres to the Open XML standard, was introduced with Office 2007. XLSX is capable of supporting sophisticated features such as charts and conditional formatting, making it ideal for data analysis and business-related tasks.

Conversely, XLS is the older, binary format used in former versions of Excel. It does not support the extensive features found in XLSX and has become increasingly rare.

IronXL enables developers to generate both XLSX and XLS files effortlessly with a single line of code.

## Example: Creating a Spreadsheet

To create an Excel workbook, which can house a set of sheets or worksheets, use the static `Create` method from IronXL. By default, this method generates a workbook in the XLSX format.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CreateSpreadsheet
{
    public class Section1
    {
        public void Run()
        {
            // Initialize a new workbook
            WorkBook workBook = WorkBook.Create();
        }
    }
}
```

---

## Selecting the Spreadsheet Format

The `Create` method includes an option to specify the desired format of the Excel file using the **ExcelFileFormat** enum, allowing the choice between the more modern, XML-based XLSX format and the older, binary XLS format. While XLSX is preferred for its advanced features and efficiency, XLS remains an option for compatibility with older systems.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CreateSpreadsheet
{
    public class Section2
    {
        public void Run()
        {
            // Generate an XLSX file
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}
```

There is also an overloaded version of the `Create` method, which accepts a **CreatingOptions** object. This parameter currently holds a single property, DefaultFileFormat, which specifies whether the created file should be in XLSX or XLS format. Below is a sample usage:

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.CreateSpreadsheet
{
    public class Section3
    {
        public void Run()
        {
            // Generate an XLSX file
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}
```