# Setting a Password for a Worksheet

***Based on <https://ironsoftware.com/how-to/set-password-worksheet/>***


Applying a **Read-Only** protection is frequently required to ensure data security in files. IronXL simplifies the process of imposing **Read-Only** protection on worksheets within .NET applications.

## Accessing a Password-Protected Worksheet

IronXL provides the capability to access and alter a protected worksheet without needing the original password. Upon opening the spreadsheet through IronXL, all the cells within any worksheet become editable.

## Enforcing Password Protection on a Worksheet

To prevent alterations to a worksheet while still permitting view access in Excel, the `ProtectSheet` method can be employed with a password argument. For instance, calling `workSheet.ProtectSheet("IronXL")` activates a password-based **ReadOnly** protection on the specified worksheet.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.SetPasswordWorksheet
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Applying password protection to the worksheet
            workSheet.ProtectSheet("IronXL");
            
            // Save changes to the workbook
            workBook.Save();
        }
    }
}
```

### Accessing a Worksheet That is Password-Protected

![Access Protected Worksheet](https://ironsoftware.com/static-assets/excel/how-to/set-password-worksheet/set-password-worksheet-access.gif)

## Removing a Password from a Worksheet

Utilizing the `UnprotectSheet` method allows you to strip away the password from a worksheet. By invoking `workSheet.UnprotectSheet()`, password protection can be lifted without requiring the original password.

```cs
using IronXL.Excel;
namespace ironxl.SetPasswordWorksheet
{
    public class Section2
    {
        public void Run()
        {
            // Clear password protection from the worksheet
            workSheet.UnprotectSheet();
        }
    }
}
```

IronXL enables straightforward protection and removal of security features on both Excel <a href="https://ironsoftware.com/csharp/excel/how-to/set-password-workbook/">workbooks</a> and **worksheets** via simple C# commands.