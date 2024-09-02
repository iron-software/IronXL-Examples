# How to Add Password Protection to a Worksheet

Adding a **Read-Only** mode to a worksheet is frequently necessary for data files. IronXL simplifies the process of enforcing **Read-Only** protection on worksheets within .NET applications.

## Accessing a Password-Protected Worksheet

IronXL provides the capability to access and edit any password-protected worksheet without needing the actual password. Once you open the spreadsheet using IronXL, you can change any cell within any of the worksheets.

## Implementing Password Protection on a Worksheet

To prevent changes to a worksheet while still permitting users to view its contents in Excel, utilize the `ProtectSheet` method and include a password as an argument. For instance, using the code `workSheet.ProtectSheet("IronXL")` establishes a password-protected **ReadOnly** mode for the chosen worksheet.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Applying protection to the worksheet
workSheet.ProtectSheet("IronXL");

workBook.Save();
```

### Opening a Worksheet with Password Protection

<img src="https://ironsoftware.com/static-assets/excel/how-to/set-password-worksheet/set-password-worksheet-access.gif" alt="Access Protected Worksheet" class="img-responsive add-shadow" style="margin-bottom: 30px;"/>

## Disabling Password Protection from a Worksheet

To eliminate the password protection from a particular worksheet, utilize the `UnprotectSheet` method. By simply invoking `workSheet.UnprotectSheet()`, you can lift the password restriction from that worksheet.

```cs
// Disabling worksheet protection. Password is not required!
workSheet.UnprotectSheet();
```

IronXL facilitates both the protection and unprotection of any Excel <a href="https://ironsoftware.com/csharp/excel/how-to/set-password-workbook/">workbook</a> and **worksheet** effortlessly using just a line of C# code.