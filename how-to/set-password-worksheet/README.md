# How to Set a Password on a Worksheet

***Based on <https://ironsoftware.com/how-to/set-password-worksheet/>***


Setting a worksheet to **Read-Only** is a frequently required feature for data files. IronXL simplifies the process of applying **Read-Only** protection to worksheets in .NET applications.

### Initial Setup with IronXL

---

## Accessing a Password Protected Worksheet

IronXL enables the modification and access of any protected worksheet without the need for a password. Once you open the spreadsheet through IronXL, you're free to alter any cell across the worksheets.

## Implementing Password Protection on a Worksheet

To prevent changes to a worksheet while still permitting viewing access in Excel, employ the `ProtectSheet` method alongside a password argument. For instance, `workSheet.ProtectSheet("IronXL")` establishes password-based **ReadOnly** protection for the targeted worksheet.

```cs
using IronXL;

// Load an existing workbook
WorkBook workBook = WorkBook.Load("sample.xlsx");
// Access the default worksheet
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Applying ReadOnly protection with a password
workSheet.ProtectSheet("IronXL");

// Save the changes to the workbook
workBook.Save();
```

### Viewing a Worksheet Protected by Password

![Access Protected Worksheet](https://ironsoftware.com/static-assets/excel/how-to/set-password-worksheet/set-password-worksheet-access.gif)

## Disabling Password on a Worksheet

To eliminate the password protection from a specific worksheet, utilize the `UnprotectSheet` method. A simple invocation of `workSheet.UnprotectSheet()` suffices to remove the associated password.

```cs
// Unlock the worksheet without needing the password
workSheet.UnprotectSheet();
```

IronXL facilitates the protection and de-protection of any Excel [workbook](https://ironsoftware.com/csharp/excel/how-to/set-password-workbook/) and **worksheet** using merely a single line of C# code.