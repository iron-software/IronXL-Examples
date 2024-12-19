# How to Password Protect an Excel Workbook with IronXL

***Based on <https://ironsoftware.com/how-to/set-password-workbook/>***


It's important to ensure that sensitive data within an Excel file is accessible only to authorized users. Using IronXL, you can secure your data by applying password protection to both the Excel workbook and individual [worksheets](https://ironsoftware.com/csharp/excel/how-to/set-password-worksheet/).

***

***

### Getting Started with IronXL

***

## Opening a Password-Protected Workbook

To open a spreadsheet that has password protection, you must supply the password along with the file name when calling the `Load` function. This looks like: `WorkBook.Load("sample.xlsx", "IronSoftware")`.

Remember, accessing a password-protected workbook is impossible without the correct password.

## Setting a Password on a Workbook

Securing a spreadsheet with a password involves using the `Encrypt` method as shown below:

```cs
// Load an existing workbook
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Access a workbook which is already protected
WorkBook protectedWorkBook = WorkBook.Load("sample.xlsx", "IronSoftware");

// Apply password protection to the workbook
workBook.Encrypt("IronSoftware");

// Finally, save the changes to the workbook
workBook.Save();
```

### Accessing a Password-Protected Workbook

![Open Protected Spreadsheet](https://ironsoftware.com/static-assets/excel/how-to/set-password-workbook/set-password-workbook-access.gif "Effectively opening a secured workbook")

## Removing a Workbook's Password

Deleting a password from an Excel file is straightforward; set the **Password** to `null`. This step must follow successful authentication with the original password:

```cs
// To remove password protection, ensure the original password is known
workBook.Password = null;
```

IronXL simplifies the process of protecting and unprotecting Excel **workBooks** and [worksheets](https://ironsoftware.com/csharp/excel/how-to/set-password-worksheet/) with minimal coding effort, all through C#.