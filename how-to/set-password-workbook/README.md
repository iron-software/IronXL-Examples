# Setting a Password on a Workbook

It's essential to ensure that sensitive information or data is shared only with authorized individuals. By using IronXL, you have the ability to secure entire spreadsheets or individual [worksheets](https://ironsoftware.com/csharp/excel/how-to/set-password-worksheet/) with password protection.

***

***

## Opening a Password-Protected Workbook

To open a spreadsheet that is safeguarded with a password, you should supply the password as an argument in the `Load` function. For instance, `WorkBook.Load("sample.xlsx", "IronSoftware")`.

Opening a protected spreadsheet is not feasible without the correct password.

## Securing a Workbook with a Password

To implement password protection on a spreadsheet, the `Encrypt` method is used.

```cs
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Load protected spreadsheet file
WorkBook protectedWorkBook = WorkBook.Load("sample.xlsx", "IronSoftware");

// Apply password protection to the spreadsheet file
workBook.Encrypt("IronSoftware");

// Save changes to the file
workBook.Save();
```

### Accessing a Workbook with Password Protection

![Open Protected Spreadsheet](https://ironsoftware.com/static-assets/excel/how-to/set-password-workbook/set-password-workbook-access.gif)

## Removing a Password from a Workbook

To eliminate the password from a spreadsheet, assign `null` to the **Password** field. This operation can only be executed if the original password is known.

```cs
// Deactivate password protection for an opened workbook. Original password is required.
workBook.Password = null;
```

IronXL facilitates the straightforward protection and unprotection of Excel **workbooks** and [worksheets](https://ironsoftware.com/csharp/excel/how-to/set-password-worksheet/) with just a single line of C# code.