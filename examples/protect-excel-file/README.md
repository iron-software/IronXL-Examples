***Based on <https://ironsoftware.com/examples/protect-excel-file/>***

Ensuring the right person receives the correct data is crucial in affirming proper authorization. IronXL facilitates this through capabilities that allow the creation of password-protected spreadsheets, including individual `WorkSheet`s.

## Opening Password-Protected Files

To open a protected spreadsheet, use the `Load` method and supply the password as its second argument. For instance:

```csharp
WorkBook.Load("sample.xlsx", "IronSoftware");
```

## Securing Spreadsheets

To apply password protection to a spreadsheet, use the `Encrypt` method like so:

```csharp
workBook.Encrypt("IronSoftware");
```

## Removing Password Protection

To remove a password from a spreadsheet, reset the `Password` field to `null` as shown here:

```csharp
workBook.Password = null;
```
This operation should be performed after accessing the workbook, which requires knowing the original password.

With IronXL, you can easily protect and unprotect Excel `WorkBook` and `WorkSheet` objects using straightforward C# commands.