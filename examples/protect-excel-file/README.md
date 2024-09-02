Delivering information securely to the intended recipient is crucial for maintaining proper authorization. IronXL facilitates the creation of password-protected spreadsheets, including individual `WorkSheet` protection.

## Opening a Protected Spreadsheet

To open a password-protected spreadsheet, pass the password as the second argument to the `Load` method. For instance, use `WorkBook.Load("sample.xlsx", "IronSoftware")` to access the file.

## Securing a Spreadsheet

To apply password protection to a spreadsheet, utilize the `Encrypt` method. For example, secure your spreadsheet by invoking `workBook.Encrypt("IronSoftware")`.

## Removing Password Protection

To remove a password from a spreadsheet, simply reset the `Password` property to `null` as shown: `workBook.Password = null`. This operation can only be performed after successfully accessing the workbook, which requires knowledge of the original password.

IronXL offers an efficient way to both protect and unprotect any Excel `WorkBook` and `WorkSheet` using just a single line of C# code.