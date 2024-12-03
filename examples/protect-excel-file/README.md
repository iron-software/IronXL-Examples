***Based on <https://ironsoftware.com/examples/protect-excel-file/>***

Ensuring the right individuals have access to specific information or data is pivotal in maintaining proper authorization. IronXL enables users to create spreadsheets with password protection and also secure individual `WorkSheet` objects with passwords.

## Accessing

To open a password-protected spreadsheet, pass the password as the second argument of the `Load` method. For instance, `WorkBook.Load("sample.xlsx", "IronSoftware")` demonstrates how to access a locked file.

## Applying

To safeguard a spreadsheet with a password, utilize the `Encrypt` method. A typical usage is: `workBook.Encrypt("IronSoftware")`.

## Removing

To remove a password from a spreadsheet, simply set the `Password` property to `null`, as shown here: `workBook.Password = null`. This operation should be performed only after you've successfully opened the workbook, meaning you must know the original password.

IronXL provides a straightforward mechanism to protect and unprotect both Excel `WorkBook` and `WorkSheet` objects with a mere line of C# code.