# How to Secure a Workbook with a Password

***Based on <https://ironsoftware.com/how-to/set-password-workbook/>***


It is vital to ensure that sensitive data in spreadsheets is accessible only by the intended recipients. IronXL facilitates this by allowing you to secure entire workbooks and specific <a href="https://ironsoftware.com/csharp/excel/how-to/set-password-worksheet/">worksheets</a> with password protection.

***

***

## Opening a Password-Protected Workbook

To access a workbook that is secured with a password, you need to provide the password as an argument in the `Load` method. For instance: `WorkBook.Load("sample.xlsx", "IronSoftware")`.

Accessing a secure workbook without the correct password is not feasible.

## Securing a Workbook with a Password

To add a password protection to a workbook, employ the `Encrypt` method as shown below:

```cs
using IronXL.Excel;
namespace ironxl.SetPasswordWorkbook
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");

            // Accessing the secure workbook
            WorkBook protectedWorkBook = WorkBook.Load("sample.xlsx", "IronSoftware");
            
            // Applying password protection
            workBook.Encrypt("IronSoftware");
            
            // Save the changes to the workbook
            workBook.Save();
        }
    }
}
```

### Accessing a Workbook with Password Protection

<img src="https://ironsoftware.com/static-assets/excel/how-to/set-password-workbook/set-password-workbook-access.gif" alt="Accessing a Protected Workbook" class="img-responsive add-shadow" style="margin-bottom: 30px;"/>

## Removing a Password from a Workbook

You can easily remove a password from a workbook by setting the **Password** property to null. This requires knowing the original password and can only be done after accessing the workbook:

```cs
using IronXL.Excel;
namespace ironxl.SetPasswordWorkbook
{
    public class Section2
    {
        public void Run()
        {
            // Removes password protection. The original password must be known.
            workBook.Password = null;
        }
    }
}
```

IronXL provides a streamlined approach to both safeguard and remove protections from Excel **workbooks** and <a href="https://ironsoftware.com/csharp/excel/how-to/set-password-worksheet/">worksheets</a> using concise C# code.