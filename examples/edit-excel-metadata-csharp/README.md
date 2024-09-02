The above code example demonstrates how to set, modify, and retrieve metadata for an Excel file using the IronXL C# library. To assign a custom author name to a spreadsheet, simply set the `name` to the `Author` property, like so: `workBook.Metadata.Author = "Your Name"`. You can access and manipulate various metadata properties through the `Metadata` property of `WorkBook`.

Available properties include:

- Modify and retrieve detailed metadata such as:
  - `Author`
  - `Comments`
  - `LastPrinted`
  - `Keywords` and `Category`
  - `Created` and `ModifiedDate`
  - `Subject` and `Title`
- Retrieve detailed metadata such as:
  - `ApplicationName`
  - `CustomProperties`
  - `Company`
  - `Manager`
  - `Template`

For additional information and detailed examples, please visit the ["Edit Workbook Metadata" How-To](https://ironsoftware.com/csharp/excel/how-to/edit-workbook-metadata/) article.