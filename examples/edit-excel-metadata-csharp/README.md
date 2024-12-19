***Based on <https://ironsoftware.com/examples/edit-excel-metadata-csharp/>***

The preceding code sample illustrates how to manipulate the metadata of an Excel file using the IronXL C# library. To apply a custom author name to a spreadsheet, simply set the `name` to the `Author` property. For instance, you would use `workBook.Metadata.Author = "Your Name"`. You can access and manipulate a variety of metadata attributes via the `Metadata` property of `WorkBook`.

The list below highlights the metadata properties that can be adjusted:

- Modify and access detailed metadata such as:
  - `Author`
  - `Comments`
  - `LastPrinted`
  - `Keywords` and `Category`
  - `Created` and `ModifiedDate`
  - `Subject` and `Title`
  
- Access detailed metadata including:
  - `ApplicationName`
  - `CustomProperties`
  - `Company`
  - `Manager`
  - `Template`

For further details and examples, please visit the ["Edit Workbook Metadata" How-To](https://ironsoftware.com/csharp/excel/how-to/edit-workbook-metadata/) article.