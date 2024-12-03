***Based on <https://ironsoftware.com/examples/edit-excel-metadata-csharp/>***

The above example illustrates how to work with metadata in Excel files using the IronXL C# library. To insert a custom author name into a spreadsheet, you simply set the `name` to the `Author` property like this: `workBook.Metadata.Author = "Your Name"`. You can access and manipulate various metadata attributes through the `Metadata` property of the `WorkBook`.

Below is a list of the metadata properties you can set, modify, and retrieve:

- Customize and manage detailed metadata, including:
  - `Author`
  - `Comments`
  - `LastPrinted`
  - `Keywords` and `Category`
  - `Created` and `ModifiedDate`
  - `Subject` and `Title`
- Access detailed metadata such as:
  - `ApplicationName`
  - `CustomProperties`
  - `Company`
  - `Manager`
  - `Template`

For additional insights and step-by-step guidance, visit the ["Edit Workbook Metadata" How-To](https://ironsoftware.com/csharp/excel/how-to/edit-workbook-metadata/) article.