# How to Modify Excel Workbook Metadata

***Based on <https://ironsoftware.com/how-to/edit-workbook-metadata/>***


Metadata in an Excel file encompasses details such as the title, author, subject, keywords, and dates of creation and modification, among others. This metadata is crucial as it provides context, aids in organizing, and improves the management of spreadsheet files, which is especially beneficial when dealing with a large number of files.


The **IronXL** library simplifies the task of editing workbook metadata, eliminating the need for Office Interop.

***

***

## Example: Modifying Workbook Metadata

To change the author of an Excel file, you can modify the `Author` property. For instance, you might set `workBook.Metadata.Author = "New Author Name";`. The `Metadata` property of the `WorkBook` class allows for easy access and modification of such metadata.

```cs
using System;
using IronXL.Excel;
namespace ironxl.ModifyWorkbookMetadata
{
    public class ModifyMetadataSection
    {
        public void Execute()
        {
            WorkBook workBook = WorkBook.Load("example.xlsx");
            
            // Update author
            workBook.Metadata.Author = "New Author Name";
            // Update comments
            workBook.Metadata.Comments = "Annual report generation";
            // Update title
            workBook.Metadata.Title = "Annual Report 2023";
            // Update keywords
            workBook.Metadata.Keywords = "Annual,Finance,Summary";
            
            // Retrieve the creation date of the workbook
            DateTime? creationDate = workBook.Metadata.Created;
            
            // Retrieve the last printed date
            DateTime? printDate = workBook.Metadata.LastPrinted;
            
            workBook.SaveAs("updatedMetadata.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/edit-workbook-metadata/edit-workbook-metadata.png" alt="Metadata" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Editable and Accessible Metadata Properties

While you can edit several metadata fields, some are only accessible for viewing. Below is a detailed list of the properties:

#### Editable or Modifiable and Accessible
  - Author
  - Comments
  - LastPrinted
  - Keywords and Category
  - Created and ModifiedDate
  - Subject and Title

#### Accessible Only
  - ApplicationName
  - CustomProperties
  - Company
  - Manager
  - Template