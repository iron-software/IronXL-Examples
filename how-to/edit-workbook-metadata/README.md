# How to Modify Workbook Metadata

Metadata in an Excel workbook entails details such as the title, author, subject, keywords, and dates of creation and modification among other vital information. Such metadata is instrumental in providing context and streamlining the management of Excel files, particularly when handling numerous documents. It greatly enhances the ease of file searching and categorization.

The **IronXL** library enables users to modify workbook metadata effortlessly without the necessity of Office Interop.

***

***

## Example: Modifying Workbook Metadata

To modify the author of an Excel file, simply assign a new value to the `Author` property. For instance, `workBook.Metadata.Author = "Jane Doe";` assigns a new author. The `Metadata` class of the `WorkBook` offers other modifiable properties.

```cs
using IronXL;
using System;

WorkBook workBook = WorkBook.Load("exampleFile.xlsx");

// Updating author information
workBook.Metadata.Author = "Jane Doe";
// Adding comments
workBook.Metadata.Comments = "Quarterly Financials";
// Updating the title
workBook.Metadata.Title = "Q1 Results";
// Adding relevant keywords
workBook.Metadata.Keywords = "Finance,Quarterly";

// Fetch creation date of the file
DateTime? creationDate = workBook.Metadata.Created;

// Fetch the last printed date of the file
DateTime? printDate = workBook.Metadata.LastPrinted;

workBook.SaveAs("updatedMetadata.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/edit-workbook-metadata/edit-workbook-metadata.png" alt="Metadata Editing" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Metadata Fields: Editable and Non-Editable

Certain metadata fields can be both modified and retrieved, while some are only accessible and cannot be changed. Here’s what you can do:

#### Editable and Retrievable
  - Author
  - Comments
  - LastPrinted
  - Keywords and Category
  - Created and ModifiedDate
  - Subject and Title

#### Read-Only Fields
  - ApplicationName
  - CustomProperties
  - Company
  - Manager
  - Template

This flexibility provided by IronXL makes it a powerful tool for managing workbook metadata effectively.