# How to Add a Comment to Excel Cells

***Based on <https://ironsoftware.com/how-to/add-comment/>***


In Excel, comments serve as helpful annotations or notes attached to a cell. They provide supplementary information, adding context or reminders without altering the content of the cell itself. This additional layer of information is especially useful for explaining the data or calculations in a cell.

<h3>Getting Started with IronXL</h3>

----------------------------------

## Example of Adding a Comment

To embed a comment into a cell, select the cell and then utilize the `AddComment` method. By default, the comment will not be visible. To view the comment, simply hover over the cell.

```cs
using IronXL;
using System.Linq;

WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

Cell cellA1 = worksheet["A1"].First();
Cell cellD1 = worksheet["D1"].First();

// Adding comments to cells
cellA1.AddComment("Hello World!", "John Doe"); // Embed a comment with specified content and author. Invisible by default.
cellD1.AddComment(null, null, true); // Embed an empty comment that is automatically visible.

workbook.SaveAs("addComment.xlsx");
```

* * *

## Example of Editing a Comment

To modify a comment, access the cell's **Comment** property to retrieve and manipulate the Comment object of that cell. This allows you to alter properties such as the Author, Content, and visibility of the comment.

```cs
using IronXL;
using System.Linq;

WorkBook workbook = WorkBook.Load("addComment.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

Cell cellA1 = worksheet["A1"].First();

// Access the existing comment
var comment = cellA1.Comment;

// Modify the comment details
comment.Author = "Jane Doe";
comment.Content = "Bye World";
comment.IsVisible = true;

workbook.SaveAs("editComment.xlsx");
```

* * *

## Example of Removing a Comment

To remove a comment from a cell, first, ensure you can access the relevant cell. Once accessed, execute the `RemoveComment` method on the specific cell.

```cs
using IronXL;
using System.Linq;

WorkBook workbook = WorkBook.Load("addComment.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

Cell cellA1 = worksheet["A1"].First();

// Deleting the comment
cellA1.RemoveComment();

workbook.SaveAs("removeComment.xlsx");
```