# How to Insert Comments in Excel Cells

***Based on <https://ironsoftware.com/how-to/add-comment/>***


In Excel, comments are notes or annotations that can be attached to cells to provide extra insights without interfering with the cell's actual data. These comments can serve as explanations, context, or reminders regarding the data or formulas in a particular cell.

## Example of Adding a Comment

To add a comment, simply select the desired cell and utilize the `AddComment` method. Comments will not be visible until you hover over the cell.

```cs
using System.Linq;
using IronXL;
namespace ironxl.AddComment
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Create();
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            
            Cell cellA1 = worksheet["A1"].First();
            Cell cellD1 = worksheet["D1"].First();
            
            // Adding comments to cells
            cellA1.AddComment("Hello World!", "John Doe");  // Inserts a hidden comment with content and author.
            cellD1.AddComment(null, null, true);  // Inserts an empty, visible comment without content or author.
            
            workbook.SaveAs("addComment.xlsx");
        }
    }
}
```

***

## Example of Editing a Comment

To modify a comment, access the cell's `Comment` property to fetch the pertinent Comment object. This object allows modification of its Author, Content, and Visibility settings.

```cs
using System.Linq;
using IronXL;
namespace ironxl.AddComment
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load("addComment.xlsx");
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            
            Cell cellA1 = worksheet["A1"].First();
            
            // Accessing the comment
            var comment = cellA1.Comment;
            
            // Modifying the comment details
            comment.Author = "Jane Doe";
            comment.Content = "Bye World";
            comment.IsVisible = true;
            
            workbook.SaveAs("editComment.xlsx");
        }
    }
}
```

***

## Example of Removing a Comment

To delete a comment from a cell, first retrieve the cell's instance, then invoke the `RemoveComment` method.

```cs
using System.Linq;
using IronXL;
namespace ironxl.AddComment
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Load("addComment.xlsx");
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            
            Cell cellA1 = worksheet["A1"].First();
            
            // Deleting the comment
            cellA1.RemoveComment();
            
            workbook.SaveAs("removeComment.xlsx");
        }
    }
}
```