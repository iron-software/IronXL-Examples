using System.Linq;
using IronXL.Excel;
namespace ironxl.AddComment
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addComment.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            Cell cellA1 = workSheet["A1"].First();
            
            // Retrieve comment
            var comment = cellA1.Comment;
            
            // Edit comment
            comment.Author = "Jane Doe";
            comment.Content = "Bye World";
            comment.IsVisible = true;
            
            workBook.SaveAs("editComment.xlsx");
        }
    }
}