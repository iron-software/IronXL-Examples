using System.Linq;
using IronXL.Excel;
namespace ironxl.AddComment
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            Cell cellA1 = workSheet["A1"].First();
            Cell cellD1 = workSheet["D1"].First();
            
            // Add comments
            cellA1.AddComment("Hello World!", "John Doe"); // Add comment with content and author. The comment is invisible by default.
            cellD1.AddComment(null, null, true); // Add comment with no content and no author. The comment is set to be visible.
            
            workBook.SaveAs("addComment.xlsx");
        }
    }
}