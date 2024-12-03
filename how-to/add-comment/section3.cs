using System.Linq;
using IronXL.Excel;
namespace ironxl.AddComment
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addComment.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            Cell cellA1 = workSheet["A1"].First();
            
            // Remove comment
            cellA1.RemoveComment();
            
            workBook.SaveAs("removeComment.xlsx");
        }
    }
}