using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AddComment
{
    public static class Section3
    {
        public static void Run()
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