using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("addComment.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

Cell cellA1 = workSheet["A1"].First();

// Remove comment
cellA1.RemoveComment();

workBook.SaveAs("removeComment.xlsx");