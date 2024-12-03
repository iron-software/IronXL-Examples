using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Create freeze pane from column(A-B) and row(1-3)
workSheet.CreateFreezePane(2, 3);

// Overwriting freeze or split pane to column(A-E) and row(1-5) as well as applying prescroll
// The column will show E,G,... and the row will show 5,8,...
workSheet.CreateFreezePane(5, 5, 6, 7);

workBook.SaveAs("createFreezePanes.xls");

// Remove all existing freeze or split pane
workSheet.RemovePane();
