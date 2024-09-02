using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Overwriting freeze or split pane to column(A-E) and row(1-5) as well as applying prescroll
// The column will show E,G,... and the row will show 5,8,...
workSheet.CreateFreezePane(5, 5, 6, 7);

workBook.SaveAs("createFreezePanes.xlsx");