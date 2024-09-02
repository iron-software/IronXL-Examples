using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Create freeze pane from column(A-B) and row(1-3)
workSheet.CreateFreezePane(2, 3);

workBook.SaveAs("createFreezePanes.xlsx");