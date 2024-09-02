WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
bool max2 = workSheet["A1:A4"].Max(c => c.IsFormula);
Console.WriteLine(max2);