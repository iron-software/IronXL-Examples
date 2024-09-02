WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
WorkSheet workSheet = workBook.WorkSheets.First();
string cell = workSheet["A1"].StringValue;
Console.WriteLine(cell);