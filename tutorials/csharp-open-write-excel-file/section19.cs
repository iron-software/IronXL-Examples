WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();
decimal max = workSheet["A2:A4"].Max();
Console.WriteLine(max);