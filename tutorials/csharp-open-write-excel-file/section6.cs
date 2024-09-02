DataSet xmldataset = new DataSet();
xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
WorkSheet workSheet = workBook.WorkSheets.First();