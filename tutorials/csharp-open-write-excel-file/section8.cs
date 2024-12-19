using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section8
    {
        public static void Run()
        {
            StreamReader jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
            var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
            var xmldataset = countryList.ToDataSet();
            WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
            WorkSheet workSheet = workBook.WorkSheets.First();
        }
    }
}