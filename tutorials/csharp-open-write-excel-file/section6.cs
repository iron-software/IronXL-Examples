using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section6
    {
        public static void Run()
        {
            DataSet xmldataset = new DataSet();
            xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
            WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
            WorkSheet workSheet = workBook.WorkSheets.First();
        }
    }
}