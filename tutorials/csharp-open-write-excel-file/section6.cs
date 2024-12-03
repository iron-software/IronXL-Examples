using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section6
    {
        public void Run()
        {
            DataSet xmldataset = new DataSet();
            xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
            WorkBook workBook = IronXL.WorkBook.Load(xmldataset);
            WorkSheet workSheet = workBook.WorkSheets.First();
        }
    }
}