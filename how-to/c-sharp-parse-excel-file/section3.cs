using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            //Access the Data by Cell Addressing
            string val = ws ["Cell Address"].ToString();
        }
    }
}