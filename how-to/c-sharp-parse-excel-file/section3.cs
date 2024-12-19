using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section3
    {
        public static void Run()
        {
            //Access the Data by Cell Addressing
            string val = ws ["Cell Address"].ToString();
        }
    }
}