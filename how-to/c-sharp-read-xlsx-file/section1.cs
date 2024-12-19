using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section1
    {
        public static void Run()
        {
            /**
            Load Workbook
            anchor-load-workbook
            **/
            WorkBook wb = WorkBook.Load("sample.xlsx");//Excel file path
        }
    }
}