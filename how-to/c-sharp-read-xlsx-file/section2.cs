using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section2
    {
        public static void Run()
        {
            /**
            Access Sheet by Name
            anchor-access-specific-worksheet
            **/
            WorkSheet ws = wb.GetWorkSheet("Sheet1"); //by sheet name
        }
    }
}