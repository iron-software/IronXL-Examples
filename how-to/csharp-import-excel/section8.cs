using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section8
    {
        public static void Run()
        {
            //to find the sum of specific cell range 
            WorkSheet ["Starting Cell Address : Ending Cell Address"].Sum();
        }
    }
}