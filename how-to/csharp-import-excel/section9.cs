using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section9
    {
        public static void Run()
        {
            //to find the average of specific cell range 
            WorkSheet ["Starting Cell Address : Ending Cell Address"].Avg()
        }
    }
}