using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section11
    {
        public static void Run()
        {
            //to find the Max in specific cell range 
            WorkSheet ["Starting Cell Address : Ending Cell Address"].Max()
        }
    }
}