using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section6
    {
        public static void Run()
        {
            /**
            Insert Data in Range
            anchor-insert-data-in-range
            **/
            WorkSheet ["From Cell Address : To Cell Address"].Value="value";
        }
    }
}