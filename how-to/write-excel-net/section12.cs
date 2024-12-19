using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section12
    {
        public static void Run()
        {
            workSheet["From Cell Address : To Cell Address"].Replace("old value", "new value");
        }
    }
}