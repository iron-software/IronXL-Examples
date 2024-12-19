using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section6
    {
        public static void Run()
        {
            workSheet["From Cell Address:To Cell Address"].Value = "New Value";
        }
    }
}