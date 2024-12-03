using IronXL.Excel;
namespace ironxl.WriteExcelNet
{
    public class Section12
    {
        public void Run()
        {
            workSheet["From Cell Address : To Cell Address"].Replace("old value", "new value");
        }
    }
}