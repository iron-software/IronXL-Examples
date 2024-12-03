using IronXL.Excel;
namespace ironxl.WriteExcelNet
{
    public class Section6
    {
        public void Run()
        {
            workSheet["From Cell Address:To Cell Address"].Value = "New Value";
        }
    }
}