using IronXL.Excel;
namespace ironxl.WriteExcelNet
{
    public class Section10
    {
        public void Run()
        {
            workSheet.Rows[RowIndex].Replace("old value", "new value");
        }
    }
}