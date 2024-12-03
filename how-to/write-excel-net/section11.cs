using IronXL.Excel;
namespace ironxl.WriteExcelNet
{
    public class Section11
    {
        public void Run()
        {
            workSheet.Columns[ColumnIndex].Replace("old value", "new Value");
        }
    }
}