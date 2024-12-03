using IronXL;
using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section4
    {
        public void Run()
        {
            workSheet.Merge("A1:A3");
            
            workSheet.AutoSizeRow(0, false);
            workSheet.AutoSizeRow(1, false);
            workSheet.AutoSizeRow(2, false);
        }
    }
}