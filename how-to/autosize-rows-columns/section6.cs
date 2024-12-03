using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section6
    {
        public void Run()
        {
            workSheet.Merge("A1:B1");
            
            workSheet.AutoSizeColumn(0, false);
            workSheet.AutoSizeColumn(1, false);
        }
    }
}