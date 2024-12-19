using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section5
    {
        public static void Run()
        {
            /**
            Insert Data in Cell Address
            anchor-insert-cell-data
            **/
            WorkSheet ["CellAddress"].Value = "Value";
        }
    }
}