using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section12
    {
        public static void Run()
        {
            //import WorkBook into dataset
            DataSet ds = WorkBook.ToDataSet();
        }
    }
}