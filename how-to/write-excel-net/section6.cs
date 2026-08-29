using IronXL;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet["From Cell Address:To Cell Address"].Value = "New Value";
        }
    }
}