using System.Linq;
using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            int ItemIndex = 0;
            IronXL.Cell[] array = workSheet["A1:A10"].ToArray();

            string item = array [ItemIndex].Value.ToString();
        }
    }
}