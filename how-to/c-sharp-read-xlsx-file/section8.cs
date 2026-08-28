using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section8
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            foreach (var cell in ws ["A2:A10"])
            {
                Console.WriteLine("value is: {0}",  cell.Text);
            }
        }
    }
}