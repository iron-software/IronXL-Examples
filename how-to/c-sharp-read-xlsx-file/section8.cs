using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section8
    {
        public static void Run()
        {
            foreach (var cell in ws ["A2:A10"])
            {
                Console.WriteLine("value is: {0}",  cell.Text);
            }
        }
    }
}