using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section5
    {
        public static void Run()
        {
            Random r = new Random();
            for (int i = 2 ; i <= 11 ; i++)
            {
                workSheet["A" + i].Value = r.Next(1, 1000);
                workSheet["B" + i].Value = r.Next(1000, 2000);
                workSheet["C" + i].Value = r.Next(2000, 3000);
                workSheet["D" + i].Value = r.Next(3000, 4000);
                workSheet["E" + i].Value = r.Next(4000, 5000);
                workSheet["F" + i].Value = r.Next(5000, 6000);
                workSheet["G" + i].Value = r.Next(6000, 7000);
                workSheet["H" + i].Value = r.Next(7000, 8000);
                workSheet["I" + i].Value = r.Next(8000, 9000);
                workSheet["J" + i].Value = r.Next(9000, 10000);
                workSheet["K" + i].Value = r.Next(10000, 11000);
                workSheet["L" + i].Value = r.Next(11000, 12000);
            }
        }
    }
}