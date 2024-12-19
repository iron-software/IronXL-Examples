using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section5
    {
        public static void Run()
        {
            /**
            Set Cell Value Dynamically
            anchor-set-cell-values-dynamically
            **/
            Random r = new Random();
            for (int i = 2; i <= 11; i++)
            {
            	sheet ["A" + i].Value = r.Next(1, 1000);
            	sheet ["B" + i].Value = r.Next(1000, 2000);
            	sheet ["C" + i].Value = r.Next(2000, 3000);
            	sheet ["D" + i].Value = r.Next(3000, 4000);
            	sheet ["E" + i].Value = r.Next(4000, 5000);
            	sheet ["F" + i].Value = r.Next(5000, 6000);
            	sheet ["G" + i].Value = r.Next(6000, 7000);
            	sheet ["H" + i].Value = r.Next(7000, 8000);
            	sheet ["I" + i].Value = r.Next(8000, 9000);
            	sheet ["J" + i].Value = r.Next(9000, 10000);
            	sheet ["K" + i].Value = r.Next(10000, 11000);
            	sheet ["L" + i].Value = r.Next(11000, 12000);
            }
        }
    }
}