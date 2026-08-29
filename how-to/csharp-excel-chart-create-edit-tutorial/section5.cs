using IronXL.Drawing.Charts;
using System.Collections.Generic;
using IronXL;
namespace IronXL.Examples.HowTo.CsharpExcelChartCreateEditTutorial
{
    public static class Section5
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("pieChart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Retrieve the chart
            List<IChart> chart = workSheet.Charts;
            
            // Remove the chart
            workSheet.RemoveChart(chart[0]);
            
            workBook.SaveAs("removedChart.xlsx");
        }
    }
}