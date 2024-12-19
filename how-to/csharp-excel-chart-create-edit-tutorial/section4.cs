using IronXL.Drawing.Charts;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpExcelChartCreateEditTutorial
{
    public static class Section4
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("pieChart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Retrieve the chart
            IChart chart = workSheet.Charts[0];
            
            // Edit the legend position
            chart.SetLegendPosition(LegendPosition.Top);
            
            // Edit the chart title
            chart.SetTitle("Edited Chart");
            
            workBook.SaveAs("editedChart.xlsx");
        }
    }
}