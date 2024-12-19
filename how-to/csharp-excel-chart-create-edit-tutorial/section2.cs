using IronXL.Drawing.Charts;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpExcelChartCreateEditTutorial
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("chart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set the chart type and position
            IChart chart = workSheet.CreateChart(ChartType.Column, 5, 5, 20, 10);
            
            string xAxis = "A2:A7";
            
            // Add the series
            IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
            series.Title = workSheet["B1"].StringValue;
            
            // Add the series
            series = chart.AddSeries(xAxis, "C2:C7");
            series.Title = workSheet["C1"].StringValue;
            
            // Add the series
            series = chart.AddSeries(xAxis, "D2:D7");
            series.Title = workSheet["D1"].StringValue;
            
            // Set the chart title
            chart.SetTitle("Line Chart");
            
            // Set the legend position
            chart.SetLegendPosition(LegendPosition.Bottom);
            
            // Plot the chart
            chart.Plot();
            
            workBook.SaveAs("lineChart.xlsx");
        }
    }
}