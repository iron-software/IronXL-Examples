using IronXL;
using IronXL.Drawing.Charts;

WorkBook workBook = WorkBook.Load("chart.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set the chart type and position
IChart chart = workSheet.CreateChart(ChartType.Pie, 5, 5, 20, 10);

string xAxis = "A2:A7";

// Add the series
IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
series.Title = workSheet["B1"].StringValue;

// Set the chart title
chart.SetTitle("Pie Chart");

// Set the legend position
chart.SetLegendPosition(LegendPosition.Bottom);

// Plot the chart
chart.Plot();

workBook.SaveAs("pieChart.xlsx");