# How to Create and Edit Excel Charts in C#

Excel charts offer a graphical depiction of data, making it easier to understand and interpret. Excel supports a range of charts like bar, line, pie charts, etc., each designed for different data interpretations.

IronXL enhances chart creation in C# by supporting various chart types including column, scatter, line, pie, bar, and area charts. It allows for customization of series names, legend positions, chart titles, and the positioning of the charts on the worksheet.

## Create Charts Example

IronXL simplifies chart creation in C#. To build a chart, follow these steps:

1. Use the `CreateChart` method to define the chart type and its location within the worksheet.
2. Employ the `AddSeries` method to introduce data series to the chart. This method accepts a column of data corresponding to minimal chart types. The first parameter denotes the horizontal axis values, while the second one represents the vertical axis values.
3. Optionally, you can set the series name, chart's title, and the position of the legend.
4. Use the `Plot` method to draw the chart on the worksheet. Repeated calls to this method can create multiple charts.

Check out the example charts created from the data in the [chart.xlsx](https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/chart.xlsx) file. Below is how the data looks like:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/data.webp" alt="Data" class="img-responsive add-shadow">
    </div>
</div>

### Column Chart

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workBook = WorkBook.Load("chart.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Configure the chart type and position
IChart chart = workSheet.CreateChart(ChartType.Column, 5, 5, 20, 10);

string xAxis = "A2:A7";

// Series definitions
IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
series.Title = workSheet["B1"].StringValue;

series = chart.AddSeries(xAxis, "C2:C7");
series.Title = workSheet["C1"].StringValue;

series = chart.AddSeries(xAxis, "D2:D7");
series.Title = workSheet["D1"].StringValue;

// Chart title and legend
chart.SetTitle("Column Chart");
chart.SetLegendPosition(LegendPosition.Bottom);

// Draw the chart
chart.Plot();

workBook.SaveAs("columnChart.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/column-chart.webp" alt="Column chart" class="img-responsive add-shadow">
    </div>
</div>

### Line Chart

Converting between a line and column chart is trivial, only requiring a change in the chart type.

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workBook = WorkBook.Load("chart.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set chart specifics
IChart chart = workSheet.CreateChart(ChartType.Line, 5, 5, 20, 10);

// Set axes for the chart
string xAxis = "A2:A7";

// Adding series to chart
IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
series.Title = workSheet["B1"].StringValue;

series = chart.AddSeries(xAxis, "C2:C7");
series.Title = workSheet["C1"].StringValue;

series = chart.AddSeries(xAxis, "D2:D7");
series.Title = workSheet["D1"].StringValue;

// Configuring title and legend position
chart.SetTitle("Line Chart");
chart.SetLegendPosition(LegendPosition.Bottom);

// Render the chart
chart.Plot();

workBook.SaveAs("lineChart.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/line-chart.webp" alt="Line chart" class="img-responsive add-shadow">
    </div>
</div>

### Pie Chart

Pie charts need only one series of data.

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workBook = WorkBook.Load("chart.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Prepare the pie chart
IChart chart = workSheet.CreateChart(ChartType.Pie, 5, 5, 20, 10);

string xAxis = "A2:A7";

// Adding a single series
...