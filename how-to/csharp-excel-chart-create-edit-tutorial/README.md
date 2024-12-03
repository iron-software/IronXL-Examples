# How to Create and Edit Excel Charts in C#

***Based on <https://ironsoftware.com/how-to/csharp-excel-chart-create-edit-tutorial/>***


Excel charts serve as visual representations of data, significantly enhancing the interpretability and meaningfulness of the information presented. Excel offers an assortment of chart types such as bar, line, pie charts and others, catering to various data analysis needs.

IronXL provides robust support for creating and manipulating various types of charts, including column, scatter, line, pie, bar, and area charts. These charts can be customized in terms of series names, legend positions, chart titles, and other properties.

## Creating Charts with IronXL

IronXL facilitates the creation of several chart types including column, scatter, line, pie, bar, and area charts. The chart creation process involves several straightforward steps:

1. Utilize the `CreateChart` method to define the chart type and its position on the spreadsheet.
2. Use the `AddSeries` method to input data. This method can accept a single data column, which suffices for certain types of charts. The first parameter specifies the x-axis values and the second, the y-axis values.
3. You can optionally set the series name, chart name, and position of the legend.
4. The `Plot` method renders the chart using the specified data. Multiple executions of this method will create multiple charts rather than modifying an existing one.

You can start creating charts using data from the Excel file available here: [Download chart.xlsx](https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/chart.xlsx).

Here's a preview of the data:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/data.webp" alt="Data" class="img-responsive add-shadow">
    </div>
</div>

### Example: Creating a Column Chart

```cs
using IronXL.Drawing.Charts;
using IronXL.Excel;
namespace ironxl.CsharpExcelChartCreateEditTutorial
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("chart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Define the chart type and position
            IChart chart = workSheet.CreateChart(ChartType.Column, 5, 5, 20, 10);
            
            string xAxis = "A2:A7";
            
            // Insert the series
            IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
            series.Title = workSheet["B1"].StringValue;
            
            // Additional series
            series = chart.AddSeries(xAxis, "C2:C7");
            series.Title = workSheet["C1"].StringValue;
            series = chart.AddSeries(xAxis, "D2:D7");
            series.Title = workSheet["D1"].StringValue;
            
            // Configure chart title and legend
            chart.SetTitle("Column Chart");
            chart.SetLegendPosition(LegendPosition.Bottom);
            
            // Render the chart
            chart.Plot();
            
            workBook.SaveAs("columnChart.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/column-chart.webp" alt="Column chart" class="img-responsive add-shadow">
    </div>
</div>

### Example: Creating a Line Chart

Simply switching the chart type can transform a column chart into a line chart, while retaining the other properties.

```cs
using IronXL.Drawing.Charts;
using IronXL.Excel;
namespace ironxl.CsharpExcelChartCreateEditTutorial
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("chart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Define the chart type and position
            IChart chart = workSheet.CreateChart(ChartType.Line, 5, 5, 20, 10);
            
            string xAxis = "A2:A7";
            
            // Insert the series
            IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
            series.Title = workSheet["B1"].StringValue;
            
            // Additional series
            series = chart.AddSeries(xAxis, "C2:C7");
            series.Title = workSheet["C1"].StringValue;
            series = chart.AddSeries(xAxis, "D2:D7");
            series.Title = workSheet["D1"].StringValue;
            
            // Configure chart title and legend
            chart.SetTitle("Line Chart");
            chart.SetLegendPosition(LegendPosition.Bottom);
            
            // Render the chart
            chart.Plot();
            
            workBook.SaveAs("lineChart.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/line-chart.webp" alt="Line chart" class="img-responsive add-shadow">
    </div>
</div>

### Example: Creating a Pie Chart

A pie chart is typically generated using only one data column.

```cs
using IronXL.Drawing.Charts;
using IronXL.Excel;
namespace ironxl.CsharpExcelChartCreateEditTutorial
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("chart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Define the chart type and position
            IChart chart = workSheet.CreateChart(ChartType.Pie, 5, 5, 20, 10);
            
            string xAxis = "A2:A7";
            
            // Insert the series
            IChartSeries series = chart.AddSeries(xAxis, "B2:B7");
            series.Title = workSheet["B1"].StringValue;
            
            // Configure chart title and legend
            chart.SetTitle("Pie Chart");
            chart.SetLegendPosition(LegendPosition.Bottom);
            
            // Render the chart
            chart.Plot();
            
            workBook.SaveAs("pieChart.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/pie-chart.webp" alt="Pie chart" class="img-responsive add-shadow">
    </div>
</div>

## Editing an Existing Chart Example

Modifying existing charts is straightforward in IronXL. You can change properties such as the legend position and the chart title.

```cs
using IronXL.Drawing.Charts;
using IronXL.Excel;
namespace ironxl.CsharpExcelChartCreateEditTutorial
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("pieChart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Access the chart
            IChart chart = workSheet.Charts[0];
            
            // Modify the legend position
            chart.SetLegendPosition(LegendPosition.Top);
            
            // Update the chart title
            chart.SetTitle("Edited Chart");
            
            workBook.SaveAs("editedChart.xlsx");
        }
    }
}
```

Before and after images of the edited chart:

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 48%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/before.webp" alt="Before" class="img-responsive add-shadow">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Before</p>
    </div>
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/after.webp" alt="After" class="img-responsive add-shadow">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">After</p>
    </div>
</div>

## Removing a Chart Example

To delete a chart, retrieve the specific chart object and use the `RemoveChart` method.

```cs
using System.Collections.Generic;
using IronXL.Excel;
namespace ironxl.CsharpExcelChartCreateEditTutorial
{
    public class Section5
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("pieChart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Access the list of charts
            List<IChart> charts = workSheet.Charts;
            
            // Delete the chart
            workSheet.RemoveChart(charts[0]);
            
            workBook.SaveAs("removedChart.xlsx");
        }
    }
}
```