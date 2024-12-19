# Generating Excel Charts with C# Using IronXL

***Based on <https://ironsoftware.com/how-to/csharp-create-excel-chart-programmatically/>***


This guide provides step-by-step instructions for creating Excel charts in C# utilizing the IronXL library.

<div class="learn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Programmatic Chart Creation in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-2-create-excel-chart-for-net">Using code to generate Excel graphs</a></li>
        <li><a href="#anchor-2-create-excel-chart-for-net">Incorporating series, titles, and legends</a></li>
      </ul>
    </div>
     <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h2>Steps to Creating an Excel Chart in C#</h2>

1. Install the library needed for Excel chart generation.
2. Load your existing Excel document into a `Workbook` object.
3. Utilize `CreateChart` to forge a new chart.
4. Assign a title and set up the legend for your chart.
5. Execute the `Plot` method to draw the chart.
6. Persist the modified `Workbook` back to an Excel file.

<p class="main-content__segment-title">Step 1</p>

## 1. Setting Up IronXL

The quickest way to install IronXL is by using the NuGet Package Manager in Visual Studio:

* Go to the Project menu;
* Choose Manage NuGet Packages;
* Search for IronXL.Excel;
* Click on Install.

Alternatively, you can run the following command in the Developer Command Prompt:

```shell
Install-Package IronXL.Excel
```

Or, download the package directly through this link: <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">IronXL Package</a>

<hr class="separator">

<p class="main-content__segment-title">How to Tutorial</p>

## 2. Creating an Excel Chart for .NET

Begin your project by populating a new Excel Spreadsheet as illustrated below:

<div class="content-img-align-center">
	<div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/data-to-be-used-for-charting.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/data-to-be-used-for-charting.png"
        alt="Data to be used for charting"
      >
    </a>
    <p><strong>Figure 1</strong> – <em>Data to be used for charting</em></p>
	</div>
</div>

Include the necessary namespaces to deal with Excel charts in IronXL:

```cs
using IronXL;
using IronXL.Drawing.Charts;
```

Implement code to dynamically generate an Excel chart using IronXL:

```cs
private void button1_Click(object sender, EventArgs e)
{
    WorkBook wb = WorkBook.Load("Chart_Ex.xlsx");
    WorkSheet ws = wb.DefaultWorkSheet;
            
    var chart = ws.CreateChart(ChartType.Column, 10, 15, 25, 20);

    // Define the X-Axis range.
    const string xAxis = "A2:A7";

    var series = chart.AddSeries(xAxis, "B2:B7");
    series.Title = ws["B1"].StringValue;
            
    series = chart.AddSeries(xAxis, "C2:C7");
    series.Title = ws["C1"].StringValue;
            
    series = chart.AddSeries(xAxis, "D2:D7");
    series.Title = ws["D1"].StringValue;
            
    chart.SetTitle("Column Chart");
    chart.SetLegendPosition(LegendPosition.Bottom);
    chart.Plot();
    wb.SaveAs("Exported_Column_Chart.xlsx");
}
```

In this example, we configure the chart type and location using the `CreateChart` method of the `WorkSheet` object, then add title and legend to the series as demonstrated.

<div class="content-img-align-center">
	<div class="center-image-wrapper">
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/chart-output.png"
        alt="Chart output"
      >
   	<p><strong>Figure 2</strong> – <em>Chart output</em></p>
	</div>
</div>

<hr class="separator">

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Documentation</h3>
      <p>Explore additional capabilities like merging, unmerging, and manipulating cells in Excel sheets through the detailed IronXL API Reference Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Learn More <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>