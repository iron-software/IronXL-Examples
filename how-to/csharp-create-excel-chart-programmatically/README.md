# Generating Excel Charts with C# and IronXL

***Based on <https://ironsoftware.com/how-to/csharp-create-excel-chart-programmatically/>***


This tutorial guides you through the process of programmatically creating Excel charts using IronXL in C#.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Automation of Excel Charts in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-2-create-excel-chart-for-net">Programmatically generating Excel charts</a></li>
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

<h2>Generating Excel Charts in C#</h2>

1. Install the IronXL library for Excel chart generation.
2. Load an existing Excel document into a `Workbook` object.
3. Initialize a chart using the `CreateChart` method.
4. Define the chart's title and legend.
5. Execute the `Plot` method.
6. Save the updated `Workbook` back to an Excel file.

<p class="main-content__segment-title">Step 1</p>

## 1. Install IronXL

Beginning with the installation, IronXL can be conveniently added using NuGet in Visual Studio:

* Navigate to the Project menu
* Select Manage NuGet Packages
* Look for IronXL.Excel
* Click Install

Alternatively, install via the Developer Command Prompt:

```shell
Install-Package IronXL.Excel
```

Or download directly from here: <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">IronXL Package</a>

<hr class="separator">

<p class="main-content__segment-title">How to Tutorial</p>

## 2. Craft an Excel Chart for .NET

Let’s start the practical component of this tutorial!

Prepare your data within an Excel sheet as illustrated below:

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
    <p><strong>Figure 1</strong> – <em> Data to be used for charting</em></p>
	</div>
</div>

Incorporate the necessary namespaces to manage Excel charts inside IronXL:

```cs
using IronXL;
using IronXL.Drawing.Charts;
```

Implement this code to create the Excel chart programmatically using IronXL:

```cs
private void button1_Click(object sender, EventArgs e)
{
    WorkBook wb = WorkBook.Load("Chart_Ex.xlsx");
    WorkSheet ws = wb.DefaultWorkSheet;
            
    var chart = ws.CreateChart(ChartType.Column, 10, 15, 25, 20);

    const string xAxis = "A2:A7";

    var series = chart.AddSeries(xAxis, "B2:B7");
    series.Title = ws["B1"].StringValue;
            
    series = chart.AddSeries(xAxis, "C2:C7");
    series.Title = ws["C1"].StringValue;
            
    series = chart.AddSeries(xAxis, "D2:D7");
    series.Title = ws["D1"].StringValue;
            
    chart.SetTitle("Sales Performance Chart");
    chart.SetLegendPosition(LegendPosition.Bottom);
    chart.Plot();
    wb.SaveAs("Generated_Column_Chart.xlsx");
}
```

In this method, we start by opening a Workbook and accessing its default Worksheet. A chart is initiated indicating the type and position. Titles and legend positions are set for each added series, and then chart plotting takes place.

<div class="content-img-align-center">
	<div class="center-image-wrapper">
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/chart-output.png"
        alt="Chart output"
      >
   	<p><strong>Figure 2</strong> – <em>Generated chart output</em></p>
	</div>
</div>

<hr class="separator">

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore IronXL API Reference Documentation</h3>
      <p>Delve into the IronXL API Reference Documentation for in-depth knowledge on handling various Excel operations like merging, unmerging, and manipulating cells.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">IronXL API Reference Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>