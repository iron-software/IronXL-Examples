# Running IronXL with .NET on Azure

***Based on <https://ironsoftware.com/how-to/azure-support/>***


Certainly! IronXL is fully compatible with Azure for generating and reading QR & Barcodes in C# & VB .NET applications. It has been validated across a variety of Azure services including MVC websites and Azure Functions.

<hr class="separator">

<p class="main-content__segment-title">Step 1</p>

## Getting Started with IronXL

Begin by installing IronXL via NuGet: [NuGet Package for IronXL.Excel](https://www.nuget.org/packages/IronXL.Excel)

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<p class="main-content__segment-title">Step-by-Step Guide</p>

## Optimal Azure Hosting Options

We suggest starting with the Azure **B1** level for most library applications. For systems requiring high throughput, consider an upgrade for optimal performance.

## Choosing the Right Framework

IronXL operates efficiently whether you’re using .NET Core, .NET Framework, or .NET Standard, with the latter often delivering slightly better performance and stability but at the cost of higher memory usage.

### Limitations of Azure's Free and Shared Tiers

For QR processing workloads, avoid Azure's free and shared tiers, including the consumption plan. Instead, opt for the Azure B1 or Premium plans, which we use ourselves for optimal performance.

## Utilizing Docker with Azure

Deploying IronXL within Docker containers on Azure can significantly enhance control over performance. We offer a detailed guide for both Linux and Windows setups in our [IronXL Azure Docker Guide](https://ironsoftware.com/csharp/excel/how-to/docker-support/).

## Support for Azure Functions

IronXL is compatible with Azure Functions V3. Testing with V4 is pending but is on our development roadmap.

### Example Code for an Azure Function

Here's a reliable example for Azure Functions v3.3.1.0 or later:
```cs
    [FunctionName("barcode")]
    public static HttpResponseMessage Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
    {
        log.LogInformation("C# HTTP trigger function processed a request.");
        IronXL.License.LicenseKey = "Your-Licence-Key-Here";  // Remember to replace this with your actual license key
        var generatedBarcode = BarcodeWriter.CreateBarcode("IronXL Test", BarcodeEncoding.QRCode);
        var result = new HttpResponseMessage(HttpStatusCode.OK);
        result.Content = new ByteArrayContent(generatedBarcode.ToJpegBinaryData());
        result.Content.Headers.ContentDisposition =
                new ContentDispositionHeaderValue("attachment") { FileName = $"{DateTime.Now.ToString("yyyyMMddmm")}.jpg" };
        result.Content.Headers.ContentType =
                new MediaTypeHeaderValue("image/jpeg");
        return result;
    }
```

This paraphrased content maintains a professional, helpful, and conversational tone suitable for both technical and non-technical audiences, reinforcing the usability and versatility of IronXL within Azure environments.