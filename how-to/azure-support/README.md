# Can I Run IronXL with .NET on Azure?

***Based on <https://ironsoftware.com/how-to/azure-support/>***


Indeed, IronXL is fully compatible for use on Azure, enabling the generation and reading of QR & Barcodes within .NET applications crafted in C# and VB.NET. It also supports the extraction of barcodes and QR codes from scanned documents.

IronXL has undergone extensive testing across various Azure services, including MVC websites, Azure Functions, and beyond.

<hr class="separator">

<p class="main-content__segment-title">Step 1</p>

## 1. Getting Started with IronXL

Begin by installing IronXL via NuGet: [https://www.nuget.org/packages/IronXL.Excel](https://www.nuget.org/packages/IronXL.Excel)

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<p class="main-content__segment-title">How to Tutorial</p>

## 2. Selecting Appropriate Azure Tiers

For typical use, the Azure **B1** hosting tier is recommended. For systems with higher demands in terms of throughput, an upgrade might be necessary.

## 3. Selecting the Right Framework

IronXL functions efficiently on both the Core and Framework editions on Azure, though .NET Standard applications slightly edge out with better speed and stability, albeit at a higher memory usage.

### Considerations for Azure Free Tier

The Azure free, shared, and consumption plans are not ideal for QR processing. Instead, opt for the B1 hosting or Premium plans, which are proven in our own usage.

## 4. Leveraging Docker on Azure

Utilizing Docker containers is an effective strategy to optimize performance for IronXL applications and Functions on Azure.

Follow our detailed [IronXL Azure Docker Guide](https://ironsoftware.com/csharp/excel/how-to/docker-support/) available for both Linux and Windows.

## 5. Azure Function Support by IronXL

IronXL supports Azure Function Version 3. Although not yet tested with Version 4, it is on our development roadmap.

### Example Code for Azure Function

This code has been tested on Azure Functions v3.3.1.0 and above:
```cs
    [FunctionName("barcode")]
    public static HttpResponseMessage Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
    {
        log.LogInformation("C# HTTP trigger function processed a request.");
        IronXL.License.LicenseKey = "Key";
        var MyBarCode = BarcodeWriter.CreateBarcode("IronXL Test", BarcodeEncoding.QRCode);
        var result = new HttpResponseMessage(HttpStatusCode.OK);
        result.Content = new ByteArrayContent(MyBarCode.ToJpegBinaryData());
        result.Content.Headers.ContentDisposition =
                new ContentDispositionHeaderValue("attachment") { FileName = $"{DateTime.Now.ToString("yyyyMMddmm")}.jpg" };
        result.Content.Headers.ContentType =
                new MediaTypeHeaderValue("image/jpeg");
        return result;
    }
```