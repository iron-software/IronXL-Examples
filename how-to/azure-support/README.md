# Can I Deploy IronXL on Azure with .NET?

Certainly! IronXL is fully compatible with Azure, allowing for the creation and reading of QR codes and barcodes in C# & VB.NET applications using Azure services. It's rigorously validated across various Azure environments, including MVC websites, Azure Functions, and more.

---

## Installation of IronXL

Begin by installing IronXL via NuGet:

[NuGet Package for IronXL.Excel](https://www.nuget.org/packages/IronXL.Excel)

```shell
Install-Package IronXL.Excel
```

---

## Using IronXL with Azure

### 2. Azure Hosting Recommendations

For typical use-cases, the Azure **B1** tier is recommended. Those needing more robust performance should consider higher service tiers.

### 3. Choosing the Right .NET Framework

IronXL functions well on both Core and Framework variants in Azure environments. .NET Standard might offer slightly better performance but demands more memory.

#### Note on Azure's Free Tier

The free and shared Azure plans, including the consumption plan, do not perform well for QR processing. We use and recommend at least the Azure B1 hosting or Premium plans for optimal performance.

### 4. Implementing Docker on Azure

Using Docker containers is an effective way to enhance the performance of your IronXL applications on Azure. For detailed guidance, check out our in-depth tutorial:

[Comprehensive IronXL Azure Docker Tutorial](https://ironsoftware.com/csharp/excel/how-to/docker-support/)

### 5. Azure Functions Integration

IronXL is compatible with Azure Functions V3. Currently, version V4 is under review but not yet endorsed.

#### Azure Function Example Code

This example is validated for Azure Functions v3.3.1.0 onward:

```cs
[FunctionName("barcode")]
public static HttpResponseMessage Run(
    [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
    ILogger log)
{
    log.LogInformation("Processing request with C# HTTP trigger function.");
    IronXL.License.LicenseKey = "Your-License-Key-Here";
    var barCode = BarcodeWriter.CreateBarcode("IronXL Test", BarcodeEncoding.QRCode);
    var result = new HttpResponseMessage(HttpStatusCode.OK);
    result.Content = new ByteArrayContent(barCode.ToJpegBinaryData());
    result.Content.Headers.ContentDisposition =
            new ContentDispositionHeaderValue("attachment") { FileName = $"{DateTime.Now.ToString("yyyyMMddmm")}.jpg" };
    result.Content.Headers.ContentType =
            new MediaTypeHeaderValue("image/jpeg");
    return result;
}
```
This example demonstrates setting up an HTTP-triggered Azure function for generating and returning a QR code as a JPEG image.