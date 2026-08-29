# Can IronXL be Operated with .NET on Azure?

> Full guide: [Can IronXL be Operated with .NET on Azure?](https://ironsoftware.com/csharp/excel/get-started/azure/?utm_source=github)


IronXL reads and writes Excel spreadsheets from C# and VB.NET applications hosted in Azure. It has been deployed and tested across Azure environments including MVC websites and Azure Functions.

---

<p class="main-content__segment-title">Step 1</p>

## 1. Setting Up IronXL

To begin, add IronXL to your project via NuGet: [NuGet IronXL.Excel Package](https://www.nuget.org/packages/IronXL.Excel)

```shell
Install-Package IronXL.Excel
```

---

<p class="main-content__segment-title">How to Tutorial</p>

## 2. Selecting the Right Azure Tier

For most scenarios, we suggest starting with Azure **B1** hosting tiers, unless you are designing a system with very high throughput demands.

## 3. Deciding on the .NET Framework

While IronXL works robustly on both .NET Core and .NET Framework on Azure, applications built on .NET Standard may experience slightly better performance in terms of speed and stability albeit potentially consuming more memory.

### Considerations for Azure's Lower Tiers

It's important to note that the Azure free and shared tiers, including the consumption plan, do not perform well for QR processing. We generally recommend opting for at least an Azure B1 hosting or a Premium plan, which is also what we utilize.

## 4. Utilizing Docker for Enhanced Performance on Azure

Docker offers enhanced control over performance for IronXL applications on Azure. We provide detailed guidance in our [IronXL Azure Docker Guide](https://ironsoftware.com/csharp/excel/how-to/docker-support/?utm_source=github) tailored for both Linux and Windows setups.

## 5. Azure Function Compatibility

IronXL is compatible with Azure Function (specifically Azure Functions V3). While we haven't tested integration with Azure Functions V4 yet, it is on our roadmap.

### Example Code for Azure Function

Here's a working example for Azure Functions v3.3.1.0 and higher:

```csharp
using System.Net;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using IronXL;
using System.Net.Http.Headers;

// Process an HTTP request and return an Excel file via an Azure Function
[FunctionName("excel")]
public static HttpResponseMessage Run(
    [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
    ILogger log)
{
    // Logging the request processing
    log.LogInformation("Processing request via C# HTTP trigger function.");

    // Applying the IronXL license
    IronXL.License.LicenseKey = "Your-Licensce-Key-Here";

    // Loading an Excel workbook
    var workbook = WorkBook.Load("path-to-your-file.xlsx");

    // Constructing the response to include the workbook as an attachment
    var response = new HttpResponseMessage(HttpStatusCode.OK);
    response.Content = new ByteArrayContent(workbook.ToByteArray());
    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
    {
        FileName = $"exported-{DateTime.Now:yyyyMMddHHmm}.xlsx"
    };
    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    
    // Sending the prepared Excel file
    return response;
}
```
This approach ensures IronXL's capabilities are effectively utilized within Azure's cloud environment.