using IronXL;
namespace IronXL.Examples.HowTo.AzureSupport
{
    public static class Section1
    {
        public static void Run()
        {
            // This snippet is a member of a larger component from the accompanying README, not a standalone program.
            // Kept verbatim; see README.md for the full context.
            // [FunctionName("barcode")]
            // public static HttpResponseMessage Run(
            // [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            // ILogger log)
            // {
            // log.LogInformation("C# HTTP trigger function processed a request.");
            // IronXL.License.LicenseKey = "Key";
            // var MyBarCode = BarcodeWriter.CreateBarcode("IronXL Test", BarcodeEncoding.QRCode);
            // var result = new HttpResponseMessage(HttpStatusCode.OK);
            // result.Content = new ByteArrayContent(MyBarCode.ToJpegBinaryData());
            // result.Content.Headers.ContentDisposition =
            // new ContentDispositionHeaderValue("attachment") { FileName = $"{DateTime.Now.ToString("yyyyMMddmm")}.jpg" };
            // result.Content.Headers.ContentType =
            // new MediaTypeHeaderValue("image/jpeg");
            // return result;
            // }
        }
    }
}