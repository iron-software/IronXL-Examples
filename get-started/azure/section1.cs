using System.Net.Http.Headers;
using IronXL;
namespace IronXL.Examples.GettingStarted.Azure
{
    public static class Section1
    {
        public static void Run()
        {
            // This is an Azure Function that processes an HTTP request and returns an Excel file
            [FunctionName("excel")]
            public static HttpResponseMessage Run(
                [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
                ILogger log)
            {
                // Log the processing of the request
                log.LogInformation("C# HTTP trigger function processed a request.");
            
                // Set the IronXL license key
                IronXL.License.LicenseKey = "Key";
            
                // Load an existing workbook
                var workBook = WorkBook.Load("test-wb.xlsx");
            
                // Create a response with the workbook content as an attachment
                var result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new ByteArrayContent(workBook.ToByteArray());
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                { 
                    FileName = $"{DateTime.Now:yyyyMMddmm}.xlsx" 
                };
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                
                // Return the response
                return result;
            }
        }
    }
}