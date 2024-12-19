using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section12
    {
        public static void Run()
        {
            /**
            Data API to Spreadsheet
            anchor-download-data-from-an-api-to-spreadsheet
            **/
            var client = new Client(new Uri("https://restcountries.eu/rest/v2/"));
            List<RestCountry> countries = await client.GetAsync<List<RestCountry>>();
        }
    }
}