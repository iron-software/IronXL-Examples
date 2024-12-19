using IronXL.Excel;
namespace IronXL.Examples.HowTo.LicenseKeys
{
    public static class Section2
    {
        public static void Run()
        {
            // Check if a given license key string is valid.
            bool result = IronXL.License.IsValidLicense("IronXL-MYLICENSE-KEY-1EF01");
            
            // Check if IronXL is licensed successfully 
            bool is_licensed = IronXL.License.IsLicensed;
        }
    }
}