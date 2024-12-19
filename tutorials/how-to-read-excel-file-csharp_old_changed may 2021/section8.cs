using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section8
    {
        public static void Run()
        {
            /**
            Validate Spreadsheet Data
            anchor-validate-spreadsheet-data
            **/
            //Iterate through the rows
            for (var i = 2; i <= 101; i++)
            {
                var result = new PersonValidationResult { Row = i };
                results.Add(result);
            
                //Get all cells for the person
                var cells = worksheet [$"A{i}:E{i}"].ToList();
            
                //Validate the phone number (1 = B)
                var phoneNumber = cells [1].Value;
                result.PhoneNumberErrorMessage = ValidatePhoneNumber(phoneNumberUtil, (string)phoneNumber);
            
                //Validate the email address (3 = D)
                result.EmailErrorMessage = ValidateEmailAddress((string)cells [3].Value);
            
                //Get the raw date in the format of Month Day [suffix], Year (4 = E)
                var rawDate = (string)cells [4].Value;
                result.DateErrorMessage = ValidateDate(rawDate);
            }
        }
    }
}