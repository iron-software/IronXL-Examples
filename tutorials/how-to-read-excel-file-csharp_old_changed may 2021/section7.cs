using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section7
    {
        public static void Run()
        {
            /**
            Edit Cell Values in Range
            anchor-edit-cell-values-within-a-range
            **/
            //Iterate through the rows
            for (var y = 2; y <= 101; y++)
            {
                var result = new PersonValidationResult { Row = y };
                results.Add(result);
            
                //Get all cells for the person
                var cells = worksheet [$"A{y}:E{y}"].ToList();
            
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