# IronXL Integration with AWS Lambda using .NET Core

IronXL offers full support for integrating with AWS Lambda Functions across .NET Standard Libraries, Core applications, as well as .NET 5 and .NET 6 environments.

To include the AWS Toolkit in Visual Studio, refer to the comprehensive guide available at [Using the AWS Lambda Templates in the AWS Toolkit for Visual Studio](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

By installing the AWS Toolkit in Visual Studio, you gain the ability to initiate AWS Lambda Function Projects. Discover the steps to create such projects in Visual Studio by visiting this [link](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

### Example of AWS Lambda Function Code with IronXL

Once you have set up a new AWS Lambda Function project, you can use the following C# code snippet:
```cs
namespace AWSLambdaIronXL
{
    public class Function
    {
        /// <summary>
        /// A straightforward function that takes input string to perform ToUpper transformation.
        /// </summary>
        /// <param name="input">Input string</param>
        /// <param name="context">Lambda context parameter</param>
        /// <returns>The result in a base64 encoded string of the Excel file</returns>
        public string FunctionHandler(string input, ILambdaContext context)
        {
            WorkBook workBook = WorkBook.Open(ExcelFileFormat.XLS);
            var newSheet = workBook.CreateWorkSheet("new_sheet");

            string columnNames = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            foreach (char column in columnNames)
            {
                for (int row = 1; row <= 50; row++)
                {
                    var cellAddress = $"{column}{row}";
                    newSheet[cellAddress].Value = $"Cell: {cellAddress}";
                }
            }

            return Convert.ToBase64String(workBook.ToByteArray());
        }
    }
}
```

For information on availability and the setup of IronXL's NuGet packages, please visit the documentation at [**IronXL NuGet Installation Guide**](https://ironsoftware.com/csharp/excel/docs/).