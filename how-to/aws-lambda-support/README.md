# IronXL Integration with AWS Lambda using .NET Core

***Based on <https://ironsoftware.com/how-to/aws-lambda-support/>***


IronXL offers complete compatibility with AWS Lambda for .NET Standard Libraries, and .NET Core, including .NET 5 and .NET 6 projects.

For guidance on adding the AWS Toolkit to Visual Studio, refer to the comprehensive tutorial here: [Using the AWS Lambda Templates in the AWS Toolkit for Visual Studio](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

Installing the AWS Toolkit in Visual Studio simplifies the process of creating a project tailored for AWS Lambda. Detailed instructions on setting up your AWS Lambda project in Visual Studio can be found following this [link](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

### Example Code for an AWS Lambda Function

Once you have established a new AWS Lambda project, consider the following code sample:
```cs
    namespace AWSLambdaIronXL
    {
        public class Function
        {
        
        /// <summary>
        /// A simple method receiving a string and converting it to uppercase
        /// </summary>
        /// <param name="input"></param>
        /// <param name="context"></param>
        /// <returns>Base64 encoded string of an Excel file</returns>
        public string FunctionHandler(string input, ILambdaContext context)
        {
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLS);

            var newSheet = workbook.CreateWorkSheet("new_sheet");
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            foreach (char column in alphabet)
            {
                for (int row = 1; row <= 50; row++)
                {
                    var cellName = $"{column}{row}";
                    newSheet[cellName].Value = $"Cell: {cellName}";
                }
            }

            return Convert.ToBase64String(workbook.ToByteArray());
        }
        }
    }
```

For deployment, IronXL NuGet Packages are detailed in our documentation, which can be accessed [here](https://ironsoftware.com/csharp/excel/docs/).