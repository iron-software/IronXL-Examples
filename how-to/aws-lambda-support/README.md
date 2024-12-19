# IronXL Integration for AWS Lambda with .NET Core

***Based on <https://ironsoftware.com/how-to/aws-lambda-support/>***


IronXL provides robust support for AWS Lambda functions across .NET Standard libraries, .NET Core, .NET 5, and .NET 6 projects.

For incorporating the AWS Toolkit into Visual Studio, consult the documentation here: [Working with AWS Lambda Templates in the AWS Toolkit for Visual Studio](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

Once you install the AWS Toolkit in Visual Studio, you can begin setting up your AWS Lambda Function Project. Detailed steps on creating this project in Visual Studio can be found at this [link](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

### Sample AWS Lambda Function Using IronXL

Upon setting up a new AWS Lambda Function project, consider implementing the following code snippet:
```cs
    namespace AWSLambdaIronXL
    {
        public class Function
        {
        /// <summary>
        /// A straightforward function that converts a string to uppercase
        /// </summary>
        /// <param name="input">Input string</param>
        /// <param name="context">Lambda context</param>
        /// <returns>Base64 representation of the Excel file</returns>
        public string FunctionHandler(string input, ILambdaContext context)
        {
            WorkBook workbook = WorkBook.LoadOrCreate(ExcelFileFormat.XLS); // Ensures a workbook is created if not existing

            var sheet = workbook.CreateWorkSheet("new_sheet");
            string columnNames = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            foreach (char column in columnNames)
            {
                for (int row = 1; row <= 50; row++)
                {
                    var cell = $"{column}{row}";
                    sheet[cell].Value = $"Content: {cell}";
                }
            }

            return Convert.ToBase64String(workbook.ToByteArray());
        }
        }
    }
```

For information regarding deploying IronXL with NuGet packages, refer to the [IronXL NuGet Installation Guide](https://ironsoftware.com/csharp/excel/docs/).