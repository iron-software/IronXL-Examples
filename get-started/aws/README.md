# IronXL Integration with AWS Lambda Functions Using .NET Core

***Based on <https://ironsoftware.com/get-started/aws/>***


IronXL offers full support for implementing AWS Lambda Functions across .NET Standard Libraries, as well as for .NET Core applications, .NET 5, and .NET 6 frameworks.

To integrate AWS ToolKit in Visual Studio, please refer to the instructions here: [Using the AWS Lambda Templates in the AWS Toolkit for Visual Studio](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

Once you install the AWS Toolkit into Visual Studio, you are equipped to develop an AWS Lambda Function Project. For a step-by-step guide on constructing an AWS Lambda Project using Visual Studio, click this [link](https://docs.aws.amazon.com/toolkit-for-visual-studio/latest/user-guide/lambda-creating-project-in-visual-studio.html).

### Sample AWS Lambda Function Code with IronXL

Following the setup of your AWS Lambda Function project in Visual Studio, you can utilize the following example code:

```csharp
using System;
using Amazon.Lambda.Core;
using IronXL;

// This attribute indicates the target of your Lambda function
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.Json.JsonSerializer))]

namespace AWSLambdaIronXL
{
    public class Function
    {
        /// This method processes a string input and utilizes IronXL to manipulate an Excel workbook.
        /// In this example, a new Excel workbook is created and populated with labeled data in cells.
        /// <param name="input">The input string provided to the function</param>
        /// <param name="context">Lambda context provided during execution</param>
        /// <returns>A string that represents the Excel file in Base64 format</returns>
        public string FunctionHandler(string input, ILambdaContext context)
        {
            // Initialize a new Excel workbook in the XLS format
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLS);
            
            // Create and name a new worksheet
            var newSheet = workBook.CreateWorkSheet("new_sheet");
            
            string columnNames = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            foreach (char col in columnNames)
            {
                for (int row = 1; row <= 50; row++)
                {
                    // Generate cell identifier and populate it
                    var cellName = $"{col}{row}";
                    newSheet[cellName].Value = $"Cell: {cellName}";
                }
            }

            // Convert the Excel workbook to a Base64-encoded string
            return Convert.ToBase64String(workBook.ToByteArray());
        }
    }
}
```

For details on the IronXL NuGet package and deployment information, refer to our comprehensive [IronXL NuGet Installation Guide](https://ironsoftware.com/csharp/excel/docs/).