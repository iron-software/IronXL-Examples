# Working with IronXL License Keys

***Based on <https://ironsoftware.com/how-to/license-keys/>***


## Obtaining a License Key

Deploying your project without any limitations or watermarks requires an IronXL license key.

You can [purchase a license here](https://ironsoftware.com/csharp/excel/licensing/) or opt for a [free 30-day trial key here](https://ironsoftware.com/#trial-license), accessible through an interactive modal.

---

## Step 1: Install the Latest IronXL

Begin by incorporating the IronXL.Excel library to enhance your .NET project with Excel capabilities.

### Installation via NuGet Package

1. Open Visual Studio, right-click on your project and select 'Manage NuGet Packages...'
2. Look for the IronXL.Excel package and proceed to install it

Alternatively,

1. Open the Package Manager Console
2. Execute the command below to install the package:

```shell
Install-Package IronXL.Excel
```

Check out [the package on the NuGet website](https://www.nuget.org/packages/IronXL.Excel/).

### Installation via Direct DLL Download

[Download the IronXL .NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually integrate it into Visual Studio.

---

## Step 2: Activating Your License Key

### Embed the License Key in Your Code

Ensure to insert this line of code at your application's startup, prior to utilizing IronXL.

```cs
IronXL.License.LicenseKey = "IRONXL-MYLICENSE-KEY-1EF01";
```

### Configuration File Integration in .NET Framework Applications

To globally activate your key within a .NET Framework application using a configuration file:

```xml
<configuration>
  ...
  <appSettings>
    <add key="IronXL.LicenseKey" value="IRONXL-MYLICENSE-KEY-1EF01"/>
  </appSettings>
  ...
</configuration>
```

Please note an ongoing licensing issue affecting versions [2023.4.13](https://www.nuget.org/packages/IronXL.Excel/2023.4.13) through [2024.3.20](https://www.nuget.org/packages/IronXL.Excel/2024.3.20) for projects using:
- ASP.NET
- .NET Framework version 4.6.2 or higher

The configuration details provided in a `Web.config` file may not be recognized. For guidance, refer to the ['Setting License Key in Web.config'](https://ironsoftware.com/csharp/excel/troubleshooting/license-key-web.config/) article.

It's crucial to verify the licensing status with `IronXL.License.IsLicensed`.

### Applying Key in .NET Core Applications

To apply a license key to your .NET Core application:

* Add an 'appsettings.json' file at the root directory of your project
* Insert the 'IronXL.LicenseKey' with your key as its value
* Set the file properties to *Copy to Output Directory: Copy always*
* Confirm the licensing status by checking `IronXL.License.IsLicensed`

Example file: *appsettings.json*
```json
{
  "IronXL.LicenseKey": "IronXL-MYLICENSE-KEY-1EF01"
}
```

---

## Step 3: Verifying Your Key

Ensure the accuracy of your license application:
```cs
// Validate the license key string
bool result = IronXL.License.IsValidLicense("IRONXL-MYLICENSE-KEY-1EF01");

// Confirm successful IronXL licensing
bool is_licensed = IronXL.License.IsLicensed;
```

*Note:* Always clean and republish your application post licensing to prevent deployment issues.

---

## Step 4: Kickstart Your Project

Explore our step-by-step guide on [Getting Started with IronXL](https://ironsoftware.com/csharp/excel/docs/).

---

## Need Help?

For any inquiries, please contact [support@ironsoftware.com](mailto:support@ironsoftware.com).