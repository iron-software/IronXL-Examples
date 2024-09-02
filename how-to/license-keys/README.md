# Using IronXL License Keys

## Acquiring a License Key

Activating an IronXL license key enables you to deploy your applications live without any restrictions or watermarks.

You can [purchase a license here](https://ironsoftware.com/csharp/excel/licensing/) or obtain a [free 30-day trial key here](https://ironsoftware.com/csharp/excel/licensing/).

<hr class="separator">

## Step 1: Install IronXL

To begin, we need to install the IronXL.Excel library to add Excel functionality to the .NET environment.

### Installation via NuGet Package

1. In Visual Studio, right-click on your project and select "Manage NuGet Packages..."
2. Search for `IronXL.Excel` and install it.

Alternatively,

1. Open the Package Manager Console.
2. Execute the command:

```shell
Install-Package IronXL.Excel
```

[View the package on NuGet](https://www.nuget.org/packages/IronXL.Excel/).

### Manual DLL Installation

Download the [IronXL .NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and add it manually to your Visual Studio project.

<hr class="separator">

## Step 2: Activate Your License Key

### Embedding the License Key in Your Code

Incorporate this line in your application's initialization code, before utilizing IronXL:

```cs
IronXL.License.LicenseKey = "IRONXL-MYLICENSE-KEY-1EF01";
```

<hr class="separator">

### Configuring Your Key via Web.Config or App.Config for .NET Framework Applications

To implement the license key across your application, insert it into your config file under appSettings:

```xml
<configuration>
...
  <appSettings>
    <add key="IronXL.LicenseKey" value="IRONXL-MYLICENSE-KEY-1EF01"/>
  </appSettings>
...
</configuration>
```

Please note that between IronXL versions [2023.4.13](https://www.nuget.org/packages/IronXL.Excel/2023.4.13) and [2024.3.20](https://www.nuget.org/packages/IronXL.Excel/2024.3.20), there is an issue with ASP.NET projects and .NET Framework versions >= 4.6.2 where the key in `Web.config` is not recognized. More information is available on the [license troubleshooting page](https://ironsoftware.com/csharp/excel/troubleshooting/license-key-web.config/).

Ensure that `IronXL.License.IsLicensed` evaluates to `true`.

<hr class="separator">

### Configuring Your Key in a .NET Core appsettings.json File

For global application settings in .NET Core:

- Create an `appsettings.json` file in your project's root directory.
- Add the `IronXL.LicenseKey` to your JSON configuration. The value should be your license key.
- Set the file properties to *Copy to Output Directory: Copy always*.
- Verify the license with `IronXL.License.IsLicensed`.

File example: *appsettings.json*

```json
{
	"IronXL.LicenseKey": "IronXL-MYLICENSE-KEY-1EF01"
}
```

<hr class="separator">

## Step 3: License Verification

Ensure your license key is properly activated:

```cs
// Validate if the license key is correct.
bool result = IronXL.License.IsValidLicense("IRONXL-MYLICENSE-KEY-1EF01");

// Confirm if IronXL is successfully licensed.
bool is_licensed = IronXL.License.IsLicensed;
```

*Note:* Always clean and republish your application after license configuration to ensure proper deployment.

<hr class="separator">

## Step 4: Begin Your Project

Explore our guide on [Getting Started with IronXL](https://ironsoftware.com/csharp/excel/docs/).

<hr class="separator">

## Need Help?

For any inquiries, contact us at [support@ironsoftware.com](mailto:support@ironsoftware.com).