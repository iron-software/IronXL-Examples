# Using IronXL License Keys

***Based on <https://ironsoftware.com/how-to/license-keys/>***


## Obtaining a License Key

Obtaining an IronXL license key allows you to move your project to a live environment devoid of any limitations or watermark impositions.

You can [purchase a license here](https://ironsoftware.com/csharp/excel/licensing/) or register for a [free 30-day trial key](https://ironsoftware.com/csharp/excel/licensing/).

<hr class="separator">

## Step 1: Install the Latest IronXL Version


The initial step is to integrate IronXL.Excel, which extends Excel capabilities to the .NET environment.

<h3>Via NuGet Package</h3>

1. In Visual Studio, right-click on your project and select "Manage NuGet Packages..."
2. Look for the IronXL.Excel package and install it

Alternatively,

1. Open the Package Manager Console
2. Execute: `Install-Package IronXL.Excel`

```shell
Install-Package IronXL.Excel
```

<br>
[Check out the package on NuGet](https://www.nuget.org/packages/IronXL.Excel/)

<h3>Manual DLL Download</h3>

Download the IronXL [.NET Excel DLL here](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually include it in your Visual Studio project.

<hr class="separator">

## Step 2: Activate Your License Key

### Embed the License Key in Your Code###

Insert this code at the beginning of your application, prior to employing IronXL.

```cs
IronXL.License.LicenseKey = "IRONXL-MYLICENSE-KEY-1EF01";
```

<hr class="separator">

### Apply Your License Using `Web.Config` or `App.Config` in .NET Framework Projects###

To assign the license key across your application via `Web.Config` or `App.Config`, insert this setting into your config file under `appSettings`.

```xml
<configuration>
...
  <appSettings>
    <add key="IronXL.LicenseKey" value="IronXL-MYLICENSE-KEY-1EF01"/>
  </appSettings>
...
</configuration>
```

Between IronXL version [2023.4.13](https://www.nuget.org/packages/IronXL.Excel/2023.4.13) and [2024.3.20](https://www.nuget.org/packages/IronXL.Excel/2024.3.20), there's an issue for:
- **ASP.NET** projects
- **.NET Framework version >= 4.6.2**

The key configured in `Web.config` might not be properly recognized. Learn more from the ['Setting License Key in Web.config'](https://ironsoftware.com/csharp/excel/troubleshooting/license-key-web.config/) guide.

Verify the license activation with `IronXL.License.IsLicensed`.

<hr class="separator">

### Implementing the Key in a .NET Core `appsettings.json` File###

To globally apply the key in .NET Core:

* Include a JSON file named `appsettings.json` at the root of your project
* Insert a 'IronXL.LicenseKey' setting. Assign your license key to this setting.
* Make sure the file property is set to *Copy to Output Directory: Copy always*
* Verify the license with `IronXL.License.IsLicensed`.

File: *appsettings.json*
```json
{

	"IronXL.LicenseKey":"IronXL-MYLICENSE-KEY-1EF01"

}  
```

<hr class="separator">

## Step 3: Verify Your License

Check if your license is effectively installed.

```cs
// Validate the license key.
bool result = IronXL.License.IsValidLicense("IronXL-MYLICENSE-KEY-1EF01");

// Determine if IronXL is officially licensed
bool is_licensed = IronXL.License.IsLicensed;
```

*Note:* Always clean and republish your application after license integration to avoid deployment errors.

<hr class="separator">

## Step 4: Kickstart Your Project

Follow our guide on how to [Get Started with IronXL](https://ironsoftware.com/csharp/excel/docs/).

<hr class="separator">

## Need Help?

Feel free to contact [support@ironsoftware.com](mailto:support@ironsoftware.com) with any inquiries.