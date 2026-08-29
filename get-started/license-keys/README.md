# Utilizing IronXL License Keys

> Full guide: [Utilizing IronXL License Keys](https://ironsoftware.com/get-started/license-keys/)

## Acquiring a License Key

Implementing an IronXL license key in your project means you can go live without any operational restrictions or watermarks.

Purchase a license key on the [buy a license page](https://ironsoftware.com/csharp/excel/licensing/) or begin a [free 30-day trial here](https://ironsoftware.com/trial-license).

<hr class="separator">

The initial step is to integrate the IronXL.Excel library to enable Excel capabilities within the .NET framework.

### Installation via NuGet Package Manager

1. Within Visual Studio, right-click the project and choose "Manage NuGet Packages ..."
2. Locate the IronXL.Excel package and proceed with installation

Alternatively,

1. Open the Package Manager Console
2. Execute the command:

   ```shell
   Install-Package IronXL.Excel
   ```

[Inspect the NuGet package details here.](https://www.nuget.org/packages/IronXL.Excel/)

### Installation via DLL Direct Download

Directly download the IronXL [.NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually integrate it within Visual Studio.

<hr class="separator">

## Step 2: Implement Your License Key

### Embedding the license key in your code

Incorporate this snippet at the beginning of your application, prior to utilizing IronXL.

```csharp
// Initialize the IronXL license key for this project
IronXL.License.LicenseKey = "IRONXL-MYLICENSE-KEY-1EF01";
```

<hr class="separator">

### Configuration via Web.Config or App.Config in .NET Framework Applications

For a global application key setting using Web.Config or App.Config, include the following in the `appSettings` of your config file:

```xml
<configuration>
  ...
  <appSettings>
    
    <add key="IronXL.LicenseKey" value="IRONXL-MYLICENSE-KEY-1EF01"/>
  </appSettings>
  ...
</configuration>
```

A noted issue affects IronXL versions from [2023.4.13](https://www.nuget.org/packages/IronXL.Excel/2023.4.13) to [2024.3.20](https://www.nuget.org/packages/IronXL.Excel/2024.3.20) in projects:
- **ASP.NET**
- **.NET Framework version >= 4.6.2**

The key inside a `Web.config` is **NOT** recognized. Visit the [Troubleshooting License Key in Web.config](https://ironsoftware.com/csharp/excel/troubleshooting/license-key-web.config/) for details.

Ensure that `IronXL.License.IsLicensed` reports `true`.

<hr class="separator">

### Setting the Key in .NET Core via appsettings.json

For applying the key across your .NET Core application:

- Add an `appsettings.json` in your project's root directory
- Enter 'IronXL.LicenseKey' with your license key in your JSON configuration
- Set *Copy to Output Directory* in the file properties to *Copy always*
- Confirm with `IronXL.License.IsLicensed`.

File: *appsettings.json*

```json
{
  "IronXL.LicenseKey": "IRONXL-MYLICENSE-KEY-1EF01"
}
```

<hr class="separator">

## Step 3: Verify Your Key

Confirm the proper setup of your license key.

```csharp
// Validate the license key format
bool result = IronXL.License.IsValidLicense("IRONXL-MYLICENSE-KEY-1EF01");

// Ensure IronXL is correctly licensed
bool is_licensed = IronXL.License.IsLicensed;
```

*Important:* Always clean and republish your application after including a license to prevent deployment errors.

<hr class="separator">

## Step 4: Begin Your Project

Explore our guide on [Getting Started with IronXL](https://ironsoftware.com/csharp/excel/docs/).

<hr class="separator">

## Need Help?

For any inquiries, please contact [support@ironsoftware.com](mailto:support@ironsoftware.com).