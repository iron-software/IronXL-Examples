# IronXL Linux Compatibility & Setup Instruction

***Based on <https://ironsoftware.com/how-to/linux/>***


IronXL is engineered exclusively with .NET Standard, enabling it to function seamlessly across all Linux distributions that support **.NET Core**, **.NET 5**, and **.NET 6**. Furthermore, it operates flawlessly on Docker, Azure, macOS, and Windows environments that support the .NET frameworks.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu">
    <img src="https://img.icons8.com/color/96/000000/debian--v1.png" alt="Debian">
</div>

We advise deploying IronXL on .NET Core 3.1, .NET Core 5, or .NET Core 6, especially those versions designated as [LTS by Microsoft](https://dotnet.microsoft.com/platform/support/policy) due to their extended support and thorough testing on Linux platforms.

IronXL is designed to work immediately upon installation, requiring no modifications to code on Linux, thanks to extensive testing and configurations performed by our team of engineers.

Support for Linux is crucial as numerous cloud services like Azure Web Apps, Azure Functions, AWS EC2, AWS Lambda, and Docker on Azure Devops predominantly utilize Linux. At Iron Software, we frequently utilize these cloud services and acknowledge their importance to our Enterprise and SAAS clientele.

### Full Support for All Linux Distros Compliant with .NET

We **fully support** and suggest the newest **64-bit** versions of Linux listed below for straightforward "zero configuration" installation of IronXL:

*   Ubuntu 20
*   Ubuntu 18
*   Debian 11
*   Debian 10 _\[The Default Linux Distribution on Microsoft Azure\]_
*   CentOS 7
*   CentOS 8

Please consult "Other Linux Distros" later in this document for guidance on installing IronXL on an unsupported Linux version.

It's recommended to use Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/) for seamless deployment. Other Linux distributions may partly support IronXL but could necessitate manual setup using `apt-get`. Refer to "Linux Manual Setup" at the document’s conclusion for more details.

## IronXL NuGet Packages

```shell
Install-Package IronXL.Excel
```

## Ubuntu Compatibility

As one of our most frequently tested operating systems due to its extensive use in Azure's infrastructure, Ubuntu provides an optimal environment for continuous testing and deployment. This platform benefits from official Microsoft .NET support and readily available Official Docker Images.

### Ubuntu 20

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/48/000000/microsoft.png" alt="Microsoft">
    <img src="https://img.icons8.com/color/48/000000/ubuntu--v1.png" alt="Ubuntu">
    <img src="https://img.icons8.com/color/48/000000/chrome--v1.png" alt="Chrome">
    <img src="https://img.icons8.com/color/48/000000/safari--v1.png" alt="Safari">
    <img src="https://img.icons8.com/color/48/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/48/000000/azure-1.png" alt="Azure">
</div>

**Official Microsoft Docker Images:**

*   [64-bit Ubuntu 20.04 Docker Image for .NET Runtime 3.1 ('3.1-focal')](https://hub.docker.com/_/microsoft-dotnet-runtime/)
*   [64-bit Ubuntu 20.04 Docker Image for .NET Runtime 5.0 ('5.0-focal')](https://hub.docker.com/_/microsoft-dotnet-runtime/)

### Ubuntu 18

The structure and content for Ubuntu 18 are largely similar to Ubuntu 20, with high compatibility and extensive support, despite the absence of an official Docker image for .NET 5.

### Debian 11 and Debian 10

Both Debian 11 and Debian 10 are significantly supported with official Docker images suited for various .NET runtimes. Debian 10 is notably the default choice in Visual Studio when configuring Docker support for .NET projects.

**CentOS 7 & CentOS 8** are well-integrated for use with IronXL; installing the necessary NuGet package is sufficient for operation without any special configuration.

**Other Linux Distributions** supporting .NET will also be compatible with IronXL following the same straightforward NuGet package installation, provided administrative privileges are available.