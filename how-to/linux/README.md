# IronXL Linux Compatibility & Setup Guide

***Based on <https://ironsoftware.com/how-to/linux/>***


IronXL is designed with pure .NET Standard, ensuring it operates seamlessly on any Linux distribution that supports `.NET Core`, `.NET 5`, and `.NET 6`. Additionally, it functions perfectly on Docker, Azure, macOS (all of which support .NET frameworks), and Windows.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu">
    <img src="https://img.icons8.com/color/96/000000/debian--v1.png" alt="Debian">
</div>

We advocate for the use of `.NET Core 3.1`, `.NET 5` or `.NET 6`, especially versions endorsed as LTS (Long Term Support) by Microsoft, as highlighted here [Microsoft's Support Policy](https://dotnet.microsoft.com/platform/support/policy), due to their robustness and ensured long-term support when running on Linux.

IronXL requires no modifications to function on Linux, providing an immediate, ready-to-use solution thanks to extensive testing and fine-tuning by our dedicated engineering team.

The significance of Linux compatibility cannot be understated, as many significant cloud services like Azure Web Apps, Azure Functions, AWS EC2, AWS Lambda, and Azure DevOps Docker predominantly rely on Linux. At Iron Software, our routine use of these services aligns with the requirements of our Enterprise and SAAS clientele.

### Officially Supported All Linux Distros That Support .NET

IronXL **officially supports** the use on the latest **64-bit** Linux operating systems listed below, offering a straightforward, "zero configuration" installation process:

* Ubuntu 20
* Ubuntu 18
* Debian 11
* Debian 10 _\[Currently the Microsoft Azure Default Linux Distro\]_
* Centos 7
* Centos 8

For installations on Linux distributions not officially supported, please refer to the "Other Linux Distros" section further in this document for setup advice.

We recommend using Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/) for optimal compatibility. Other Linux distributions may be supported but could require manual configuration via apt-get. Further details are provided in the "Linux Manual Setup" at the end of this document.

## IronXL NuGet Packages

```shell
Install-Package IronXL.Excel
```

## Ubuntu Compatibility

Ubuntu, being the most extensively tested Linux OS in our framework, is heavily utilized within the Azure ecosystem that supports our continuous testing and deployment processes. This platform benefits from official Microsoft .NET support and Official Docker Images.

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

* [64 bit Ubuntu 20.04 Docker Image for .NET Runtime 3.1 ('3.1-focal')](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64 bit Ubuntu 20.04 Docker Image for .NET Runtime 5.0 ('5.0-focal')](https://hub.docker.com/_/microsoft-dotnet-runtime/)

### Ubuntu 18

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/48/000000/microsoft.png" alt="Microsoft">
    <img src="https://img.icons8.com/color/48/000000/ubuntu--v1.png" alt="Ubuntu">
    <img src="https://img.icons8.com/color/48/000000/chrome--v1.png" alt="Chrome">
    <img src="https://img.icons8.com/color/48/000000/safari--v1.png" alt="Safari">
    <img src="https://img.icons8.com/color/48/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/48/000000/azure-1.png" alt="Azure">
</div> 

**Official Microsoft Docker Images:**

* [64 bit Ubuntu 18.04 Docker Image for .NET Runtime 3.1 ('3.1-bionic')](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* High compatibility with .NET 5 without an official Docker image.

### Debian 11 and Debian 10

*Debian 10 remains the default Linux distribution for Microsoft* when integrating Docker support in .NET projects via Visual Studio.

**Official Microsoft Docker Images for Debian:**

* [64 bit Debian 11 Docker Image for .NET Runtime 3.1](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64 bit Debian 11 Docker Image for .NET Runtime 5.0](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64 bit Debian 10 Docker Image for .NET Runtime 3.1](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64 bit Debian 10 Docker Image for .NET Runtime 5.0](https://hub.docker.com/_/microsoft-dotnet-runtime/)

**CentOS 7 & CentOS 8** Only require administrative privileges with _sudo_ to install IronXL via NuGet, with no special configuration needed.

**Other Linux Distros** Just ensure compatibility with .NET and administrative privileges, and IronXL can be installed and run directly with no special configuration required.