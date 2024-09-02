# IronXL Linux Compatibility & Setup Guide

IronXL is designed with pure .NET Standard, making it compatible with all Linux distributions that support **.NET Core**, **.NET 5**, and **.NET 6**. Furthermore, it seamlessly integrates into environments such as Docker, Azure, macOS, and of course, Windows.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu">
    <img src="https://img.icons8.com/color/96/000000/debian--v1.png" alt="Debian">
</div>

It is advisable to use .NET Core versions 3.1, 5, or 6, especially those marked as [Long Term Support (LTS) by Microsoft](https://dotnet.microsoft.com/platform/support/policy), as they offer prolonged support and reliable performance on Linux platforms.

IronXL typically requires no code modifications to run effectively on Linux, thanks to extensive testing and optimization performed by our engineering team.

Linux compatibility is crucial, given its extensive use in major cloud services like Azure Web Apps, Azure Functions, AWS EC2, AWS Lambda, and Docker operations. Iron Software regularly employs these cloud tools, recognizing their importance to our Enterprise and SAAS clients.

### Fully Supported Linux Distributions for .NET

We **officially endorse** and recommend the following **64-bit** Linux operating systems for effortless configuration of IronXL:

* Ubuntu 20
* Ubuntu 18
* Debian 11
* Debian 10 _\[Currently the Microsoft Azure Default Linux Distro\]_
* Centos 7
* Centos 8

For installation on other Linux distributions that are not **officially supported**, please refer to the "Other Linux Distros" section below.

We suggest using Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/). For other distributions, partial support may be available, potentially requiring manual configuration through `apt-get`. More details can be found in the "Linux Manual Setup" section later in this guide.

## IronXL NuGet Package Installation

```shell
Install-Package IronXL.Excel
```

## Ubuntu Compatibility

Ubuntu is the platform we test most extensively due to its heavy usage in the Azure infrastructure, which supports our continuous testing and deployment processes. This system is fully backed by official Microsoft .NET support and Docker Images.

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

* [64-bit Ubuntu 20.04 Docker Image for .NET Runtime 3.1 ('3.1-focal')](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64-bit Ubuntu 20.04 Docker Image for .NET Runtime 5.0 ('5.0-focal')](https://hub.docker.com/_/microsoft-dotnet-runtime/)

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

* [64-bit Ubuntu 18.04 Docker Image for .NET Runtime 3.1 ('3.1-bionic')](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* Although there is no official docker image for .NET 5 on Ubuntu 18, compatibility remains high.

### Debian 11 and Debian 10

On the Debian front, both Debian 10 and 11 are extensively supported and integrated into Microsoft's Docker optimization strategies for .NET projects in Visual Studio.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/48/000000/debian.png" alt="Debian">
    <img src="https://img.icons8.com/color/48/000000/microsoft.png" alt="Microsoft">
    <img src="https://img.icons8.com/color/48/000000/chrome--v1.png" alt="Chrome">
    <img src="https://img.icons8.com/color/48/000000/safari--v1.png" alt="Safari">
    <img src="https://img.icons8.com/color/48/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/48/000000/azure-1.png" alt="Azure">
</div>

**Official Microsoft Docker Images for Debian 11 & 10:**

* [64-bit Debian 11 Docker Image for .NET Runtime 3.1](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64-bit Debian 11 Docker Image for .NET Runtime 5.0](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64-bit Debian 10 Docker Image for .NET Runtime 3.1](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64-bit Debian 10 Docker Image for.NET Runtime 5.0](https://hub.docker.com/_/microsoft-dotnet-runtime/)

**CentOS 7 & CentOS 8**

For CentOS installations, administrative rights are essential, but no special configurations are necessary. Simply install the NuGet package to get started.

**Other Linux Distributions**

Ensure the Linux distribution you choose supports .NET and grants you administrative access to avoid any complications. Installation of IronXL in these environments generally does not require additional configurations.