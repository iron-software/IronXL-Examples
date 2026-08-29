# IronXL Linux Compatibility & Setup Guide

> Full guide: [IronXL Linux Compatibility & Setup Guide](https://ironsoftware.com/csharp/excel/get-started/linux/?utm_source=github)


IronXL is engineered entirely using .NET Standard, enabling functionality across all Linux distributions that support **.NET Core**, **.NET 5**, and **.NET 6**. Compatibility extends to Docker, Azure, macOS platforms, and Windows—all of which support .NET frameworks.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu">
    <img src="https://img.icons8.com/color/96/000000/debian--v1.png" alt="Debian">
</div>

We suggest utilizing .NET Core 3.1, .NET 5, or .NET 6, along with other runtimes designated as [long-term support (LTS) by Microsoft](https://dotnet.microsoft.com/platform/support/policy) because of their testing and support on Linux platforms.

Running IronXL on Linux does not require any modifications to the code. IronXL typically operates flawlessly right off the bat, thanks to thorough testing and optimization by our development team.

Linux compatibility is crucial for many cloud environments such as Azure Web Apps, Azure Functions, AWS EC2, AWS Lambda, and containers in Azure DevOps, since these services predominantly utilize Linux. At Iron Software, we use these cloud services extensively and recognize their importance to our Enterprise and SAAS clients.

### Officially Supported Linux Distributions for .NET

IronXL **officially supports** the following latest **64-bit** Linux operating systems, on which no configuration is needed:

* Ubuntu 20
* Ubuntu 18
* Debian 11
* Debian 10 _\[Presently the default Linux distribution on Microsoft Azure\]_
* CentOS 7
* CentOS 8

For versions of Linux not **officially supported**, see the "Other Linux Distros" section below for guidance on installing IronXL.

We advise using Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/). Partial support exists for other Linux distributions, which may need manual setups via `apt-get`. Refer to "Linux Manual Setup" at the end of this document.

## IronXL NuGet Packages

```shell
# To install IronXL using the dotnet CLI, use the following command:

dotnet add package IronXL
```

## Ubuntu Compatibility

Ubuntu is our extensively tested Linux distribution because of its significant use in Azure infrastructures, which supports continuous testing and deployment. This environment is also backed by official Microsoft .NET and Docker support.

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
* Ubuntu 18 boasts high compatibility with .NET 5, although no official Docker image is available for this version.

### Debian 11

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/48/000000/debian.png" alt="Debian">
    <img src="https://img.icons8.com/color/48/000000/microsoft.png" alt="Microsoft">
    <img src="https://img.icons8.com/color/48/000000/chrome--v1.png" alt="Chrome">
    <img src="https://img.icons8.com/color/48/000000/safari--v1.png" alt="Safari">
    <img src="https://img.icons8.com/color/48/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/48/000000/azure-1.png" alt="Azure">
</div>

**Official Microsoft Docker Images:**

* [64-bit Debian 11 Docker Image for .NET Runtime 3.1](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64-bit Debian 11 Docker Image for .NET Runtime 5.0](https://hub.docker.com/_/microsoft-dotnet-runtime/)

### Debian 10

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/48/000000/debian.png" alt="Debian">
    <img src="https://img.icons8.com/color/48/000000/microsoft.png" alt="Microsoft">
    <img src="https://img.icons8.com/color/48/000000/chrome--v1.png" alt="Chrome">
    <img src="https://img.icons8.com/color/48/000000/safari--v1.png" alt="Safari">
    <img src="https://img.icons8.com/color/48/000000/docker.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/48/000000/azure-1.png" alt="Azure">
</div>

**Official Microsoft Docker Images:**

* [64-bit Debian 10 Docker Image for .NET Runtime 3.1](https://hub.docker.com/_/microsoft-dotnet-runtime/)
* [64-bit Debian 10 Docker Image for .NET Runtime 5.0](https://hub.docker.com/_/microsoft-dotnet-runtime/)

**CentOS 7 & CentOS 8:** It is important to have _sudo_ admin rights. No intricate configurations are necessary to start using IronXL; simply install the NuGet package and initiate.

**Other Linux Distros:** Confirm your Linux distro is compatible with .NET and you have _sudo_ admin privileges. Similar to CentOS, no detailed configuration is required; install the NuGet package to begin.