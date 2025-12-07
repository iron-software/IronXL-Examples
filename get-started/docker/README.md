# Setting up IronXL in Docker Containers

***Based on <https://ironsoftware.com/get-started/docker/>***


Explore how to [interact with Excel spreadsheets using C#](https://ironsoftware.com/csharp/excel/)?

IronXL offers complete support for Docker environments, including Linux and Windows-based Azure Docker Containers.

<div class="main-content__small-images-inline"><img src="https://img.icons8.com/color/96/000000/docker--v1.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services--v1.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/windows-logo--v1.png" alt="Windows"></div>

## Why Choose Docker?

Docker simplifies the deployment process, enabling developers to package and deploy applications into lightweight, portable containers that can operate anywhere.

## Introduction to IronXL with Linux on Docker

New to Docker and .NET? Explore this informative resource on [debugging Docker with Visual Studio and integrating within projects](https://docs.microsoft.com/en-us/visualstudio/containers/edit-and-refresh?view=vs-2019).

We strongly suggest reviewing our [IronXL Linux Setup and Compatibility Guide](https://ironsoftware.com/csharp/excel/how-to/linux/).

### Preferred Linux Docker Distributions for IronXL

For an easy configuration of IronPDF, opt for the latest 64-bit releases of these Linux OSs:

- Ubuntu 20
- Ubuntu 18
- Debian 11
- Debian 10 _[Currently the Microsoft Azure Default Linux Distro]_
- CentOS 7
- CentOS 8

We advise utilizing Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/). While other Linux distributions are compatible, they might require `apt-get` for manual setup. Refer to our "[Linux Manual Setup](https://ironsoftware.com/csharp/excel/how-to/linux/)" guide.

Included below are Docker files for Ubuntu and Debian:

## Essential Steps for IronXL Linux Docker Installation

### Deploying via Our NuGet Package

It's recommended to install the [IronXL](https://www.nuget.org/packages/BarCode) NuGet Package, compatible with development on Windows, macOS, and Linux.

```shell
Install-Package IronXL.Excel
```

## DockerFiles for Ubuntu Linux

<div class="main-content__small-images-inline"><img src="https://img.icons8.com/color/96/000000/docker--v1.png" alt="Docker"> 
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu"></div>

### Ubuntu 20 Using .NET 5

```Dockerfile
# Base image for run-time environment (Ubuntu 20 with .NET runtime)

***Based on <https://ironsoftware.com/get-started/docker/>***

FROM mcr.microsoft.com/dotnet/runtime:5.0-focal AS base
WORKDIR /app

# Base image for development environment (Ubuntu 20 with .NET SDK)

***Based on <https://ironsoftware.com/get-started/docker/>***

FROM mcr.microsoft.com/dotnet/sdk:5.0-focal AS build
WORKDIR /src

# NuGet package restoration

***Based on <https://ironsoftware.com/get-started/docker/>***

COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"

# Compiling the project

***Based on <https://ironsoftware.com/get-started/docker/>***

COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build

# Publishing the project

***Based on <https://ironsoftware.com/get-started/docker/>***

FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish

# Executing the application

***Based on <https://ironsoftware.com/get-started/docker/>***

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

### Ubuntu 20 with .NET 3.1 LTS

```Dockerfile
# Base runtime image (for Ubuntu 20 with .NET 3.1 LTS)

***Based on <https://ironsoftware.com/get-started/docker/>***

FROM mcr.microsoft.com/dotnet/runtime:3.1-focal AS base
WORKDIR /app

# Developing environment setup (for Ubuntu 20 with .NET 3.1 SDK)

***Based on <https://ironsoftware.com/get-started/docker/>***

FROM mcr.microsoft.com/dotnet/sdk:3.1-focal AS build
WORK