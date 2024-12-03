# Configuring IronXL with Docker Containers

***Based on <https://ironsoftware.com/how-to/docker-support/>***


Learn how to [manage and manipulate Excel files using C# with IronXL](https://ironsoftware.com/csharp/excel/). IronXL provides full support for Docker, including on both Linux and Windows Azure Docker Containers.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/docker--v1.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services--v1.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/windows-logo--v1.png" alt="Windows">
</div>

## Why Opt for Docker?

Docker simplifies the packaging, transporting, and running of applications by using lightweight, standalone, executable containers. These containers are configurable to run on almost any system.

## Getting Started with IronXL and Docker on Linux

If you're new to Docker and .NET, we recommend starting with this informative guide on Docker integration and debugging with Visual Studio projects here: [Setting Up Docker with Visual Studio](https://docs.microsoft.com/en-us/visualstudio/containers/edit-and-refresh?view=vs-2019).

Explore our [IronXL Setup and Compatibility Guide for Linux](https://ironsoftware.com/csharp/excel/how-to/linux/) for detailed instructions.

### Recommended Docker Distributions for Linux

Here we list Linux operating systems that are fully compatible and easy to configure with IronXL:

- Ubuntu 20
- Ubuntu 18
- Debian 11
- Debian 10 (default Linux distribution on Microsoft Azure)
- CentOS 7
- CentOS 8

Use the [Microsoft Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/) for these setups. Partial support for other distributions is available, but additional manual configurations using `apt-get` could be necessary. Refer to our [Manual Linux Setup Guide](https://ironsoftware.com/csharp/excel/how-to/linux/).

**Sample Docker configurations for Ubuntu and Debian are provided below.**

## Key Installation Instructions for IronXL on Linux Using Docker

### Integrating IronXL via NuGet

It's a good practice to use IronXL via the NuGet Package when developing across different platforms like Windows, macOS, and Linux.
```shell
Install-Package IronXL.Excel
```

## Ubuntu Linux Docker Configuration Examples

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/docker--v1.png" alt="Docker">
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu">
</div>

### Setup for Ubuntu 20 using .NET 5

Here’s how you set up a project on Ubuntu 20 using .NET 5.0:

```dockerfile
# Base image with runtime (Ubuntu 20 with .NET 5.0 runtime)

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM mcr.microsoft.com/dotnet/runtime:5.0-focal AS base
WORKDIR /app

# Development base image (Ubuntu 20 with .NET 5.0 SDK)

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM mcr.microsoft.com/dotnet/sdk:5.0-focal AS build
WORKDIR /src
# Restore NuGet packages

***Based on <https://ironsoftware.com/how-to/docker-support/>***

COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"
# Build the project

***Based on <https://ironsoftware.com/how-to/docker-support/>***

COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build
# Publish the project

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish
# Final app setup

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

Similarly, configurations are detailed for Ubuntu 20 with .NET 3.1 LTS, Ubuntu 18 with .NET 3.1 LTS, Debian with various .NET versions, and CentOS 7 and 8 using .NET 3.1 LTS. Each setup provides a robust framework for deploying .NET applications in Docker environments tailored to specific Linux distributions.