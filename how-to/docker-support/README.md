# Implementing IronXL with Docker Containers

Interested in [managing Excel spreadsheet files with C#](https://ironsoftware.com/csharp/excel/)?

IronXL is now fully compatible with Docker, supporting both Azure Docker Containers and those hosted on Linux and Windows environments.

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/docker--v1.png" alt="Docker">
    <img src="https://img.icons8.com/fluency/96/000000/azure-1.png" alt="Azure">
    <img src="https://img.icons8.com/color/96/000000/linux--v1.png" alt="Linux">
    <img src="https://img.icons8.com/color/96/000000/amazon-web-services--v1.png" alt="Amazon">
    <img src="https://img.icons8.com/color/96/000000/windows-logo--v1.png" alt="Windows">
</div>

## Benefits of Using Docker

Docker provides a streamlined approach for developers to package, deploy, and run applications using lightweight, stand-alone containers that work seamlessly across various computing environments.

## Getting Started with IronXL and Docker on Linux

For newcomers to Docker within the .NET ecosystem, this detailed guide on setting up Docker for debugging and integration with Visual Studio projects is invaluable: [https://docs.microsoft.com/en-us/visualstudio/containers/edit-and-refresh?view=vs-2019](https://docs.microsoft.com/en-us/visualstudio/containers/edit-and-refresh?view=vs-2019).

We also suggest consulting our [IronXL Linux Setup and Compatibility Guide](https://ironsoftware.com/csharp/excel/how-to/linux/).

### Recommended Linux Docker Environments

For optimal configuration with IronXL, we advocate using the latest 64-bit Linux OS versions noted below:

* Ubuntu 20
* Ubuntu 18
* Debian 11
* Debian 10 _\[The default Linux Distro on Microsoft Azure\]_
* CentOS 7
* CentOS 8

For best practices, utilize Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/). Some Linux distributions might need manual setups via `apt-get`. View our "[Linux Manual Setup](https://ironsoftware.com/csharp/excel/how-to/linux/)" documentation for assistance.

This document includes ready-to-use Docker files for Ubuntu and Debian:

## Essential Installation Steps for IronXL on Linux Using Docker

### NuGet Package

It's recommended to utilize the [IronXL](https://www.nuget.org/packages/BarCode) NuGet Package for development across Windows, macOS, and Linux.
```shell
Install-Package IronXL.Excel
```
## Docker Files for Ubuntu Linux

<div class="main-content__small-images-inline">
    <img src="https://img.icons8.com/color/96/000000/docker--v1.png" alt="Docker"> 
    <img src="https://img.icons8.com/color/96/000000/ubuntu--v1.png" alt="Ubuntu">
</div>

Below are Docker files for different versions of Ubuntu with specific .NET frameworks.

### Ubuntu 20 with .NET 5

```dockerfile
# Base runtime image (Ubuntu 20 with .NET runtime)
FROM mcr.microsoft.com/dotnet/runtime:5.0-focal AS base
WORKDIR /app

# Base development image (Ubuntu 20 with .NET SDK)
FROM mcr.microsoft.com/dotnet/sdk:5.0-focal AS build
WORKDIR /src
# Restore NuGet packages
COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"
# Build the project
COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build
# Publish the project
FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish
# Run the app
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

### Ubuntu 20 with .NET 3.1 LTS

```dockerfile
# Base runtime image (Ubuntu 20 with .NET runtime)
FROM mcr.microsoft.com/dotnet/runtime:3.1-focal AS base
WORKDIR /app

# Base development image (Ubuntu 20 with .NET SDK)
FROM mcr.microsoft.com/dotnet/sdk:3.1-focal AS build
WORKDIR /src
# Restore NuGet packages
COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"
# Build the project
COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build
# Publish the project
FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish
# Run the app
FROM base AS final
WORKDI