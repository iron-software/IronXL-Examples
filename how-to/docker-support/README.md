# Implementing IronXL in Docker Environments

***Based on <https://ironsoftware.com/how-to/docker-support/>***


Discover how to [manage Excel files using C# in Docker containers](https://ironsoftware.com/csharp/excel/). IronXL offers seamless integration with Docker, fully supporting various environments such as Azure Docker Containers on both Linux and Windows platforms.

![Docker](https://img.icons8.com/color/96/000000/docker--v1.png) ![Azure](https://img.icons8.com/fluency/96/000000/azure-1.png) ![Linux](https://img.icons8.com/color/96/000000/linux--v1.png) ![Amazon](https://img.icons8.com/color/96/000000/amazon-web-services--v1.png) ![Windows](https://img.icons8.com/color/96/000000/windows-logo--v1.png)

## The Advantages of Using Docker

Docker simplifies the process of packaging, delivering, and running applications by using lightweight, standalone containers that can operate virtually anywhere.

## Getting Started with IronXL on Linux and Docker

For those new to Docker within the .NET framework, we suggest this comprehensive guide on [debugging and integrating Docker with Visual Studio](https://docs.microsoft.com/en-us/visualstudio/containers/edit-and-refresh?view=vs-2019).

Explore our detailed [guide on setting up IronXL with Linux](https://ironsoftware.com/csharp/excel/how-to/linux/).

### Suggested Linux Distros for Docker

The following 64-bit Linux distributions are recommended for straightforward installation of IronXL:

- Ubuntu 20
- Ubuntu 18
- Debian 11
- Debian 10 _(The default Linux Distro on Microsoft Azure)_
- CentOS 7
- CentOS 8

For optimal setup, consider using [Microsoft's Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/). For other Linux distributions, manual configuration might be necessary. Refer to our [Linux Manual Setup](https://ironsoftware.com/csharp/excel/how-to/linux/) for detailed instructions.

Find Dockerfiles for select Linux distributions mentioned below in this document.

## Essential Setup for IronXL on Linux Docker

### Utilizing the IronXL NuGet Package

Utilize the [IronXL NuGet Package](https://www.nuget.org/packages/BarCode) which proves effective across Windows, macOS, and Linux platforms.
```shell
Install-Package IronXL.Excel
```

## Ubuntu Linux DockerFiles

### Docker Configuration for Ubuntu 20 with .NET 5
```dockerfile
# Base runtime image (Ubuntu 20 w/ .NET runtime)

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM mcr.microsoft.com/dotnet/runtime:5.0-focal AS base
WORKDIR /app

# Base development image (Ubuntu 20 w/ .NET SDK)

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM mcr.microsoft.com/dotnet/sdk:5.0-focal AS build
WORKDIR /src
# Restore NuGet packages

***Based on <https://ironsoftware.com/how-to/docker-support/>***

COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"
# Build project

***Based on <https://ironsoftware.com/how-to/docker-support/>***

COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build
# Publish project

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish
# Run app

***Based on <https://ironsoftware.com/how-to/docker-support/>***

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```
Repeat the above steps for other versions of Ubuntu and Debian as well as for CentOS, adapting the Docker file directives to the specific Linux version and .NET framework as demonstrated.

This guide outlines the complete steps to efficiently deploy IronXL within Docker containers, utilizing different Linux distributions, while ensuring compatibility and performance.