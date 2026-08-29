# Setting up IronXL in Docker Containers

> Full guide: [Setting up IronXL in Docker Containers](https://ironsoftware.com/csharp/excel/get-started/docker/)


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

For an easy configuration of IronXL, opt for the latest 64-bit releases of these
Linux OSs, which are the distributions Microsoft currently ships official .NET 8
container images for:

- Ubuntu 24.04 (Noble)
- Ubuntu 22.04 (Jammy)
- Debian 12 (Bookworm) _[Currently the Microsoft Azure Default Linux Distro]_
- RHEL / UBI 8 or 9, for RHEL-family deployments. Microsoft no longer publishes
  official CentOS runtime images; CentOS 7 and 8 are both end of life.

We advise utilizing Microsoft's [Official Docker Images](https://hub.docker.com/_/microsoft-dotnet-runtime/). While other Linux distributions are compatible, they might require `apt-get` for manual setup. Refer to our "[Linux Manual Setup](https://ironsoftware.com/csharp/excel/how-to/linux/)" guide.

Included below are Docker files for Ubuntu, Debian and RHEL/UBI:

## Essential Steps for IronXL Linux Docker Installation

### Deploying via Our NuGet Package

It's recommended to install the [IronXL](https://www.nuget.org/packages/IronXL.Excel) NuGet Package, compatible with development on Windows, macOS, and Linux.

```shell
Install-Package IronXL.Excel
```

## DockerFiles for Ubuntu Linux

### Ubuntu 24.04 (Noble) with .NET 8 LTS

```dockerfile
# Base runtime image (Ubuntu 24.04 with .NET runtime)
FROM mcr.microsoft.com/dotnet/runtime:8.0-noble AS base
WORKDIR /app

# Base development image (Ubuntu 24.04 with .NET SDK)
FROM mcr.microsoft.com/dotnet/sdk:8.0-noble AS build
WORKDIR /src

# Restore NuGet packages
COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"

# Build project
COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build

# Publish project
FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish

# Run app
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

### Ubuntu 22.04 (Jammy) with .NET 8 LTS

```dockerfile
# Base runtime image (Ubuntu 22.04 with .NET runtime)
FROM mcr.microsoft.com/dotnet/runtime:8.0-jammy AS base
WORKDIR /app

# Base development image (Ubuntu 22.04 with .NET SDK)
FROM mcr.microsoft.com/dotnet/sdk:8.0-jammy AS build
WORKDIR /src

# Restore NuGet packages
COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"

# Build project
COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build

# Publish project
FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish

# Run app
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

## DockerFiles for Debian Linux

### Debian 12 (Bookworm) with .NET 8 LTS

```dockerfile
# Base runtime image (Debian 12 with .NET runtime)
FROM mcr.microsoft.com/dotnet/runtime:8.0-bookworm-slim AS base
WORKDIR /app

# Base development image (Debian 12 with .NET SDK)
FROM mcr.microsoft.com/dotnet/sdk:8.0-bookworm-slim AS build
WORKDIR /src

# Restore NuGet packages
COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"

# Build project
COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build

# Publish project
FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish

# Run app
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

## DockerFiles for RHEL and UBI

Microsoft no longer publishes official CentOS runtime images (CentOS 7 and 8 are both end-of-life). For RHEL-family deployments, use Red Hat's own Universal Base Image (UBI) .NET containers instead: they're freely usable without a Red Hat subscription and are kept current with supported .NET releases.

### RHEL UBI 8 with .NET 8 LTS

```dockerfile
# Base runtime image (UBI 8 with .NET 8 runtime)
FROM registry.access.redhat.com/ubi8/dotnet-80-runtime:8.0 AS base
WORKDIR /app

# Base development image (UBI 8 with .NET 8 SDK)
FROM registry.access.redhat.com/ubi8/dotnet-80:8.0 AS build
WORKDIR /src

# Restore NuGet packages
COPY ["Example/Example.csproj", "Example/"]
RUN dotnet restore "Example/Example.csproj"

# Build project
COPY . .
WORKDIR "/src/Example"
RUN dotnet build "Example.csproj" -c Release -o /app/build

# Publish project
FROM build AS publish
RUN dotnet publish "Example.csproj" -c Release -o /app/publish

# Run app
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Example.dll"]
```

## Frequently Asked Questions

**How can I set up IronXL in a Docker container?**
To set up IronXL in a Docker container, you need to use the IronXL NuGet package, which is compatible with Windows, macOS, and Linux. Install it using the command: dotnet add package IronXL. For Docker, integrate the package in your Dockerfile and ensure your application can access the necessary libraries and dependencies.

**What are the benefits of using Docker for Excel applications?**
Docker allows you to package, ship, and run Excel applications as lightweight, portable containers, ensuring consistency and efficiency across different environments. This helps in maintaining a stable and reproducible development and production environment.

**Which Linux distributions work best with IronXL in Docker?**
The recommended Linux distributions for configuring IronXL in Docker are Ubuntu 24.04, Ubuntu 22.04, Debian 12, and RHEL/UBI 8 or 9. These are the distributions Microsoft currently ships official .NET 8 container images for, providing a stable environment for Docker containers running Excel applications.

**Can I use IronXL in both Windows and Linux Docker containers?**
Yes, IronXL supports Docker containers on both Windows and Linux platforms. This includes containers hosted on Azure, allowing for flexible deployment options.

**What Docker images are recommended for .NET applications using IronXL?**
For .NET applications using IronXL, it is recommended to use Microsoft's official Docker images for .NET runtime and SDK. These images are optimized for .NET applications and can be found on Docker Hub.

**How can I troubleshoot issues with IronXL in a Docker environment?**
If you encounter issues with IronXL in Docker, ensure all dependencies are correctly installed within your Docker container. Check your Dockerfile configuration and ensure the correct .NET version is being used. Refer to the IronXL documentation and Docker's official troubleshooting guides for more help.

**What resources are available for learning more about Docker and IronXL integration?**
For further learning on Docker and IronXL integration, refer to Microsoft's documentation on Docker for .NET and Visual Studio projects. Additionally, IronXL’s Linux setup guide provides valuable information for setting up Docker environments.
