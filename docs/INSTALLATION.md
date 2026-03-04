# Oneberry Installation Guide

This document provides detailed instructions for installing Oneberry on various platforms.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation Methods](#installation-methods)
   - [Building from Source](#building-from-source)
   - [Docker Installation](#docker-installation)
   - [Pre-built Packages](#pre-built-packages)
3. [Platform-Specific Instructions](#platform-specific-instructions)
   - [Debian/Ubuntu](#debianubuntu)
   - [Fedora/RHEL/CentOS](#fedorarhel-centos)
   - [Arch Linux](#arch-linux)
   - [Ingenic A1](#ingenic-a1)
   - [Raspberry Pi](#raspberry-pi)
4. [Post-Installation Setup](#post-installation-setup)
5. [Upgrading](#upgrading)
6. [Uninstallation](#uninstallation)

## Prerequisites

Before installing Oneberry, ensure your system meets the following requirements:

- **Processor**: Any Linux-compatible processor (ARM, x86, MIPS, etc.)
- **Memory**: unknown what the minimum is
- **Storage**: Any storage device accessible by the OS
- **Network**: Ethernet or WiFi connection
- **OS**: Linux with kernel 4.4 or newer

## Installation Methods

### Building from Source

Building from source is the recommended method for most installations, as it ensures compatibility with your specific system.

#### 1. Clone the Repository

```bash
git clone https://github.com/opensensor/oneberry.git
cd oneberry
```

#### 2. Install Dependencies

See the [Platform-Specific Instructions](#platform-specific-instructions) section for dependency installation commands for your distribution.

#### 3. Build the Software

```bash
# Build in debug mode (default)
./scripts/build.sh

# Or build in release mode (recommended for production)
./scripts/build.sh --release
```

#### 4. Install the Software

```bash
# Install (requires root privileges)
sudo ./scripts/install.sh
```

The installation script will:
1. Install the binary to `/usr/local/bin/oneberry`
2. Install configuration files to `/etc/oneberry/`
3. Create data directories in `/var/lib/oneberry/`
4. Create a systemd service file

You can customize the installation paths using options:

```bash
sudo ./scripts/install.sh --prefix=/opt --config-dir=/etc/custom/oneberry
```

See `./scripts/install.sh --help` for all available options.

### Docker Installation

Docker provides an easy way to run Oneberry without installing dependencies directly on your system.

#### Option 1: Using Docker Compose (Recommended)

Docker Compose simplifies the deployment and ensures proper volume configuration.

```bash
# Clone the repository
git clone https://github.com/opensensor/oneberry.git
cd oneberry

# Start the container
docker-compose up -d
```

The default `docker-compose.yml` creates two volumes:
- `./config` - Configuration files (mounted to `/etc/oneberry`)
- `./data` - Persistent data including database, recordings, and models (mounted to `/var/lib/oneberry/data`)

To customize the configuration:

```bash
# Edit the configuration file
nano config/oneberry.ini

# Restart the container to apply changes
docker-compose restart
```

#### Option 2: Using Docker Run

##### 1. Pull the Docker Image

```bash
docker pull oneberry/oneberry:latest
```

##### 2. Create Directories for Persistent Storage

```bash
mkdir -p /path/to/config
mkdir -p /path/to/data
```

**Important:** The data directory must be persisted to avoid losing the database and recordings on container restart.

##### 3. Run the Container

```bash
docker run -d \
  --name oneberry \
  --restart unless-stopped \
  -p 8080:8080 \
  -p 1984:1984 \
  -v /path/to/config:/etc/oneberry \
  -v /path/to/data:/var/lib/oneberry/data \
  oneberry/oneberry:latest
```

##### 4. Create a Configuration File

```bash
# Copy the default configuration
docker cp oneberry:/etc/oneberry/oneberry.ini /path/to/config/oneberry.ini

# Edit the configuration
nano /path/to/config/oneberry.ini
```

**Note:** Ensure the paths in `oneberry.ini` point to `/var/lib/oneberry/data` subdirectories:
- Database: `/var/lib/oneberry/data/database/oneberry.db`
- Recordings: `/var/lib/oneberry/data/recordings`
- MP4 recordings: `/var/lib/oneberry/data/recordings/mp4`
- Models: `/var/lib/oneberry/data/models`

##### 5. Restart the Container

```bash
docker restart oneberry
```

### Pre-built Packages

Pre-built packages are available from GitHub Releases.

#### Downloading from GitHub Releases

1. Visit the [Oneberry Releases page](https://github.com/opensensor/oneberry/releases)
2. Download the appropriate package for your platform:
   - `.deb` packages for Debian/Ubuntu
   - `.tar.gz` archives for other Linux distributions

#### Debian/Ubuntu

```bash
# Download the latest .deb package from GitHub Releases
wget https://github.com/opensensor/oneberry/releases/latest/download/lightnvr_<version>_<arch>.deb

# Install the package
sudo dpkg -i lightnvr_<version>_<arch>.deb

# Install any missing dependencies
sudo apt-get install -f
```

#### Other Distributions

For other distributions, download the tarball and install manually:

```bash
# Download and extract
wget https://github.com/opensensor/oneberry/releases/latest/download/oneberry-<version>-linux-<arch>.tar.gz
tar -xzf oneberry-<version>-linux-<arch>.tar.gz

# Install (adjust paths as needed)
sudo cp oneberry /usr/local/bin/
sudo mkdir -p /etc/oneberry
sudo cp oneberry.ini.default /etc/oneberry/oneberry.ini
```

## Platform-Specific Instructions

### Debian/Ubuntu

#### Install Dependencies

```bash
sudo apt-get update
sudo apt-get install -y \
    build-essential \
    cmake \
    pkg-config \
    git \
    libsqlite3-dev \
    libavcodec-dev \
    libavformat-dev \
    libavutil-dev \
    libswscale-dev \
    libcurl4-openssl-dev \
    libmbedtls-dev \
    curl \
    wget
```

**Note**: `libmbedtls-dev` is **required** for ONVIF support and authentication system (cryptographic functions).

### Fedora/RHEL/CentOS

#### Install Dependencies

```bash
sudo dnf install -y \
    gcc \
    gcc-c++ \
    make \
    cmake \
    pkgconfig \
    git \
    sqlite-devel \
    ffmpeg-devel \
    libcurl-devel \
    mbedtls-devel \
    curl \
    wget
```

**Note**: `mbedtls-devel` is **required** for ONVIF support and authentication system (cryptographic functions).

### Arch Linux

#### Install Dependencies

```bash
sudo pacman -S \
    base-devel \
    cmake \
    git \
    sqlite \
    ffmpeg \
    curl \
    wget \
    mbedtls
```

**Note**: `mbedtls` is **required** for ONVIF support and authentication system (cryptographic functions).

### Ingenic A1

The Ingenic A1 SoC requires cross-compilation. A detailed guide is provided below.

#### 1. Set Up Cross-Compilation Toolchain

```bash
# Download and extract the toolchain
wget https://github.com/Ingenic-community/mips-linux-toolchain/releases/download/latest/mips-linux-uclibc-toolchain.tar.gz
sudo mkdir -p /opt/mips-linux-toolchain
sudo tar -xzf mips-linux-uclibc-toolchain.tar.gz -C /opt/mips-linux-toolchain
```

#### 2. Install Dependencies for Cross-Compilation

```bash
# Install build tools
sudo apt-get update
sudo apt-get install -y build-essential cmake pkg-config

# Clone and build cross-compiled dependencies
git clone https://github.com/oneberry/ingenic-dependencies.git
cd ingenic-dependencies
./build-all.sh
```

#### 3. Build Oneberry for Ingenic A1

```bash
# Clone the repository
git clone https://github.com/opensensor/oneberry.git
cd oneberry

# Build using the cross-compilation script
./scripts/build-ingenic.sh
```

#### 4. Deploy to Ingenic A1 Device

```bash
# Copy the binary and configuration files to the device
scp build/Ingenic/bin/oneberry root@ingenic-device:/usr/local/bin/
scp config/oneberry.conf.default root@ingenic-device:/etc/oneberry/oneberry.conf

# Create necessary directories on the device
ssh root@ingenic-device "mkdir -p /var/lib/oneberry/recordings /var/lib/oneberry/www /var/log/oneberry"

# Copy web interface files
scp -r web/* root@ingenic-device:/var/lib/oneberry/www/
```

### Raspberry Pi

Raspberry Pi installation is similar to Debian/Ubuntu, with some optimizations.

#### 1. Install Dependencies

```bash
sudo apt-get update
sudo apt-get install -y \
    build-essential \
    cmake \
    pkg-config \
    git \
    libsqlite3-dev \
    libavcodec-dev \
    libavformat-dev \
    libavutil-dev \
    libswscale-dev \
    libcurl4-openssl-dev \
    libmbedtls-dev \
    curl \
    wget
```

**Note**: `libmbedtls-dev` is **required** for ONVIF support and authentication system (cryptographic functions).

#### 2. Build and Install

```bash
# Clone the repository
git clone https://github.com/opensensor/oneberry.git
cd oneberry

# Build with Raspberry Pi optimizations
./scripts/build.sh --release --platform=raspberry-pi

# Install
sudo ./scripts/install.sh
```

## Post-Installation Setup

After installing Oneberry, follow these steps to complete the setup:

### 1. Configure Oneberry

Edit the configuration file:

```bash
sudo nano /etc/oneberry/oneberry.conf
```

At minimum, you should:
- Set a secure password for the web interface
- Configure storage paths
- Set up your camera streams

See [CONFIGURATION.md](CONFIGURATION.md) for detailed configuration options.

### 2. Start the Service

```bash
# Start the service
sudo systemctl start oneberry

# Enable the service to start at boot
sudo systemctl enable oneberry
```

### 3. Check the Status

```bash
sudo systemctl status oneberry
```

### 4. Access the Web Interface

Open a web browser and navigate to:

```
http://your-device-ip:8080
```

Log in with the username and password configured in the configuration file.

## Upgrading

### Upgrading from Source

```bash
# Navigate to the repository
cd oneberry

# Pull the latest changes
git pull

# Rebuild
./scripts/build.sh --release

# Stop the service
sudo systemctl stop oneberry

# Install the new version
sudo ./scripts/install.sh

# Start the service
sudo systemctl start oneberry
```

### Upgrading Docker Installation

#### Using Docker Compose

```bash
# Navigate to the repository
cd oneberry

# Pull the latest changes
git pull

# Rebuild and restart
docker-compose down
docker-compose build
docker-compose up -d
```

#### Using Docker Run

```bash
# Pull the latest image
docker pull oneberry/oneberry:latest

# Stop and remove the container
docker stop oneberry
docker rm oneberry

# Run a new container with the latest image
docker run -d \
  --name oneberry \
  --restart unless-stopped \
  -p 8080:8080 \
  -p 1984:1984 \
  -v /path/to/config:/etc/oneberry \
  -v /path/to/data:/var/lib/oneberry/data \
  oneberry/oneberry:latest
```

**Note:** Your data is preserved in the volumes, so upgrading will not affect your database or recordings.

## Uninstallation

### Uninstalling Source Installation

```bash
# Stop the service
sudo systemctl stop oneberry
sudo systemctl disable oneberry

# Remove the service file
sudo rm /etc/systemd/system/oneberry.service
sudo systemctl daemon-reload

# Remove the binary
sudo rm /usr/local/bin/oneberry

# Remove configuration and data (optional)
sudo rm -rf /etc/oneberry
sudo rm -rf /var/lib/oneberry
sudo rm -rf /var/log/oneberry
```

### Uninstalling Docker Installation

#### Using Docker Compose

```bash
# Navigate to the repository
cd oneberry

# Stop and remove the container
docker-compose down

# Remove the image
docker rmi oneberry

# Remove volumes (optional - this will delete all data)
rm -rf ./config
rm -rf ./data
```

#### Using Docker Run

```bash
# Stop and remove the container
docker stop oneberry
docker rm oneberry

# Remove the image
docker rmi oneberry/oneberry:latest

# Remove volumes (optional - this will delete all data)
rm -rf /path/to/config
rm -rf /path/to/data
```
