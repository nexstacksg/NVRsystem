# Oneberry Quickstart Guide

This guide is for new users who want a smooth, step-by-step way to install Oneberry NVR with motion detection.

> [!NOTE]
> Every command below is required. You must run them in your terminal, one after another, to build and install the software onto your computer.

---

## 1. Install Prerequisites (Required)
These are the tools your computer needs to build and run the software.

**Run this command to install everything at once:**
```bash
sudo apt update && sudo apt install -y \
    build-essential cmake pkg-config git \
    libsqlite3-dev libavcodec-dev libavformat-dev \
    libavutil-dev libswscale-dev libswresample-dev \
    libcurl4-openssl-dev libmbedtls-dev libcjson-dev \
    libmosquitto-dev curl wget nodejs npm \
    python3 python3-pip python3-venv
```
*This includes the C compiler, video tools, and Python (needed for motion detection).*

---

## 2. Download and Build (Required)
We will now turn the source code into a working program.

### Step A: Download the code
```bash
git clone https://github.com/opensensor/oneberry.git
cd oneberry
```

### Step B: Build the Web Interface (Required)
This creates the dashboard you will see in your browser. You must run this to have a working UI.
```bash
./scripts/build_web_vite.sh
```

### Step C: Build the Core Engine (Required)
This creates the main "Oneberry" program. You must run this to create the executable engine.
```bash
./scripts/build.sh --release
```

---

## 3. Installation and Services (Required)
Now we install the program into your system so it starts automatically.

### Step A: System Install (Required)
```bash
sudo ./scripts/install.sh
```

### Step B: Setup WebRTC (Required for Live Video)
This tool allows you to see live video with zero delay in your browser.
```bash
sudo ./scripts/install_go2rtc.sh
```

### Step C: Start the Service (Required)
```bash
sudo systemctl enable oneberry
sudo systemctl start oneberry
```

---

## 4. Setting up Motion Detection
Oneberry uses a smart AI engine to detect movement.

### Step A: Install the AI Brain
```bash
# This installs the detection software using the Python tools from Step 1.
pip install light-object-detect

# Start the AI engine (it will run on port 9001 and wait for Oneberry).
light-object-detect --host 0.0.0.0 --port 9001
```

### Step B: Configure via Web Browser
1. Open your browser and go to: `http://localhost:8080`
2. **Login**: Username: `admin` | Password: `admin`
3. Go to **Streams** -> **Configure** on your camera.
4. Enable **Detection Based Recording**.
5. Set **API Detection URL** to: `http://localhost:9001/api/v1/detect`

---

## 5. Quick Troubleshooting
- **Is it running?**: `sudo systemctl status oneberry`
- **See what's happening**: `tail -f /var/log/oneberry/oneberry.log`
- **Something not working?**: Check [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)

*Enjoy your new NVR system!*
