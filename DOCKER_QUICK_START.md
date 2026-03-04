# Oneberry Docker Quick Start Guide

Get Oneberry running in under 5 minutes!

## Prerequisites

- Docker installed ([Get Docker](https://docs.docker.com/get-docker/))
- Docker Compose installed (included with Docker Desktop)

## Quick Start

### Option 1: Docker Compose (Recommended)

```bash
# Clone the repository
git clone https://github.com/opensensor/oneberry.git
cd oneberry

# Start Oneberry
docker-compose up -d

# View logs
docker-compose logs -f
```

**Access the Web UI:** http://localhost:8080

**Default Credentials:**
- Username: `admin`
- Password: `admin`

### Option 2: Docker Run

```bash
# Create directories for persistent data
mkdir -p config data

# Run Oneberry
docker run -d \
  --name oneberry \
  --restart unless-stopped \
  -p 8080:8080 \
  -p 8554:8554 \
  -p 8555:8555 \
  -p 8555:8555/udp \
  -p 1984:1984 \
  -v $(pwd)/config:/etc/oneberry \
  -v $(pwd)/data:/var/lib/oneberry/data \
  -e TZ=America/New_York \
  ghcr.io/opensensor/oneberry:latest

# View logs
docker logs -f oneberry
```

## What Happens on First Run?

The container automatically:

1. ✅ Creates default configuration files
2. ✅ Initializes the database
3. ✅ Sets up web assets
4. ✅ Configures go2rtc with WebRTC support
5. ✅ Creates admin user with default credentials

**No manual configuration needed!**

## Port Reference

| Port | Service | Description |
|------|---------|-------------|
| 8080 | Web UI | Main web interface |
| 8554 | RTSP | RTSP streaming server |
| 8555 | WebRTC | WebRTC streaming (TCP/UDP) |
| 1984 | go2rtc API | go2rtc REST API |

## Common Commands

### View Logs
```bash
# Docker Compose
docker-compose logs -f

# Docker Run
docker logs -f oneberry
```

### Restart Container
```bash
# Docker Compose
docker-compose restart

# Docker Run
docker restart oneberry
```

### Stop Container
```bash
# Docker Compose
docker-compose down

# Docker Run
docker stop oneberry
```

### Update to Latest Version
```bash
# Docker Compose
docker-compose pull
docker-compose up -d

# Docker Run
docker pull ghcr.io/opensensor/oneberry:latest
docker stop oneberry
docker rm oneberry
# Then run the docker run command again
```

## Adding Your First Camera

1. **Access Web UI:** http://localhost:8080
2. **Login** with admin/admin
3. **Go to Streams** page
4. **Click "Add Stream"**
5. **Enter camera details:**
   - Name: `Front Door`
   - RTSP URL: `rtsp://username:password@camera-ip:554/stream`
   - Enable recording if desired
6. **Click "Save"**

## Viewing Streams

### Web UI
- Navigate to **Streams** page
- Click on stream name to view live feed
- WebRTC streaming works automatically!

### RTSP Client
```bash
# Using ffplay
ffplay rtsp://localhost:8554/stream-name

# Using VLC
vlc rtsp://localhost:8554/stream-name
```

## Customizing Configuration

### Edit Main Configuration
```bash
# Edit oneberry.ini
nano ./config/oneberry.ini

# Restart to apply changes
docker-compose restart
```

### Edit go2rtc Configuration
```bash
# Edit go2rtc.yaml
nano ./config/go2rtc/go2rtc.yaml

# Restart to apply changes
docker-compose restart
```

## Troubleshooting

### Web UI Not Loading

**Check if container is running:**
```bash
docker ps | grep oneberry
```

**Check logs for errors:**
```bash
docker-compose logs -f
```

**Verify port is accessible:**
```bash
curl http://localhost:8080
```

### WebRTC Not Working

**Ensure UDP port is exposed:**
```bash
docker ps | grep oneberry
# Should show 8555/tcp and 8555/udp
```

**Check go2rtc is running:**
```bash
docker exec oneberry ps aux | grep go2rtc
```

**Test go2rtc API:**
```bash
curl http://localhost:1984/api/streams
```

### Database Lost After Restart

**Verify volume mounts:**
```bash
docker inspect oneberry | grep -A 10 Mounts
```

**Should show:**
- `/etc/oneberry`
- `/var/lib/oneberry/data`

**⚠️ Important:** Do NOT mount `/var/lib/oneberry` directly!

### Recordings Not Saving

**Check disk space:**
```bash
df -h
```

**Check recordings directory:**
```bash
ls -la ./data/recordings/
```

**Check configuration:**
```bash
grep -A 5 "\[storage\]" ./config/oneberry.ini
```

## Environment Variables

Customize behavior with environment variables:

```yaml
environment:
  # Set your timezone
  - TZ=America/New_York
  
  # Persist go2rtc config (default: true)
  - GO2RTC_CONFIG_PERSIST=true
  
  # Auto-initialize configs (default: true)
  - LIGHTNVR_AUTO_INIT=true
```

## Volume Structure

```
./config/                    # Configuration files
├── oneberry.ini            # Main configuration
└── go2rtc/
    └── go2rtc.yaml         # go2rtc configuration

./data/                      # Persistent data
├── database/
│   └── oneberry.db         # SQLite database
├── recordings/
│   ├── hls/                # HLS recordings
│   └── mp4/                # MP4 recordings
└── models/                 # Object detection models
```

## Security Best Practices

1. **Change Default Password:**
   - Login to web UI
   - Go to Users page
   - Change admin password

2. **Use Strong Passwords:**
   - For web UI
   - For camera RTSP URLs

3. **Firewall Configuration:**
   - Only expose necessary ports
   - Use reverse proxy for HTTPS

4. **Regular Updates:**
   ```bash
   docker-compose pull
   docker-compose up -d
   ```

## Next Steps

- 📖 Read the [Full Docker Guide](docs/DOCKER.md)
- 🔧 Check [Configuration Guide](docs/CONFIGURATION.md)
- 🐛 See [Troubleshooting Guide](docs/TROUBLESHOOTING.md)
- 💬 Join discussions on [GitHub](https://github.com/opensensor/oneberry/discussions)

## Getting Help

- **Documentation:** https://github.com/opensensor/oneberry/tree/main/docs
- **Issues:** https://github.com/opensensor/oneberry/issues
- **Discussions:** https://github.com/opensensor/oneberry/discussions

## Example docker-compose.yml

```yaml
version: '3.8'

services:
  oneberry:
    image: ghcr.io/opensensor/oneberry:latest
    container_name: oneberry
    restart: unless-stopped
    ports:
      - "8080:8080"     # Web UI
      - "8554:8554"     # RTSP
      - "8555:8555"     # WebRTC TCP
      - "8555:8555/udp" # WebRTC UDP
      - "1984:1984"     # go2rtc API
    volumes:
      - ./config:/etc/oneberry
      - ./data:/var/lib/oneberry/data
    environment:
      - TZ=America/New_York
      - GO2RTC_CONFIG_PERSIST=true
      - LIGHTNVR_AUTO_INIT=true
    networks:
      - oneberry

networks:
  oneberry:
    driver: bridge
```

---

**That's it! You now have a fully functional NVR system with WebRTC streaming. 🎉**

