#!/bin/bash
# Deploy new Oneberry binary and trigger recording sync

set -e

echo "=== Deploying new Oneberry binary ==="

# Stop service if running
echo "Stopping oneberry service..."
sudo systemctl stop oneberry || true

# Install new binary
echo "Installing new binary..."
sudo cp build/bin/oneberry /usr/local/bin/oneberry
sudo chmod +x /usr/local/bin/oneberry

# Start service
echo "Starting oneberry service..."
sudo systemctl start oneberry

# Wait for service to start
echo "Waiting for service to start..."
sleep 5

# Check service status
echo "Checking service status..."
sudo systemctl status oneberry --no-pager || true

# Wait a bit more for web server to be ready
sleep 2

# Trigger sync via API
echo ""
echo "=== Triggering recording sync via API ==="
curl -X POST http://admin:admin@localhost:8080/api/recordings/sync 2>&1 || {
    echo "Failed to call sync API, trying to check logs..."
    sudo journalctl -u oneberry --since "1 minute ago" --no-pager | tail -20
}

echo ""
echo "=== Done ==="

