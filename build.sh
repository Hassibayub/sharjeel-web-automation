#!/bin/bash
set -e

echo "Building Windows executable with Docker..."

# Create output directory
mkdir -p output

# Build the Docker image and run
docker build -f Dockerfile.build -t windows-builder .

# Run container to build and copy exe
docker run --rm -v "$(pwd)/output:/output" windows-builder

echo ""
echo "Build complete!"
echo "Output: ./output/SharjeelAutomation.exe"
