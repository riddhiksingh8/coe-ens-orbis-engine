#!/bin/bash

# Download and extract Node.js v20.x (adjust version as needed)
curl -fsSL https://nodejs.org/dist/latest-v20.x/node-v20.x.x-linux-x64.tar.xz | tar -xJ 

# Set Node.js binary path
export PATH="$PWD/node-v20.x.x-linux-x64/bin:$PATH"

# Install required dependencies (without sudo)
apt-get update
apt-get install -y unoconv libreoffice libreoffice-common

# Start LibreOffice in headless mode (required for unoconv)
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" --nofirststartwizard &

# Install dependencies
npm install
