#!/bin/bash

# Download and extract Node.js v20.x (adjust version as needed)
curl -fsSL https://nodejs.org/dist/latest-v20.x/node-v20.x.x-linux-x64.tar.xz | tar -xJ 

# Set Node.js binary path
export PATH="$PWD/node-v20.x.x-linux-x64/bin:$PATH"

# Install dependencies
npm install
