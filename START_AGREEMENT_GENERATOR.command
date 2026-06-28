#!/bin/bash

# Astonish Listing Agreement Generator - Start Script
# This script starts the local development server and opens it in your browser

PROJECT_DIR="/Users/michaelbergman/Library/CloudStorage/GoogleDrive-mbergman@astonishcommercial.com/Shared drives/MB:GC:KS/++ Apps/Listing Agreement Sync"
PORT=3737
URL="http://localhost:$PORT"

# Change to project directory
cd "$PROJECT_DIR" || exit 1

# Start the server
npm start &
SERVER_PID=$!

# Wait a few seconds for the server to start
sleep 3

# Open in browser
open "$URL"

# Keep the script running so the server continues
wait $SERVER_PID
