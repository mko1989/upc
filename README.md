# Ultimate Presentation Control

A macOS helper application that integrates with Bitfocus Companion / Stream Deck to control Keynote and PowerPoint presentations remotely.

## Quick Start

### 1. Installation

- grab a package from releases or build from source

```bash
# Clone the repository
git clone <repository-url>
cd UPC

# Install dependencies
npm install

# Start the application
npm start
```

### 2. Setup

1. **Launch the app** - A tray icon will appear in your menu bar
2. **Select folder** - Right-click the tray icon and choose "Select Folder" to pick a directory containing your presentation files (.key, .pptx, .ppt)
3. **Note the auth token** - Check the console output for the authentication token (8-character code)

### 3. Connect from Companion

The WebSocket server runs on `localhost:8765` and requires authentication using the generated token.

## WebSocket API

### Authentication

First, authenticate with the server:

```json
{
  "type": "auth",
  "token": "your-8-char-token"
}
```

### Available Commands

#### File Management
```json
// List all presentation files in the selected folder
{ "type": "listFiles" }

// Open a specific presentation file
{ "type": "openFile", "filePath": "/path/to/presentation.key" }
```

#### Presentation Control
```json
// Start slideshow
{ "type": "start" }

// Stop slideshow
{ "type": "stop" }

// Navigate slides
{ "type": "next" }
{ "type": "prev" }

// Get current status
{ "type": "status" }

// Test connection
{ "type": "ping" }
```

### Response Format

All responses include a `type` field. Common response types:

- `authResult` - Authentication response
- `status` - Current presentation status
- `fileList` - Available presentation files
- `commandResult` - Result of a command execution
- `error` - Error message

## Supported Applications

### Keynote
- Full AppleScript support
- Start/stop slideshow
- Navigate slides (next/prev/goto)
- Get slide count and current position
- Extract presenter notes


### PowerPoint
- AppleScript support
- Start/stop slideshow
- Navigate slides
- Get slide count and current position
- Notes extraction


### Components

- **Main Process** (`src/main/main.js`) - Electron app entry point and tray UI
- **WebSocket Server** (`src/main/websocket-server.js`) - Handles Companion communication
- **Presentation Manager** (`src/managers/presentation-manager.js`) - Routes commands to appropriate drivers
- **File Manager** (`src/managers/file-manager.js`) - Handles presentation file discovery and management
- **Keynote Driver** (`src/drivers/keynote-driver.js`) - AppleScript automation for Keynote
- **PowerPoint Driver** (`src/drivers/powerpoint-driver.js`) - AppleScript automation for PowerPoint


## License

MIT License - See LICENSE file for details.