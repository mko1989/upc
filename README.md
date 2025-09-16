# 🎯 Presentation Control System

A macOS helper application that integrates with Bitfocus Companion / Stream Deck to control Keynote and PowerPoint presentations remotely.

## 🚀 Quick Start

### 1. Installation

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

## 📡 WebSocket API

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
{ "type": "goto", "slideNumber": 5 }

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

## 🧪 Testing

Use the included test client to verify functionality:

```bash
node test-client.js
```

This will:
1. Connect to the WebSocket server
2. Attempt authentication (will show you where to find the actual token)
3. Run basic API tests

## 🎮 Supported Applications

### Keynote
- ✅ Full AppleScript support
- ✅ Start/stop slideshow
- ✅ Navigate slides (next/prev/goto)
- ✅ Get slide count and current position
- ✅ Extract presenter notes
- ✅ Media timing (basic support)

### PowerPoint
- ⚠️ Limited AppleScript support
- ✅ Start/stop slideshow
- ✅ Navigate slides (keyboard shortcuts)
- ⚠️ Slide positioning (may be unreliable)
- ❌ Notes extraction (requires Office.js add-in)
- ❌ Advanced features (planned for Phase 3)

## 🏗️ Architecture

```
Bitfocus Companion (WebSocket client)
        ⇅
Electron Helper App (WebSocket server)
        ⇅
 ┌───────────────┬────────────────┐
 │ Keynote Driver│ PowerPoint Driver
 │ (AppleScript) │ (AppleScript + Add-in)
 └───────────────┴────────────────┘
```

### Components

- **Main Process** (`src/main/main.js`) - Electron app entry point and tray UI
- **WebSocket Server** (`src/main/websocket-server.js`) - Handles Companion communication
- **Presentation Manager** (`src/managers/presentation-manager.js`) - Routes commands to appropriate drivers
- **File Manager** (`src/managers/file-manager.js`) - Handles presentation file discovery and management
- **Keynote Driver** (`src/drivers/keynote-driver.js`) - AppleScript automation for Keynote
- **PowerPoint Driver** (`src/drivers/powerpoint-driver.js`) - AppleScript automation for PowerPoint

## 🔧 Development

### Project Structure
```
src/
├── main/           # Main Electron process
├── managers/       # Business logic managers
├── drivers/        # Application-specific drivers
├── renderer/       # UI components
└── utils/          # Shared utilities

assets/             # Icons and static resources
scripts/            # Build and deployment scripts
powerpoint-addin/   # Office.js add-in (Phase 3)
```

### Build Commands
```bash
npm run dev         # Development mode with debugging
npm run build       # Build for production
npm run pack        # Create app package
npm run dist        # Create distributable
```

## 📋 Roadmap

### ✅ Phase 1: MVP (Current)
- Electron helper with WebSocket server
- Folder selection and file listing
- Keynote AppleScript driver
- Basic Companion integration

### 🚧 Phase 2: PowerPoint Basic
- PowerPoint AppleScript driver improvements
- Enhanced slide count and navigation
- Better error handling

### 📅 Phase 3: PowerPoint Advanced
- Office.js Add-in with shared runtime
- Rich slide information (notes, thumbnails, titles)
- Advanced PowerPoint features

### 📅 Phase 4: Media & Polish
- Media timing and playback control
- Enhanced presenter notes
- Tray UI improvements
- Authentication flow

## 🛠️ Troubleshooting

### Common Issues

**"Permission denied" errors**
- Grant Accessibility permissions to Terminal.app and the Electron app in System Preferences → Security & Privacy → Privacy → Accessibility

**Keynote/PowerPoint not responding**
- Ensure the applications are installed and can be launched normally
- Try manually opening a presentation first
- Check that no system dialogs are blocking the applications

**WebSocket connection fails**
- Verify the helper app is running (check for tray icon)
- Ensure port 8765 is not blocked by firewall
- Check console output for error messages

**Authentication fails**
- Use the exact token shown in the console output
- Tokens are regenerated each time the app starts

### Debug Mode

Run in development mode for detailed logging:
```bash
npm run dev
```

Check the logs window in the tray menu for detailed error information.

## 📝 License

MIT License - See LICENSE file for details.

## 🤝 Contributing

1. Follow the existing code structure and patterns
2. Test with both Keynote and PowerPoint
3. Update documentation for new features
4. Ensure AppleScript commands are robust and handle errors gracefully

## 📞 Support

For issues and feature requests, please check the troubleshooting section first, then create an issue with:
- macOS version
- Application versions (Keynote/PowerPoint)
- Steps to reproduce
- Error messages from console/logs
