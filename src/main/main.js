const { app, Tray, Menu, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const WebSocketServer = require('./websocket-server');
const PresentationManager = require('../managers/presentation-manager');
const FileManager = require('../managers/file-manager');
const TokenManager = require('../utils/token-manager');
const SettingsManager = require('../utils/settings-manager');

class PresentationControlApp {
  constructor() {
    this.tray = null;
    this.wsServer = null;
    this.presentationManager = null;
    this.fileManager = null;
    this.tokenManager = null;
    this.settingsManager = null;
    this.selectedFolder = null;
    this.isDevMode = process.argv.includes('--dev');
  }

  async initialize() {
    // Ensure app is ready
    await app.whenReady();
    
    // Initialize token manager and get persistent token
    this.tokenManager = new TokenManager();
    this.settingsManager = new SettingsManager();
    const authToken = this.tokenManager.getOrCreateToken();
    
    // Initialize managers
    this.fileManager = new FileManager();
    this.presentationManager = new PresentationManager(this.fileManager);
    
    // Setup WebSocket server with persistent token
    this.wsServer = new WebSocketServer(8765, this.presentationManager, authToken);
    await this.wsServer.start();
    
    // Create tray
    this.createTray();
    
    // Setup app event handlers
    this.setupAppHandlers();
    
    // Load last used folder if available
    await this.loadLastFolder();
    
    console.log('Presentation Control System initialized');
  }

  createTray() {
    // Create tray icon with fallback
    const iconPath = path.join(__dirname, '../../assets/tray-icon.png');
    
    try {
      if (fs.existsSync(iconPath)) {
        this.tray = new Tray(iconPath);
      } else {
        // Create a simple tray icon using nativeImage
        const { nativeImage } = require('electron');
        const icon = nativeImage.createEmpty();
        // Create a 16x16 template icon (macOS will handle the styling)
        this.tray = new Tray(icon);
        this.tray.setTitle('🎯'); // Use emoji as fallback visual indicator
      }
    } catch (error) {
      console.error('Error creating tray icon:', error);
      // Last resort: try to create tray without icon
      const { nativeImage } = require('electron');
      const icon = nativeImage.createEmpty();
      this.tray = new Tray(icon);
      this.tray.setTitle('PCS'); // Text fallback
    }

    // Set up initial menu and tooltip
    this.updateTrayMenu();
    this.tray.setToolTip('Presentation Control Helper');
  }

  async selectFolder() {
    // Get a valid default path or omit it
    const lastFolder = this.selectedFolder || this.settingsManager.getLastFolder();
    const dialogOptions = {
      properties: ['openDirectory'],
      title: 'Select Presentation Folder',
      // Ensure dialog appears in front
      modal: true,
      alwaysOnTop: true
    };
    
    // Only add defaultPath if we have a valid string path
    if (lastFolder && typeof lastFolder === 'string') {
      dialogOptions.defaultPath = lastFolder;
    }
    
    const result = await dialog.showOpenDialog(dialogOptions);

    if (!result.canceled && result.filePaths.length > 0) {
      this.selectedFolder = result.filePaths[0];
      
      // Save the folder preference
      this.settingsManager.setLastFolder(this.selectedFolder);
      
      // Set folder in file manager
      await this.fileManager.setFolder(this.selectedFolder);
      
      // Update tray tooltip
      this.tray.setToolTip(`Presentation Control Helper - ${path.basename(this.selectedFolder)}`);
      
      console.log(`📁 Folder selected: ${this.selectedFolder}`);
      
      // Update tray menu to reflect new folder
      this.updateTrayMenu();
      
      // Notify WebSocket clients about folder change
      this.wsServer.broadcast({
        type: 'folderChanged',
        folder: this.selectedFolder,
        files: this.fileManager.getFiles()
      });
    }
  }

  async loadLastFolder() {
    const lastFolder = this.settingsManager.getLastFolder();
    
    if (lastFolder && fs.existsSync(lastFolder)) {
      this.selectedFolder = lastFolder;
      await this.fileManager.setFolder(this.selectedFolder);
      
      // Update tray tooltip
      this.tray.setToolTip(`Presentation Control Helper - ${path.basename(this.selectedFolder)}`);
      
      console.log(`📁 Restored last folder: ${this.selectedFolder}`);
      console.log(`📊 Found ${this.fileManager.getFiles().length} presentation files`);
      
      // Update tray menu to reflect folder state
      this.updateTrayMenu();
    } else if (lastFolder) {
      console.log(`⚠️  Last folder no longer exists: ${lastFolder}`);
      // Clear the invalid folder setting
      this.settingsManager.setLastFolder(null);
    }
  }

  clearFolder() {
    this.selectedFolder = null;
    this.settingsManager.setLastFolder(null);
    this.fileManager.clearFolder();
    
    // Reset tray tooltip
    this.tray.setToolTip('Presentation Control Helper');
    
    console.log('📁 Folder cleared');
    
    // Update tray menu to reflect cleared state
    this.updateTrayMenu();
    
    // Notify WebSocket clients about folder change
    this.wsServer.broadcast({
      type: 'folderChanged',
      folder: null,
      files: []
    });
  }

  updateTrayMenu() {
    const contextMenu = Menu.buildFromTemplate([
      {
        label: 'Select Folder',
        click: () => this.selectFolder()
      },
      {
        label: 'Clear Folder',
        click: () => this.clearFolder(),
        enabled: !!this.selectedFolder
      },
      {
        label: 'Connection Status',
        click: () => this.showConnectionStatus()
      },
      {
        label: 'Auth Token',
        submenu: [
          {
            label: 'Show Current Token',
            click: () => this.showAuthToken()
          },
          {
            label: 'Copy Token to Clipboard',
            click: () => this.copyTokenToClipboard()
          },
          { type: 'separator' },
          {
            label: 'Regenerate Token',
            click: () => this.regenerateAuthToken()
          }
        ]
      },
      { type: 'separator' },
      {
        label: 'Show Logs',
        click: () => this.showLogs()
      },
      { type: 'separator' },
      {
        label: 'Quit',
        click: () => app.quit()
      }
    ]);

    this.tray.setContextMenu(contextMenu);
  }

  showConnectionStatus() {
    const connections = this.wsServer.getConnectionCount();
    const status = connections > 0 ? 'Connected' : 'No connections';
    const tokenInfo = this.tokenManager.getTokenInfo();
    
    dialog.showMessageBox({
      type: 'info',
      title: 'Connection Status',
      message: `WebSocket Server: Running on port 8765\nConnections: ${connections}\nStatus: ${status}\nFolder: ${this.selectedFolder || 'Not selected'}\n\nAuth Token: ${tokenInfo.token}\nGenerated: ${tokenInfo.generatedAt ? new Date(tokenInfo.generatedAt).toLocaleString() : 'Unknown'}`
    });
  }

  showAuthToken() {
    const tokenInfo = this.tokenManager.getTokenInfo();
    
    dialog.showMessageBox({
      type: 'info',
      title: 'Authentication Token',
      message: `Current Auth Token: ${tokenInfo.token}\n\nGenerated: ${tokenInfo.generatedAt ? new Date(tokenInfo.generatedAt).toLocaleString() : 'Unknown'}\n\nThis token persists between app restarts.\nUse this token in your Companion module configuration.`,
      buttons: ['OK', 'Copy to Clipboard'],
      defaultId: 0
    }).then((result) => {
      if (result.response === 1) {
        this.copyTokenToClipboard();
      }
    });
  }

  copyTokenToClipboard() {
    const { clipboard } = require('electron');
    const token = this.tokenManager.getCurrentToken();
    
    clipboard.writeText(token);
    
    dialog.showMessageBox({
      type: 'info',
      title: 'Token Copied',
      message: `Auth token "${token}" has been copied to clipboard!`,
      buttons: ['OK'],
      defaultId: 0
    });
  }

  regenerateAuthToken() {
    dialog.showMessageBox({
      type: 'warning',
      title: 'Regenerate Auth Token',
      message: 'This will generate a new authentication token.\n\nAll existing Companion connections will need to be updated with the new token.\n\nAre you sure you want to continue?',
      buttons: ['Cancel', 'Regenerate'],
      defaultId: 0,
      cancelId: 0
    }).then((result) => {
      if (result.response === 1) {
        const newToken = this.tokenManager.regenerateToken();
        
        // Update the WebSocket server with new token
        this.wsServer.authToken = newToken;
        
        // Disconnect all current clients to force re-authentication
        this.wsServer.disconnectAllClients();
        
        dialog.showMessageBox({
          type: 'info',
          title: 'Token Regenerated',
          message: `New auth token: ${newToken}\n\nAll existing connections have been disconnected.\nUpdate your Companion module configuration with the new token.`,
          buttons: ['OK', 'Copy to Clipboard'],
          defaultId: 0
        }).then((copyResult) => {
          if (copyResult.response === 1) {
            this.copyTokenToClipboard();
          }
        });
      }
    });
  }

  showLogs() {
    // Create a simple log window
    const logWindow = new BrowserWindow({
      width: 800,
      height: 600,
      title: 'Presentation Control Logs',
      webPreferences: {
        nodeIntegration: true,
        contextIsolation: false
      }
    });

    logWindow.loadFile(path.join(__dirname, '../renderer/logs.html'));
  }

  setupAppHandlers() {
    app.on('window-all-closed', (event) => {
      // Prevent app from quitting when all windows are closed (tray app)
      event.preventDefault();
    });

    app.on('before-quit', () => {
      if (this.wsServer) {
        this.wsServer.stop();
      }
    });

    // Handle dock icon hiding on macOS
    if (process.platform === 'darwin') {
      app.dock.hide();
    }
  }
}

// Initialize the app
const presentationApp = new PresentationControlApp();

app.whenReady().then(() => {
  presentationApp.initialize().catch(console.error);
});

app.on('activate', () => {
  // On macOS, re-create app if needed
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});
