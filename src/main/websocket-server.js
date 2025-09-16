const WebSocket = require('ws');
const { v4: uuidv4 } = require('uuid');

class WebSocketServer {
  constructor(port = 8765, presentationManager, authToken = null) {
    this.port = port;
    this.presentationManager = presentationManager;
    this.wss = null;
    this.clients = new Map(); // clientId -> { ws, authenticated, lastPing }
    this.authToken = authToken || this.generateAuthToken();
    
    console.log(`Use auth token: ${this.authToken}`);
  }

  generateAuthToken() {
    return uuidv4().substring(0, 8); // Simple 8-character token (fallback only)
  }

  async start() {
    this.wss = new WebSocket.Server({ 
      port: this.port,
      perMessageDeflate: false
    });

    this.wss.on('connection', (ws, request) => {
      const clientId = uuidv4();
      
      this.clients.set(clientId, {
        ws,
        authenticated: false,
        lastPing: Date.now(),
        ip: request.socket.remoteAddress
      });

      console.log(`Client connected: ${clientId} from ${request.socket.remoteAddress}`);

      ws.on('message', async (data) => {
        try {
          const message = JSON.parse(data.toString());
          await this.handleMessage(clientId, message);
        } catch (error) {
          console.error('Error parsing message:', error);
          this.sendError(clientId, 'Invalid JSON format');
        }
      });

      ws.on('close', () => {
        console.log(`Client disconnected: ${clientId}`);
        this.clients.delete(clientId);
      });

      ws.on('error', (error) => {
        console.error(`WebSocket error for client ${clientId}:`, error);
        this.clients.delete(clientId);
      });

      // Send initial handshake
      this.send(clientId, {
        type: 'handshake',
        message: 'Presentation Control Helper v1.0',
        authRequired: true
      });
    });

    // Setup ping interval
    this.setupPingInterval();

    console.log(`WebSocket server listening on port ${this.port}`);
    console.log(`Use auth token: ${this.authToken}`);
  }

  async handleMessage(clientId, message) {
    const client = this.clients.get(clientId);
    if (!client) return;

    console.log(`Received from ${clientId}:`, message);

    // Handle authentication
    if (message.type === 'auth') {
      if (message.token === this.authToken) {
        client.authenticated = true;
        this.send(clientId, {
          type: 'authResult',
          success: true,
          message: 'Authentication successful'
        });
        
        // Send current status
        await this.sendStatus(clientId);
      } else {
        this.send(clientId, {
          type: 'authResult',
          success: false,
          message: 'Invalid authentication token'
        });
      }
      return;
    }

    // Require authentication for all other commands
    if (!client.authenticated) {
      this.sendError(clientId, 'Authentication required');
      return;
    }

    // Handle presentation commands
    try {
      await this.handlePresentationCommand(clientId, message);
    } catch (error) {
      console.error('Error handling presentation command:', error);
      this.sendError(clientId, `Command failed: ${error.message}`);
    }
  }

  async handlePresentationCommand(clientId, message) {
    const { type, ...params } = message;

    switch (type) {
      case 'ping':
        this.send(clientId, { type: 'pong', timestamp: Date.now() });
        break;

      case 'status':
        await this.sendStatus(clientId);
        break;

      case 'listFiles':
        const files = await this.presentationManager.listFiles();
        this.send(clientId, {
          type: 'fileList',
          files: files
        });
        break;

      case 'openFile':
        if (!params.filePath) {
          throw new Error('filePath parameter required');
        }
        const result = await this.presentationManager.openFile(params.filePath);
        this.send(clientId, {
          type: 'fileOpened',
          success: result.success,
          message: result.message,
          filePath: params.filePath
        });
        
        // Broadcast status update to all clients
        await this.broadcastStatus();
        break;

      case 'start':
        await this.presentationManager.startPresentation();
        this.send(clientId, { type: 'commandResult', command: 'start', success: true });
        await this.broadcastStatus();
        break;

      case 'stop':
        await this.presentationManager.stopPresentation();
        this.send(clientId, { type: 'commandResult', command: 'stop', success: true });
        await this.broadcastStatus();
        break;

      case 'close':
        await this.presentationManager.closePresentation();
        this.send(clientId, { type: 'commandResult', command: 'close', success: true });
        await this.broadcastStatus();
        break;

      case 'next':
        await this.presentationManager.nextSlide();
        this.send(clientId, { type: 'commandResult', command: 'next', success: true });
        await this.broadcastStatus();
        break;

      case 'prev':
        await this.presentationManager.prevSlide();
        this.send(clientId, { type: 'commandResult', command: 'prev', success: true });
        await this.broadcastStatus();
        break;

      // goto command removed - too complex for reliable implementation

      // Media player commands removed - will be implemented via Office.js add-in

      default:
        throw new Error(`Unknown command: ${type}`);
    }
  }

  async sendStatus(clientId) {
    const status = await this.presentationManager.getStatus();
    this.send(clientId, {
      type: 'status',
      ...status
    });
  }

  async broadcastStatus() {
    const status = await this.presentationManager.getStatus();
    this.broadcast({
      type: 'status',
      ...status
    });
  }

  send(clientId, message) {
    const client = this.clients.get(clientId);
    if (client && client.ws.readyState === WebSocket.OPEN) {
      client.ws.send(JSON.stringify(message));
    }
  }

  sendError(clientId, error) {
    this.send(clientId, {
      type: 'error',
      message: error
    });
  }

  broadcast(message) {
    const messageStr = JSON.stringify(message);
    this.clients.forEach((client, clientId) => {
      if (client.authenticated && client.ws.readyState === WebSocket.OPEN) {
        client.ws.send(messageStr);
      }
    });
  }

  setupPingInterval() {
    setInterval(() => {
      const now = Date.now();
      this.clients.forEach((client, clientId) => {
        if (client.ws.readyState === WebSocket.OPEN) {
          // Send ping every 30 seconds
          if (now - client.lastPing > 30000) {
            client.ws.ping();
            client.lastPing = now;
          }
        } else {
          this.clients.delete(clientId);
        }
      });
    }, 10000); // Check every 10 seconds
  }

  getConnectionCount() {
    return this.clients.size;
  }

  disconnectAllClients() {
    console.log('Disconnecting all clients due to token regeneration');
    this.clients.forEach((client, clientId) => {
      if (client.ws.readyState === WebSocket.OPEN) {
        client.ws.send(JSON.stringify({
          type: 'tokenChanged',
          message: 'Authentication token has been regenerated. Please reconnect with the new token.'
        }));
        client.ws.close();
      }
    });
    this.clients.clear();
  }

  stop() {
    if (this.wss) {
      this.wss.close();
      this.clients.clear();
      console.log('WebSocket server stopped');
    }
  }
}

module.exports = WebSocketServer;
