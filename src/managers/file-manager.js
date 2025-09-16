const fs = require('fs').promises;
const path = require('path');
const chokidar = require('chokidar');

class FileManager {
  constructor() {
    this.currentFolder = null;
    this.files = [];
    this.currentFileIndex = -1;
    this.watcher = null;
    this.supportedExtensions = ['.key', '.pptx', '.ppt'];
  }

  async setFolder(folderPath) {
    // Stop watching previous folder
    if (this.watcher) {
      await this.watcher.close();
    }

    this.currentFolder = folderPath;
    this.currentFileIndex = -1;
    
    // Scan for presentation files
    await this.scanFiles();
    
    // Watch for file changes
    this.setupWatcher();
    
    console.log(`File manager set to folder: ${folderPath}`);
    console.log(`Found ${this.files.length} presentation files`);
  }

  async scanFiles() {
    if (!this.currentFolder) {
      this.files = [];
      return;
    }

    try {
      const entries = await fs.readdir(this.currentFolder, { withFileTypes: true });
      const presentationFiles = [];

      for (const entry of entries) {
        if (entry.isFile()) {
          const ext = path.extname(entry.name).toLowerCase();
          if (this.supportedExtensions.includes(ext)) {
            const fullPath = path.join(this.currentFolder, entry.name);
            const stats = await fs.stat(fullPath);
            
            presentationFiles.push({
              name: entry.name,
              path: fullPath,
              extension: ext,
              size: stats.size,
              modified: stats.mtime,
              type: this.getFileType(ext)
            });
          }
        }
      }

      // Sort by name for consistent ordering
      presentationFiles.sort((a, b) => a.name.localeCompare(b.name));
      
      this.files = presentationFiles;
    } catch (error) {
      console.error('Error scanning files:', error);
      this.files = [];
    }
  }

  getFileType(extension) {
    switch (extension) {
      case '.key':
        return 'keynote';
      case '.pptx':
      case '.ppt':
        return 'powerpoint';
      default:
        return 'unknown';
    }
  }

  setupWatcher() {
    if (!this.currentFolder) return;

    this.watcher = chokidar.watch(this.currentFolder, {
      ignored: /(^|[\/\\])\../, // ignore dotfiles
      persistent: true,
      depth: 0 // only watch the specified folder, not subdirectories
    });

    this.watcher.on('add', (filePath) => {
      const ext = path.extname(filePath).toLowerCase();
      if (this.supportedExtensions.includes(ext)) {
        console.log(`Presentation file added: ${filePath}`);
        this.scanFiles(); // Rescan to update list
      }
    });

    this.watcher.on('unlink', (filePath) => {
      const ext = path.extname(filePath).toLowerCase();
      if (this.supportedExtensions.includes(ext)) {
        console.log(`Presentation file removed: ${filePath}`);
        this.scanFiles(); // Rescan to update list
        
        // Reset current file index if the current file was deleted
        if (this.currentFileIndex >= 0 && 
            this.files[this.currentFileIndex] && 
            this.files[this.currentFileIndex].path === filePath) {
          this.currentFileIndex = -1;
        }
      }
    });

    this.watcher.on('change', (filePath) => {
      const ext = path.extname(filePath).toLowerCase();
      if (this.supportedExtensions.includes(ext)) {
        console.log(`Presentation file changed: ${filePath}`);
        this.scanFiles(); // Rescan to update metadata
      }
    });
  }

  getFiles() {
    return this.files;
  }

  getCurrentFile() {
    if (this.currentFileIndex >= 0 && this.currentFileIndex < this.files.length) {
      return this.files[this.currentFileIndex];
    }
    return null;
  }

  setCurrentFile(filePath) {
    const index = this.files.findIndex(file => file.path === filePath);
    if (index >= 0) {
      this.currentFileIndex = index;
      return this.files[index];
    }
    return null;
  }

  nextFile() {
    if (this.files.length === 0) return null;
    
    this.currentFileIndex = (this.currentFileIndex + 1) % this.files.length;
    return this.files[this.currentFileIndex];
  }

  prevFile() {
    if (this.files.length === 0) return null;
    
    this.currentFileIndex = this.currentFileIndex <= 0 
      ? this.files.length - 1 
      : this.currentFileIndex - 1;
    return this.files[this.currentFileIndex];
  }

  getFileByIndex(index) {
    if (index >= 0 && index < this.files.length) {
      this.currentFileIndex = index;
      return this.files[index];
    }
    return null;
  }

  getFileIndex(filePath) {
    return this.files.findIndex(file => file.path === filePath);
  }

  getCurrentFileIndex() {
    return this.currentFileIndex;
  }

  getFileCount() {
    return this.files.length;
  }

  async validateFile(filePath) {
    try {
      const stats = await fs.stat(filePath);
      return stats.isFile();
    } catch (error) {
      return false;
    }
  }

  formatFileInfo(file) {
    if (!file) return null;
    
    return {
      name: file.name,
      path: file.path,
      type: file.type,
      size: this.formatFileSize(file.size),
      modified: file.modified.toISOString()
    };
  }

  formatFileSize(bytes) {
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    if (bytes === 0) return '0 Bytes';
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
  }

  clearFolder() {
    // Clear current folder and files
    this.currentFolder = null;
    this.files = [];
    this.currentFileIndex = -1;
    
    // Stop watching
    if (this.watcher) {
      this.watcher.close();
      this.watcher = null;
    }
    
    console.log('File manager cleared');
  }

  destroy() {
    if (this.watcher) {
      this.watcher.close();
    }
  }
}

module.exports = FileManager;
