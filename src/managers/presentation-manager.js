const FileManager = require('./file-manager');
const KeynoteDriver = require('../drivers/keynote-driver');
const PowerPointDriver = require('../drivers/powerpoint-driver');

class PresentationManager {
  constructor(fileManager = null) {
    this.fileManager = fileManager || new FileManager();
    this.keynoteDriver = new KeynoteDriver();
    this.powerpointDriver = new PowerPointDriver();
    
    this.currentDriver = null;
    this.currentFile = null;
    this.isPresenting = false;
    this.currentSlideIndex = 0;
    this.totalSlides = 0;
    this.slideNotes = '';
  }

  async listFiles() {
    return this.fileManager.getFiles();
  }

  async openFile(filePath) {
    try {
      // Validate file exists
      const isValid = await this.fileManager.validateFile(filePath);
      if (!isValid) {
        return { success: false, message: 'File not found or invalid' };
      }

      // Set current file in file manager
      const file = this.fileManager.setCurrentFile(filePath);
      if (!file) {
        return { success: false, message: 'File not in current folder' };
      }

      // Determine which driver to use
      const driver = this.getDriverForFile(file);
      if (!driver) {
        return { success: false, message: 'Unsupported file type' };
      }

      // Open file with appropriate driver
      const result = await driver.openFile(filePath);
      if (result.success) {
        this.currentDriver = driver;
        this.currentFile = file;
        this.isPresenting = false;
        
        // Get initial presentation info
        await this.updatePresentationInfo();
        
        console.log(`Opened file: ${file.name} with ${file.type} driver`);
      }

      return result;
    } catch (error) {
      console.error('Error opening file:', error);
      return { success: false, message: error.message };
    }
  }

  getDriverForFile(file) {
    switch (file.type) {
      case 'keynote':
        return this.keynoteDriver;
      case 'powerpoint':
        return this.powerpointDriver;
      default:
        return null;
    }
  }

  async startPresentation() {
    if (!this.currentDriver || !this.currentFile) {
      throw new Error('No presentation file is currently open');
    }

    await this.currentDriver.startPresentation();
    this.isPresenting = true;
    await this.updatePresentationInfo();
    
    console.log('Presentation started');
  }

  async stopPresentation() {
    if (!this.currentDriver) {
      throw new Error('No presentation is currently active');
    }

    await this.currentDriver.stopPresentation();
    this.isPresenting = false;
    
    console.log('Presentation stopped');
  }

  async closePresentation() {
    if (!this.currentDriver) {
      throw new Error('No presentation is currently open');
    }

    await this.currentDriver.closePresentation();
    
    // Reset state
    this.currentDriver = null;
    this.currentFile = null;
    this.isPresenting = false;
    this.currentSlideIndex = 0;
    this.totalSlides = 0;
    
    console.log('Presentation closed');
  }

  async nextSlide() {
    if (!this.currentDriver) {
      throw new Error('No presentation is currently open');
    }

    await this.currentDriver.nextSlide();
    await this.updatePresentationInfo();
    
    console.log(`Advanced to slide ${this.currentSlideIndex + 1}`);
  }

  async prevSlide() {
    if (!this.currentDriver) {
      throw new Error('No presentation is currently open');
    }

    await this.currentDriver.prevSlide();
    await this.updatePresentationInfo();
    
    console.log(`Went back to slide ${this.currentSlideIndex + 1}`);
  }

  // gotoSlide removed - too complex for reliable AppleScript implementation

  async updatePresentationInfo() {
    if (!this.currentDriver) return;

    try {
      // Get current slide info
      const slideInfo = await this.currentDriver.getCurrentSlideInfo();
      this.currentSlideIndex = slideInfo.currentSlide - 1; // Convert to 0-based
      this.totalSlides = slideInfo.totalSlides;

      // Get slide notes if available
      try {
        this.slideNotes = await this.currentDriver.getSlideNotes(slideInfo.currentSlide);
      } catch (error) {
        this.slideNotes = ''; // Notes may not be available
      }
    } catch (error) {
      console.error('Error updating presentation info:', error);
    }
  }

  async getStatus() {
    const currentFile = this.currentFile;
    const files = this.fileManager.getFiles();
    const currentFileIndex = currentFile ? this.fileManager.getCurrentFileIndex() : -1;

    // Get real-time slide info if driver supports it (only for PowerPoint when presenting)
    let realTimeSlideInfo = null;
    if (this.currentDriver === this.powerpointDriver && this.isPresenting && typeof this.currentDriver.getCurrentSlideInfo === 'function') {
      try {
        realTimeSlideInfo = await this.currentDriver.getCurrentSlideInfo();
      } catch (error) {
        console.error('Error getting real-time slide info:', error);
        // Fallback to internal tracking
        realTimeSlideInfo = null;
      }
    }

    return {
      currentFile: currentFile ? {
        name: currentFile.name,
        path: currentFile.path,
        type: currentFile.type,
        index: currentFileIndex
      } : null,
      isPresenting: this.isPresenting,
      currentSlide: realTimeSlideInfo ? realTimeSlideInfo.currentSlide : (this.currentSlideIndex + 1),
      totalSlides: realTimeSlideInfo ? realTimeSlideInfo.totalSlides : this.totalSlides,
      slideNotes: this.slideNotes,
      folder: this.fileManager.currentFolder,
      fileCount: files.length,
      driverType: this.currentDriver ? 
        (this.currentDriver === this.keynoteDriver ? 'keynote' : 'powerpoint') : null
      // mediaPlayer removed - will be implemented via Office.js add-in
    };
  }

  async getSlideNotes(slideNumber) {
    if (!this.currentDriver) {
      throw new Error('No presentation is currently open');
    }

    return await this.currentDriver.getSlideNotes(slideNumber || this.currentSlideIndex + 1);
  }

  async getSlideList() {
    if (!this.currentDriver) {
      throw new Error('No presentation is currently open');
    }

    // This might only work with PowerPoint add-in
    if (typeof this.currentDriver.getSlideList === 'function') {
      return await this.currentDriver.getSlideList();
    }

    // Fallback: generate basic list
    const slides = [];
    for (let i = 1; i <= this.totalSlides; i++) {
      slides.push({
        index: i,
        title: `Slide ${i}`,
        notes: ''
      });
    }
    return slides;
  }

  // File navigation methods
  async openNextFile() {
    const nextFile = this.fileManager.nextFile();
    if (nextFile) {
      return await this.openFile(nextFile.path);
    }
    return { success: false, message: 'No next file available' };
  }

  async openPrevFile() {
    const prevFile = this.fileManager.prevFile();
    if (prevFile) {
      return await this.openFile(prevFile.path);
    }
    return { success: false, message: 'No previous file available' };
  }

  async openFileByIndex(index) {
    const file = this.fileManager.getFileByIndex(index);
    if (file) {
      return await this.openFile(file.path);
    }
    return { success: false, message: 'Invalid file index' };
  }

  // Folder management
  async setFolder(folderPath) {
    await this.fileManager.setFolder(folderPath);
    
    // Reset current state when folder changes
    this.currentDriver = null;
    this.currentFile = null;
    this.isPresenting = false;
    this.currentSlideIndex = 0;
    this.totalSlides = 0;
    this.slideNotes = '';
  }

  // Cleanup
  destroy() {
    this.fileManager.destroy();
  }

  // Media player functionality removed - will be implemented via Office.js add-in
}

module.exports = PresentationManager;
