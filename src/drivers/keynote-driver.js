const { execSync } = require('child_process');
const path = require('path');

class KeynoteDriver {
  constructor() {
    this.appName = 'Keynote';
    this.currentDocument = null;
  }

  async openFile(filePath) {
    try {
      const script = `
        tell application "Keynote"
          activate
          open POSIX file "${filePath}"
          set currentDoc to the front document
          return name of currentDoc
        end tell
      `;

      const result = this.executeAppleScript(script);
      this.currentDocument = result.trim();
      
      return { 
        success: true, 
        message: `Opened ${this.currentDocument} in Keynote`,
        document: this.currentDocument
      };
    } catch (error) {
      console.error('Keynote open file error:', error);
      return { 
        success: false, 
        message: `Failed to open file: ${error.message}` 
      };
    }
  }

  async startPresentation() {
    try {
      const script = `
        tell application "Keynote"
          tell the front document
            start slideshow
          end tell
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true, message: 'Slideshow started' };
    } catch (error) {
      console.error('Keynote start presentation error:', error);
      throw new Error(`Failed to start presentation: ${error.message}`);
    }
  }

  async stopPresentation() {
    try {
      const script = `
        tell application "Keynote"
          tell the front document
            stop slideshow
          end tell
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true, message: 'Slideshow stopped' };
    } catch (error) {
      console.error('Keynote stop presentation error:', error);
      throw new Error(`Failed to stop presentation: ${error.message}`);
    }
  }

  async closePresentation() {
    try {
      const script = `
        tell application "Keynote"
          if (count of documents) > 0 then
            tell the front document
              try
                stop slideshow
              end try
              close
            end tell
          end if
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true, message: 'Presentation closed' };
    } catch (error) {
      console.error('Keynote close presentation error:', error);
      throw new Error(`Failed to close presentation: ${error.message}`);
    }
  }

  async nextSlide() {
    try {
      // Use native AppleScript commands only
      const script = `
        tell application "Keynote"
          activate
          if (count of documents) > 0 then
            tell the front document
              show next
            end tell
          end if
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true };
    } catch (error) {
      console.error('Keynote next slide error:', error);
      throw new Error(`Failed to go to next slide: ${error.message}`);
    }
  }

  async prevSlide() {
    try {
      // Use native AppleScript commands only
      const script = `
        tell application "Keynote"
          activate
          if (count of documents) > 0 then
            tell the front document
              show previous
            end tell
          end if
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true };
    } catch (error) {
      console.error('Keynote previous slide error:', error);
      throw new Error(`Failed to go to previous slide: ${error.message}`);
    }
  }

  // gotoSlide removed - focus on core navigation functionality

  async getCurrentSlideInfo() {
    try {
      // First get total slides
      const countScript = `
        tell application "Keynote"
          if (count of documents) > 0 then
            tell the front document
              return count of slides
            end tell
          else
            return 0
          end if
        end tell
      `;
      
      const totalSlides = parseInt(this.executeAppleScript(countScript).trim()) || 0;
      
      // Then try to get current slide - use a simpler approach
      let currentSlide = 1;
      try {
        const currentScript = `
          tell application "Keynote"
            tell the front document
              try
                return slide number of current slide
              on error
                return 1
              end try
            end tell
          end tell
        `;
        currentSlide = parseInt(this.executeAppleScript(currentScript).trim()) || 1;
      } catch (slideError) {
        console.log('Could not get current slide, using 1 as default');
        currentSlide = 1;
      }
      
      return {
        currentSlide: currentSlide,
        totalSlides: totalSlides
      };
    } catch (error) {
      console.error('Keynote get slide info error:', error);
      // Return fallback values instead of throwing
      return {
        currentSlide: 1,
        totalSlides: 0
      };
    }
  }

  async getSlideNotes(slideNumber) {
    try {
      const script = `
        tell application "Keynote"
          tell the front document
            try
              set targetSlide to slide ${slideNumber}
              set slideNotes to presenter notes of targetSlide
              if slideNotes is missing value then
                return ""
              else
                return slideNotes
              end if
            on error
              return ""
            end try
          end tell
        end tell
      `;

      const result = this.executeAppleScript(script);
      return result.trim();
    } catch (error) {
      console.error('Keynote get slide notes error:', error);
      // Return empty string if notes are not available
      return '';
    }
  }

  async getTotalSlides() {
    try {
      const script = `
        tell application "Keynote"
          tell the front document
            return count of slides
          end tell
        end tell
      `;

      const result = this.executeAppleScript(script);
      return parseInt(result.trim());
    } catch (error) {
      console.error('Keynote get total slides error:', error);
      throw new Error(`Failed to get total slides: ${error.message}`);
    }
  }

  async isKeynoteRunning() {
    try {
      const script = `
        tell application "System Events"
          set keynoteRunning to (name of processes) contains "Keynote"
          return keynoteRunning
        end tell
      `;

      const result = this.executeAppleScript(script);
      return result.trim() === 'true';
    } catch (error) {
      return false;
    }
  }

  async getKeynoteStatus() {
    try {
      const script = `
        tell application "Keynote"
          if (count of documents) > 0 then
            tell the front document
              set isPlaying to playing
              set currentSlideNumber to slideNumber of current slide
              set totalSlideCount to count of slides
              set docName to name
              return docName & ":" & isPlaying & ":" & currentSlideNumber & ":" & totalSlideCount
            end tell
          else
            return "No document open"
          end if
        end tell
      `;

      const result = this.executeAppleScript(script);
      
      if (result.trim() === "No document open") {
        return {
          hasDocument: false,
          isPlaying: false,
          currentSlide: 0,
          totalSlides: 0,
          documentName: null
        };
      }

      const [docName, isPlaying, currentSlide, totalSlides] = result.trim().split(':');
      
      return {
        hasDocument: true,
        isPlaying: isPlaying === 'true',
        currentSlide: parseInt(currentSlide),
        totalSlides: parseInt(totalSlides),
        documentName: docName
      };
    } catch (error) {
      console.error('Keynote get status error:', error);
      return {
        hasDocument: false,
        isPlaying: false,
        currentSlide: 0,
        totalSlides: 0,
        documentName: null
      };
    }
  }

  executeAppleScript(script) {
    try {
      const result = execSync(`osascript -e '${script.replace(/'/g, "'\\''")}' 2>&1`, {
        encoding: 'utf8',
        timeout: 10000 // 10 second timeout
      });
      
      return result;
    } catch (error) {
      // Parse AppleScript error messages
      let errorMessage = error.message;
      if (error.stdout) {
        errorMessage = error.stdout;
      }
      
      // Clean up common AppleScript error patterns
      errorMessage = errorMessage.replace(/^[0-9]+:[0-9]+:\s*/, ''); // Remove line numbers
      errorMessage = errorMessage.replace(/execution error:\s*/i, ''); // Remove "execution error:"
      
      throw new Error(errorMessage.trim());
    }
  }
}

module.exports = KeynoteDriver;
