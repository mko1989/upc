const { execSync } = require('child_process');
const path = require('path');

class PowerPointDriver {
  constructor() {
    this.appName = 'Microsoft PowerPoint';
    this.currentPresentation = null;
    this.addinConnected = false; // Will be used for Office.js add-in in future
  }

  async openFile(filePath) {
    try {
      const script = `
        tell application "Microsoft PowerPoint"
          activate
          open POSIX file "${filePath}"
          set currentPres to the active presentation
          return name of currentPres
        end tell
      `;

      const result = this.executeAppleScript(script);
      this.currentPresentation = result.trim();
      
      return { 
        success: true, 
        message: `Opened ${this.currentPresentation} in PowerPoint`,
        presentation: this.currentPresentation
      };
    } catch (error) {
      console.error('PowerPoint open file error:', error);
      return { 
        success: false, 
        message: `Failed to open file: ${error.message}` 
      };
    }
  }

  async startPresentation() {
    try {
      const script = `
                      tell application "Microsoft PowerPoint"
	                      activate
	                        tell slide show settings of active presentation
		                        set show with presenter to true
	                        end tell
	                      run slide show slide show settings of active presentation
                      end tell
      `;

      this.executeAppleScript(script);
      return { success: true, message: 'Slideshow started' };
    } catch (error) {
      console.error('PowerPoint start presentation error:', error);
      throw new Error(`Failed to start presentation: ${error.message}`);
    }
  }

  async stopPresentation() {
    try {
      const script = `
        tell application "Microsoft PowerPoint"
          activate
          tell slide show view of slide show window 1
            exit slide show
          end tell
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true, message: 'Slideshow stopped' };
    } catch (error) {
      console.error('PowerPoint stop presentation error:', error);
      throw new Error(`Failed to stop presentation: ${error.message}`);
    }
  }

  async closePresentation() {
    try {
      const script = `
        tell application "Microsoft PowerPoint"
          set thePresentation to presentation 1
          try
            exit slide show slide show view of thePresentation
          end try
          close thePresentation
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true, message: 'Presentation closed' };
    } catch (error) {
      console.error('PowerPoint close presentation error:', error);
      throw new Error(`Failed to close presentation: ${error.message}`);
    }
  }

  async nextSlide() {
    try {
      const script = `
                      tell application "Microsoft PowerPoint"
	                      activate
	                        tell slide show view of slide show window 1
		                        go to next slide
	                        end tell
                        end tell

      `;

      this.executeAppleScript(script);
      return { success: true };
    } catch (error) {
      console.error('PowerPoint next slide error:', error);
      throw new Error(`Failed to go to next slide: ${error.message}`);
    }
  }

  async prevSlide() {
    try {
      const script = `
        tell application "Microsoft PowerPoint"
          activate
          tell slide show view of slide show window 1
            go to previous slide
          end tell
        end tell
      `;

      this.executeAppleScript(script);
      return { success: true };
    } catch (error) {
      console.error('PowerPoint previous slide error:', error);
      throw new Error(`Failed to go to previous slide: ${error.message}`);
    }
  }

  // gotoSlide removed - too complex for reliable implementation via AppleScript

  async getCurrentSlideInfo() {
    try {
      // Get total slides using working AppleScript with error handling
      const totalScript = `
        tell application "Microsoft PowerPoint"
          try
            if (count of presentations) > 0 then
              tell active presentation
                set theSlideCount to count slides
              end tell
            else
              set theSlideCount to 0
            end if
          on error
            set theSlideCount to 0
          end try
        end tell
      `;

      const totalSlides = parseInt(this.executeAppleScript(totalScript).trim()) || 0;
      
      // Try to get current slide using working AppleScript
      let currentSlide = 1;
      try {
        const currentScript = `
          tell application "Microsoft PowerPoint"
            try
              set CurrentSlide to slide index of slide of slide show view of slide show window 1 of active presentation
            on error
              1
            end try
          end tell
        `;
        
        currentSlide = parseInt(this.executeAppleScript(currentScript).trim()) || 1;
      } catch (slideError) {
        // If we can't get current slide, default to 1 (not in slideshow mode)
        currentSlide = 1;
      }
      
      return {
        currentSlide: currentSlide,
        totalSlides: totalSlides
      };
    } catch (error) {
      console.error('PowerPoint get slide info error:', error);
      return {
        currentSlide: 1,
        totalSlides: 0
      };
    }
  }

  async getSlideNotes(slideNumber) {
    try {
      // Use your exact working AppleScript - one simple line
      const script = `
        tell application "Microsoft PowerPoint"
          try
            set CurrentSlide to notes text of presenter tool of presenter view window 1 of active presentation
          on error
            ""
          end try
        end tell
      `;

      const result = this.executeAppleScript(script);
      return result.trim();
    } catch (error) {
      console.error('PowerPoint get slide notes error:', error);
      return ''; // Return empty string on error
    }
  }

  async getTotalSlides() {
    try {
      const script = `
        tell application "Microsoft PowerPoint"
          tell the active presentation
            return count of slides
          end tell
        end tell
      `;

      const result = this.executeAppleScript(script);
      return parseInt(result.trim());
    } catch (error) {
      console.error('PowerPoint get total slides error:', error);
      throw new Error(`Failed to get total slides: ${error.message}`);
    }
  }

  async isPowerPointRunning() {
    try {
      const script = `
        tell application "System Events"
          set ppRunning to (name of processes) contains "Microsoft PowerPoint"
          return ppRunning
        end tell
      `;

      const result = this.executeAppleScript(script);
      return result.trim() === 'true';
    } catch (error) {
      return false;
    }
  }

  async getPowerPointStatus() {
    try {
      const script = `
        tell application "Microsoft PowerPoint"
          if (count of presentations) > 0 then
            tell the active presentation
              set presName to name
              set totalSlideCount to count of slides
              try
                set isInSlideShow to (slide show view exists)
                if isInSlideShow then
                  set currentSlideNumber to slide index of slide show view
                  return presName & ":true:" & currentSlideNumber & ":" & totalSlideCount
                else
                  return presName & ":false:1:" & totalSlideCount
                end if
              on error
                return presName & ":false:1:" & totalSlideCount
              end try
            end tell
          else
            return "No presentation open"
          end if
        end tell
      `;

      const result = this.executeAppleScript(script);
      
      if (result.trim() === "No presentation open") {
        return {
          hasPresentation: false,
          isPlaying: false,
          currentSlide: 0,
          totalSlides: 0,
          presentationName: null
        };
      }

      const [presName, isPlaying, currentSlide, totalSlides] = result.trim().split(':');
      
      return {
        hasPresentation: true,
        isPlaying: isPlaying === 'true',
        currentSlide: parseInt(currentSlide),
        totalSlides: parseInt(totalSlides),
        presentationName: presName
      };
    } catch (error) {
      console.error('PowerPoint get status error:', error);
      return {
        hasPresentation: false,
        isPlaying: false,
        currentSlide: 0,
        totalSlides: 0,
        presentationName: null
      };
    }
  }

  // Office.js Add-in integration methods (placeholder for future implementation)
  setAddinConnection(connected) {
    this.addinConnected = connected;
  }

  isAddinConnected() {
    return this.addinConnected;
  }

  // Enhanced methods that will use add-in when available
  async getSlideList() {
    if (this.addinConnected) {
      // TODO: Implement add-in communication
      throw new Error('Add-in communication not yet implemented');
    }
    
    // Fallback: generate basic list using AppleScript
    const totalSlides = await this.getTotalSlides();
    const slides = [];
    
    for (let i = 1; i <= totalSlides; i++) {
      slides.push({
        index: i,
        title: `Slide ${i}`,
        notes: '' // Will be enhanced by add-in
      });
    }
    
    return slides;
  }

  // Media player functionality removed - will be handled by Office.js add-in

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

module.exports = PowerPointDriver;
