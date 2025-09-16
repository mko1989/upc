const { app } = require('electron');
const path = require('path');
const fs = require('fs');

class SettingsManager {
  constructor() {
    this.settingsFilePath = path.join(app.getPath('userData'), 'settings.json');
    this.settings = this.loadSettings();
  }

  loadSettings() {
    if (fs.existsSync(this.settingsFilePath)) {
      try {
        const data = fs.readFileSync(this.settingsFilePath, 'utf8');
        return JSON.parse(data);
      } catch (error) {
        console.error('Error loading settings:', error);
        return {};
      }
    }
    return {};
  }

  saveSettings() {
    try {
      fs.writeFileSync(this.settingsFilePath, JSON.stringify(this.settings, null, 2), 'utf8');
      console.log('Settings saved');
    } catch (error) {
      console.error('Error saving settings:', error);
    }
  }

  get(key, defaultValue = null) {
    return this.settings[key] !== undefined ? this.settings[key] : defaultValue;
  }

  set(key, value) {
    this.settings[key] = value;
    this.saveSettings();
  }

  // Specific getters/setters for common settings
  getLastFolder() {
    return this.get('lastFolder');
  }

  setLastFolder(folderPath) {
    this.set('lastFolder', folderPath);
    console.log(`📁 Saved folder preference: ${folderPath}`);
  }

  getWindowPosition() {
    return this.get('windowPosition');
  }

  setWindowPosition(position) {
    this.set('windowPosition', position);
  }

  getAllSettings() {
    return { ...this.settings };
  }
}

module.exports = SettingsManager;
