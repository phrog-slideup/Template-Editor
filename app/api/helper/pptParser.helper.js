const JSZip = require("jszip");

class PPTXParser {
  constructor(fileBuffer, options = {}) {
    this.fileBuffer = fileBuffer;
    this.zip = new JSZip();
    this.files = {};
    this.options = {
      logFiles: options.logFiles || false, // Optional logging
    };
  }

  // Load the ZIP file from buffer
  async load() {
    try {
      await this.zip.loadAsync(this.fileBuffer);
    } catch (error) {
      console.error(`Error loading PPTX file: ${error.message}`);
      throw error;
    }
  }

  // Unzip the PPTX file into memory
  async unzip() {
    try {
      await this.load();
      Object.keys(this.zip.files).forEach((filename) => {
        if (this.options.logFiles) {
          console.log(`File in archive: ${filename}`); // Log every file extracted
        }
        this.files[filename] = this.zip.files[filename];
      });

      return this.files;
    } catch (error) {
      console.error(`Error unzipping PPTX file: ${error.message}`);
      throw error;
    }
  }

  // Get a file's content as text
  async getFileContent(filename) {
    try {
      const file = this.files[filename];
      if (!file) {
        throw new Error(`File not found in archive: ${filename}`);
      }
      return await file.async("string");
    } catch (error) {
      console.error(`Error reading file ${filename}: ${error.message}`);
      throw error;
    }
  }
}

module.exports = PPTXParser;
