const path = require("path");
const fs = require("fs");
const pptxToHtml = require("../../template-engine/pptxToHtml");
const pptxParser = require("../helper/pptParser.helper");

let slidesHTMLCache = null; // Cache for slides HTML
const imagesInMemory = {}; // Store images in memory

// Handle file upload
async function uploadFile(req, res) {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded." });
    }

    const fileBuffer = req.file.buffer; // Read file buffer from memory

    // Pass buffer to the parser
    const parser = new pptxParser(fileBuffer);
    const unzippedFiles = await parser.unzip();

    const extractor = new pptxToHtml(unzippedFiles);

    // Override `getImageFromPicture` to store images in memory and generate URLs
    extractor.getImageFromPicture = async function (picNode, slidePath) {
      const blip = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0];
      const imageId = blip?.["$"]?.["r:embed"];
      if (!imageId) return null;

      const imagePath = await this.resolveImagePath(slidePath, imageId);
      if (imagePath) {
        const imageFile = this.files[imagePath];
        if (!imageFile) return null;

        const imageBuffer = await imageFile.async("nodebuffer");
        const imageKey = `image_${Date.now()}_${path.basename(imagePath)}`;
        imagesInMemory[imageKey] = imageBuffer; // Store in memory
        return `/api/slides/images/${imageKey}`; // Return dynamic URL
      }
      return null;
    };

    // Convert slides to HTML
    slidesHTMLCache = await extractor.convertAllSlidesToHTML();

    res.status(200).json({
      message: "File uploaded and processed successfully.",
      slides: slidesHTMLCache,
    });
  } catch (error) {
    console.error("Error uploading file:", error);
    res.status(500).json({ error: "Error uploading file." });
  }
}

// Serve images dynamically from memory
async function serveImage(req, res) {
  try {
    const imageKey = req.params.key;
    const imageBuffer = imagesInMemory[imageKey];

    if (imageBuffer) {
      const extension = path.extname(imageKey).toLowerCase().replace(".", "");
      const mimeType = extension === "jpg" ? "jpeg" : extension; // Handle JPG MIME type
      res.type(`image/${mimeType}`);
      res.send(imageBuffer);
    } else {
      res.status(404).json({ error: "Image not found." });
    }
  } catch (error) {
    console.error("Error serving image:", error);
    res.status(500).json({ error: "Error serving image." });
  }
}

// Fetch slides as HTML
async function getSlides(req, res) {
  try {
    if (!slidesHTMLCache) {
      return res.status(400).json({ error: "No slides available. Please upload a PPTX file first." });
    }

    res.json(slidesHTMLCache);
  } catch (error) {
    console.error("Error fetching slides:", error);
    res.status(500).json({ error: "Error fetching slides." });
  }
}

// Save slides (optional)
async function saveSlides(req, res) {
  try {
    const updatedSlides = req.body.updatedSlides;
    if (!updatedSlides) {
      return res.status(400).json({ error: "No slides provided." });
    }

    updatedSlides.forEach((slide, index) => {
      console.log(`Slide ${index + 1} processed.`);
    });

    res.json({ status: "Slides saved successfully." });
  } catch (error) {
    console.error("Error saving slides:", error);
    res.status(500).json({ error: "Error saving slides." });
  }
}

// Serve HTML file for download
async function downloadHTML(req, res) {
  try {
    if (!slidesHTMLCache) {
      return res.status(400).json({ error: "No slides available for download. Please upload and process a PPTX file first." });
    }

    // Combine cached slides into a single HTML file
    const fullHTML = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Converted Slides</title>
        <link rel="stylesheet" href="../../public/css/patternFill.css">
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
        </style>
      </head>
      <body>
        ${slidesHTMLCache.join("\n")}
      </body>
      </html>
    `;

    // Define the output path for the generated HTML file
    const htmlFilePath = path.resolve(__dirname, "../../public/converted-slides.html");

    // Write the HTML content to a file
    fs.writeFileSync(htmlFilePath, fullHTML, "utf8");

    // Send the file for download
    res.download(htmlFilePath, "converted-slides.html", (err) => {
      if (err) {
        console.error("Error sending file for download:", err);
        res.status(500).send("Error generating the HTML file.");
      }
    });
  } catch (error) {
    console.error("Error generating HTML file:", error);
    res.status(500).json({ error: "Error generating HTML file" });
  }
}

module.exports = {
  uploadFile,
  getSlides,
  downloadHTML,
  saveSlides,
  serveImage,
};
