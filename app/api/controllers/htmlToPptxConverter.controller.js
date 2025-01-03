const path = require("path");
const fs = require("fs");

const htmlToPptx = require("../../template-engine/htmlToPptx");


/**
 * Convert HTML to PPTX and send response
 */
async function convertToPPTX(req, res) {
  try {
    const updatedHTML = req.body.html;
    if (!updatedHTML) {
      return res.status(400).json({ error: "No HTML content provided." });
    }

    if (!updatedHTML.includes("class=\"sli-slide\"")) {
      return res.status(400).json({ error: "Invalid HTML content: No slides found." });
    }

    const outputFilePath = path.resolve(__dirname, "../../files/presentation.pptx");
    await htmlToPptx.convertHTMLToPPTX(updatedHTML, outputFilePath);

    res.download(outputFilePath, "presentation.pptx", (err) => {
      if (err) {
        console.error("Error sending PPTX file:", err);
        res.status(500).json({ error: "Error generating the PPTX file." });
      }
    });
  } catch (error) {
    console.error("Error converting HTML to PPTX:", error);
    res.status(500).json({ error: "Error during conversion." });
  }
}



module.exports = {
  convertToPPTX,
};
