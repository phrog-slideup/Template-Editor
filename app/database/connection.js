const PptxGenJS = require("pptxgenjs");
const jsdom = require("jsdom");
const { JSDOM } = jsdom;
const path = require("path");
const fs = require("fs");
const puppeteer = require("puppeteer");

// Convert HTML back to PPTX
async function convertToPPTX(req, res) {
  try {
    const updatedHTML = req.body.html;
    if (!updatedHTML) {
      return res.status(400).json({ error: "No HTML content provided." });
    }

    console.log("updatedHTML ------>>> ", updatedHTML);

    const outputFilePath = path.resolve(__dirname, "../../files/presentation.pptx");
    await convertHTMLToPPTX(updatedHTML, outputFilePath);

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

// Convert HTML to PPTX
async function convertHTMLToPPTX(htmlString, outputFilePath) {
  const dom = new JSDOM(htmlString);
  const document = dom.window.document;

  // Fetch computed styles for slides using Puppeteer
  console.log("Fetching computed styles using Puppeteer...");
  const computedStyles = await getComputedStyles(htmlString);
  console.log("Computed styles fetched:", computedStyles);

  let pptx = new PptxGenJS();
  const slides = document.querySelectorAll(".slide:not(.slide .slide)");

  if (slides.length === 0) {
    throw new Error("No slides found in the provided HTML.");
  }

  for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
    const slide = slides[slideIndex];
    const pptSlide = pptx.addSlide();

    console.log(`Processing slide ${slideIndex + 1}`);

    // Get background style from inline styles or computed styles
    const inlineStyle = slide.style;
    const computedStyle = computedStyles[slideIndex] || {};

    let bgColor = inlineStyle.backgroundColor || computedStyle.background || "#FFFFFF";

    // Check for gradients in the background
    if (bgColor.includes("linear-gradient")) {
      console.log(`Slide ${slideIndex + 1} - Gradient background detected: ${bgColor}`);
      const gradientMatch = bgColor.match(/linear-gradient\((.*)\)/);
      if (gradientMatch) {
        const gradientColors = gradientMatch[1].split(",").map((color) => color.trim());
        console.log(`Parsed gradient colors for slide ${slideIndex + 1}:`, gradientColors);
        pptSlide.background = convertGradientToXML(gradientColors);
      }
    } else if (bgColor.includes("rgb") || bgColor.startsWith("#")) {
      console.log(`Slide ${slideIndex + 1} - Solid background detected: ${bgColor}`);
      pptSlide.background = { color: rgbToHex(bgColor) };
    } else {
      console.warn(`Slide ${slideIndex + 1} - Unknown background format: ${bgColor}`);
      pptSlide.background = { color: "#FFFFFF" }; // Default to white background
    }

    // Process text boxes
    const textBoxes = slide.querySelectorAll(".text-box");
    textBoxes.forEach((textBox, boxIndex) => {
      const style = textBox.style;
      const x = parseFloat(style.left || "0") / 96; // Convert px to inches
      const y = parseFloat(style.top || "0") / 96;
      const w = parseFloat(style.width || "100") / 96;
      const h = parseFloat(style.height || "30") / 96;
      const fontSizePx = parseFloat(style["font-size"] || "12");
      const fontSize = (fontSizePx / 96) * 72; // Convert px to pt
      const textAlign = style.textAlign || "left";
      const color = rgbToHex(style.color || "#000");

      console.log(`Text box ${boxIndex + 1} on slide ${slideIndex + 1}: x=${x}, y=${y}, w=${w}, h=${h}, fontSize=${fontSize}, color=${color}`);

      const spanElements = textBox.querySelectorAll("span");
      const textParts = Array.from(spanElements).map((span) => {
        const spanStyles = extractSpanStyles(span);
        return {
          text: spanStyles.text,
          options: {
            bold: spanStyles.bold,
            italic: spanStyles.italic,
            underline: spanStyles.underline,
            fontFace: spanStyles.fontFace,
            color: spanStyles.color,
            fontSize: spanStyles.fontSize,
          },
        };
      });

      pptSlide.addText(textParts, {
        x,
        y,
        w,
        h,
        align: textAlign,
        bullet: textBox.querySelector("ul") ? true : false,
      });

      console.log(`Processed text box ${boxIndex + 1} on slide ${slideIndex + 1}`);
    });

    // Process images
    const images = slide.querySelectorAll("img");
    images.forEach((img, imgIndex) => {
      const style = img.style;
      const x = parseFloat(style.left || "0") / 96;
      const y = parseFloat(style.top || "0") / 96;
      const w = parseFloat(style.width || "100") / 96;
      const h = parseFloat(style.height || "100") / 96;

      console.log(`Image ${imgIndex + 1} on slide ${slideIndex + 1}: x=${x}, y=${y}, w=${w}, h=${h}`);

      const resolvedPath = path.resolve(__dirname, "../../files", img.src);

      if (!fs.existsSync(resolvedPath)) {
        console.warn(`Image not found: ${resolvedPath}`);
        return;
      }

      pptSlide.addImage({
        path: resolvedPath,
        x,
        y,
        w,
        h,
      });

      console.log(`Processed image ${imgIndex + 1} on slide ${slideIndex + 1}`);
    });
  }

  await pptx.writeFile({ fileName: outputFilePath });
}

// Utility function to convert RGB color to Hex format
function rgbToHex(rgb) {
  if (!rgb) return "#FFFFFF"; // Default to white if no color is provided
  const result = rgb.match(/\d+/g);
  if (result && result.length >= 3) {
    const r = parseInt(result[0]).toString(16).padStart(2, "0");
    const g = parseInt(result[1]).toString(16).padStart(2, "0");
    const b = parseInt(result[2]).toString(16).padStart(2, "0");
    return `#${r}${g}${b}`;
  }
  return "#FFFFFF"; // Default to white for invalid color formats
}

// Function to convert gradient colors to XML-compatible format
function convertGradientToXML(colors) {
  if (!colors || colors.length < 2) {
    console.warn("Gradient conversion requires at least two colors.");
    return { type: "solid", color: "#FFFFFF" }; // Fallback to white
  }

  // Convert gradient stops into PowerPoint-compatible format
  const stops = colors.map((color, index) => {
    const hexColor = rgbToHex(color);
    const position = Math.round((index / (colors.length - 1)) * 100000); // PowerPoint uses 0-100000 scale
    return `<a:gs pos="${position}"><a:srgbClr val="${hexColor.substring(1)}" /></a:gs>`;
  });

  // Create gradient XML
  return {
    type: "gradient",
    colors: stops.join(""),
  };
}

// Extract styles from inline <span> elements
function extractSpanStyles(span) {
  const style = span.style;

  return {
    text: span.textContent || "",
    bold: style.fontWeight === "bold" || style.fontWeight === "700",
    italic: style.fontStyle === "italic",
    underline: style.textDecoration?.includes("underline"),
    fontFace: style.fontFamily?.replace(/['"]/g, "") || "Arial",
    color: rgbToHex(style.color || "#000"),
    fontSize: (parseFloat(style.fontSize || "12") / 96) * 72, // Convert px to pt
  };
}

// Optional: Use Puppeteer for computed styles
async function getComputedStyles(htmlString) {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.setContent(htmlString);
  const styles = await page.evaluate(() => {
    const slides = document.querySelectorAll(".slide");
    return Array.from(slides).map((slide) => {
      const computedStyle = window.getComputedStyle(slide);
      return {
        background: computedStyle.background || computedStyle.backgroundImage,
        width: computedStyle.width,
        height: computedStyle.height,
      };
    });
  });
  await browser.close();
  return styles;
}

module.exports = {
  convertToPPTX,
};
