const PptxGenJS = require("pptxgenjs");
const jsdom = require("jsdom");
const { JSDOM } = jsdom;
const path = require("path");
const fs = require("fs");
const puppeteer = require("puppeteer");
const JSZip = require("jszip");
const axios = require("axios");

let browserInstance = null; // Shared Puppeteer browser instance

/**
 * Core Function: HTML to PPTX Conversion
 */
async function convertHTMLToPPTX(htmlString, outputFilePath) {
    const dom = new JSDOM(htmlString);
    const document = dom.window.document;

    logWithTimestamp("Fetching computed styles using Puppeteer...");
    const computedStyles = await getComputedStyles(htmlString);
    logWithTimestamp("Computed styles fetched.");

    const pptx = new PptxGenJS();

    const slides = document.querySelectorAll(".sli-slide:not(.sli-slide .sli-slide)");
    if (slides.length === 0) {
        throw new Error("No slides found in the provided HTML.");
    }

    for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
        const slideElement = slides[slideIndex];
        const pptSlide = pptx.addSlide();

        // Set slide background
        const slideStyle = slideElement.style;
        let bgColor = slideStyle.backgroundColor || "#FFFFFF";
        const computedStyle = computedStyles[slideIndex];

        if (computedStyle.background && !bgColor.includes("gradient")) {
            bgColor = computedStyle.backgroundColor || computedStyle.background || "#FFFFFF";
        }

        logWithTimestamp(`Processing Slide ${slideIndex + 1} with background color: ${bgColor}`);

        if (bgColor.includes("linear-gradient")) {
            const gradientDetails = parseLinearGradient(bgColor);
            if (gradientDetails) {
                pptSlide.background = {
                    type: "gradient",
                    stops: gradientDetails.colors.map((color, index) => ({
                        color: rgbToHex(color),
                        position: Math.round((index / (gradientDetails.colors.length - 1)) * 100000),
                    })),
                    direction: gradientDetails.directions,
                };
            } else {
                pptSlide.background = { color: "#FFFFFF" };
            }
        } else {
            pptSlide.background = { color: rgbToHex(bgColor) };
        }

        // Collect all slide elements and sort by z-index
        const elements = [
            ...slideElement.querySelectorAll(".sli-txt-box"),
            ...slideElement.querySelectorAll(".sli-svg-container"),
            ...slideElement.querySelectorAll("img"),
        ].map((el) => ({ element: el, zIndex: parseFloat(el.style.zIndex) || 0 }));

        elements.sort((a, b) => a.zIndex - b.zIndex);

        // Add elements to the slide
        for (const { element } of elements) {

            if (element.classList.contains("sli-txt-box")) {

                addTextBoxToSlide(pptSlide, element);
            } else if (element.classList.contains("sli-svg-container")) {
                const svgElement = element.querySelector("svg");
                if (svgElement) {

                    addSvgToSlide(pptSlide, svgElement, element.style);
                }
            } else if (element.tagName === "IMG") {
                await addImageToSlide(pptSlide, element); // Handle asynchronously
            }
        }
    }

    logWithTimestamp("Writing PPTX file...");
    await pptx.writeFile({ fileName: outputFilePath });
    logWithTimestamp("PPTX file created successfully at:", outputFilePath);
}



/**
 * Add Text Box to Slide
 */
function addTextBoxToSlide(pptSlide, textBox) {
    const style = textBox.style;

    const x = normalizeStyleValue(style.left) / 96;
    const y = normalizeStyleValue(style.top) / 96;
    const w = normalizeStyleValue(style.width, 100) / 96;
    const h = normalizeStyleValue(style.height, 30) / 96;

    const defaultFontSizePx = parseFloat(style.fontSize || "14");
    const defaultFontSize = (defaultFontSizePx / 96) * 72;
    const defaultColor = rgbToHex(style.color || "#000000");
    const defaultFontFamily = style.fontFamily || "Arial";
    const defaultTextAlign = style.textAlign || "left";

    const paragraphs = Array.from(textBox.querySelectorAll("p"));
    const textSegments = [];

    paragraphs.forEach((paragraph) => {
        const spans = Array.from(paragraph.querySelectorAll("span"));
        if (spans.length === 0) {
            textSegments.push({
                text: paragraph.textContent.trim(),
                options: {
                    fontFace: defaultFontFamily,
                    fontSize: defaultFontSize,
                    color: defaultColor,
                },
            });
        } else {
            spans.forEach((span) => {
                const spanStyle = span.style;
                textSegments.push({
                    text: span.textContent.trim(),
                    options: {
                        fontFace: spanStyle.fontFamily || defaultFontFamily,
                        fontSize: normalizeStyleValue(spanStyle.fontSize || "14") / 96 * 72,
                        color: rgbToHex(spanStyle.color || defaultColor),
                    },
                });
            });
        }
    });

    if (textSegments.length > 0) {
        pptSlide.addText(textSegments, { x, y, w, h, align: defaultTextAlign });
    }
}

/**
 * Add SVG to Slide
 */
function addSvgToSlide(pptSlide, svgElement, style) {
    const x = normalizeStyleValue(style.left) / 96;
    const y = normalizeStyleValue(style.top) / 96;
    const w = normalizeStyleValue(style.width, 100) / 96;
    const h = normalizeStyleValue(style.height, 100) / 96;

    const svgBuffer = Buffer.from(svgElement.outerHTML, "utf-8");
    pptSlide.addImage({
        data: `data:image/svg+xml;base64,${svgBuffer.toString("base64")}`,
        x,
        y,
        w,
        h,
    });
}

/**
 * Add Image to Slide
 */
async function addImageToSlide(pptSlide, imgElement) {
    const src = imgElement.getAttribute("src");
    const style = imgElement.style;

    console.log("imgElement --------->", imgElement);
    console.log("src --------->", src);

    // Extract position and dimensions
    const x = normalizeStyleValue(style.left) / 96; // Convert px to inches
    const y = normalizeStyleValue(style.top) / 96;
    const w = normalizeStyleValue(style.width, 100) / 96;
    const h = normalizeStyleValue(style.height, 100) / 96;

    try {
        if (src.startsWith("/api/slides/images/")) {
            // Resolve the full URL for the image
            const fullUrl = `http://localhost:4000${src}`;

            console.log(`Fetching image from URL: ${fullUrl}`);

            // Fetch the image data
            const response = await axios.get(fullUrl, { responseType: "arraybuffer" });

            console.log(`Image fetched successfully from ${fullUrl}`);

            const contentType = response.headers["content-type"];
            const base64Data = `data:${contentType};base64,${Buffer.from(response.data).toString("base64")}`;

            // Add the image to the slide
            pptSlide.addImage({
                data: base64Data,
                x,
                y,
                w,
                h,
            });
        } else if (src.startsWith("data:image/")) {
            // If already Base64-encoded, directly add it
            pptSlide.addImage({
                data: src,
                x,
                y,
                w,
                h,
            });
        } else {
            console.error(`Invalid image source: ${src}`);
        }
    } catch (error) {
        console.error(`Failed to fetch image from ${src}:`, error.message);
    }
}


/**
 * Utility: Normalize style values
 */
function normalizeStyleValue(value, defaultValue = 0) {
    return parseFloat(value || `${defaultValue}`.replace("px", "")) || defaultValue;
}


/**
 * Fetch computed styles using Puppeteer
 */

async function getComputedStyles(htmlString) {
    if (!browserInstance) {
        browserInstance = await puppeteer.launch();
    }

    const page = await browserInstance.newPage();
    await page.setContent(htmlString);

    const styles = await page.evaluate(() =>
        [...document.querySelectorAll(".sli-slide")].map((slide) => {
            const computedStyle = window.getComputedStyle(slide);
            return {
                background: computedStyle.background || computedStyle.backgroundImage,
                backgroundColor: computedStyle.backgroundColor,
            };
        })
    );

    await page.close();
    return styles;
}

/**
 * Utility: Convert RGB or RGBA color to Hex
 */
function rgbToHex(rgb) {
    if (!rgb || typeof rgb !== "string") return "#000000"; // Default to black

    if (rgb.includes("rgba") && rgb.includes("0, 0, 0, 0")) {
        return "#FFFFFF"; // Default to white for transparent
    }

    const result = rgb.match(/\d+/g);
    if (result && result.length >= 3) {
        const r = parseInt(result[0]).toString(16).padStart(2, "0");
        const g = parseInt(result[1]).toString(16).padStart(2, "0");
        const b = parseInt(result[2]).toString(16).padStart(2, "0");
        return `#${r}${g}${b}`;
    }
    return "#000000"; // Default to black for invalid format
}


/**
 * Parse Linear Gradient Backgrounds
 */
function parseLinearGradient(gradient) {
    const match = gradient.match(/linear-gradient\(([^)]+)\)/);
    if (!match) return null;

    const content = match[1].split(",");
    const directions = content[0].trim();
    const colors = content.slice(1).map((color) => color.trim());
    return { directions, colors };
}

/**
 * Log with timestamp
 */
function logWithTimestamp(message) {
    console.log(`[${new Date().toISOString()}] ${message}`);
}

/**
 * Cleanup Puppeteer browser on process exit
 */
process.on("exit", async () => {
    if (browserInstance) {
        await browserInstance.close();
    }
});



module.exports = {
    convertHTMLToPPTX,
}