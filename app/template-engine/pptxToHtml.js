const xml2js = require("xml2js");
const path = require("path");
const fs = require("fs");
const backgroundColor = require("../styling/pptBackgroundColors.js");
const pptTextAllInfo = require("../styling/pptTextAllInfo.js");

class pptxToHtml {
  constructor(unzippedFiles) {
    this.files = unzippedFiles;
    this.themeColors = null;
    this.cache = {};
  }

  // Clear the cache
  clearCache() {
    this.cache = {};
  }

  // Parse XML from the unzipped PPTX file with caching
  async parseXML(filePath) {
    if (this.cache[filePath]) {
      console.log(`Using cached version of ${filePath}`);
      return this.cache[filePath];
    }

    const file = this.files[filePath];
    if (!file) {
      console.error(`File not found: ${filePath}`);
      return null;
    }
    try {
      const xmlContent = await file.async("string");
      const parser = new xml2js.Parser();
      const result = await parser.parseStringPromise(xmlContent);

      this.cache[filePath] = result; // Cache the parsed XML
      return result;
    } catch (error) {
      console.error(`Error parsing XML file at ${filePath}:`, error);
      throw new Error(`Failed to parse XML: ${filePath}`);
    }
  }


  // Convert all slides in the presentation to HTML
  async convertAllSlidesToHTML() {
    const slideHTMLs = [];
    const slidePaths = Object.keys(this.files).filter((filename) => filename.startsWith("ppt/slides/slide"));

    for (const slidePath of slidePaths) {
      const slideHTML = await this.convertSlideToHTML(slidePath);
      if (slideHTML) slideHTMLs.push(slideHTML);
    }
    return slideHTMLs;
  }


  // Convert individual slides to HTML

  async convertSlideToHTML(slidePath) {
    const slideXML = await this.parseXML(slidePath);
    const masterPath = "ppt/slideMasters/slideMaster1.xml";
    const themePath = "ppt/theme/theme1.xml";

    let relPath;
    if (slidePath.endsWith(".rels")) {
      relPath = slidePath; // It's already a .rels file
    } else {
      const relDir = path.posix.dirname(slidePath);
      const relFile = `${path.posix.basename(slidePath)}.rels`;
      relPath = path.posix.join(relDir, "_rels", relFile);
    }

    const relationshipsXML = await this.parseXML(relPath);
    const masterXML = await this.parseXML(masterPath);
    const themeXML = await this.parseXML(themePath);

    // Get color mapping from the master slide
    const clrMap = await this.getColorMappingFromMaster(masterXML);

    let htmlContent = `<div class="sli-slide" 
                              style="position:relative; 
                              width:960px; height:720px; 
                              background-color:#FFFFFF; 
                              padding:20px;">`;

    // Get the background color or gradient
    const slideBg = await backgroundColor.getBackgroundColor(slideXML, masterXML, themeXML, relationshipsXML);

    if (slideBg) {
      let bgStyle = "";

      if (slideBg.startsWith("linear-gradient") || slideBg.startsWith("radial-gradient")) {
        bgStyle = `background: ${slideBg};`; // Gradient fill
      } else if (slideBg.startsWith("url(")) {
        bgStyle = `background-image: ${slideBg}; background-size: cover;`; // Picture/texture fill
      } else {
        bgStyle = `background-color: ${slideBg};`; // Solid color
      }

      htmlContent = `<div class="sli-slide" style="position:relative; width:960px; height:720px; ${bgStyle} padding:20px;">`;
    }

    // Process text shapes
    const shapeNodes = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
    for (const shapeNode of shapeNodes) {
      const shapeInfo = this.getAllTextInformationFromShape(shapeNode, themeXML, clrMap);
      const zIndexStyle = `z-index:${shapeInfo.zIndex};`;

      if (pptTextAllInfo.isTextShape(shapeNode)) {
        htmlContent += `<div class="sli-txt-box" contenteditable="true" 
                            style="position:absolute; 
                            left:${shapeInfo.position.x}px; 
                            top:${shapeInfo.position.y}px; 
                            width:${shapeInfo.position.width}px; 
                            height:${shapeInfo.position.height}px; 
                            color:${shapeInfo.fontColor}; padding:7px; 
                            overflow-wrap:break-word; 
                            box-sizing:border-box; 
                            transform:rotate(${shapeInfo.rotation}deg); 
                            font-size:${shapeInfo.fontSize}px; 
                            text-align:${shapeInfo.textAlign}; 
                            ${zIndexStyle}">
                              ${shapeInfo.text}
                        </div>`;
      }
    }

    // Process pictures and SVGs
    const picNodes = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:pic"] || [];
    for (const picNode of picNodes) {
      const position = pptTextAllInfo.getPositionFromShape(picNode);
      const blipNode = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0];

      const svgContent = await this.getSVGContent(blipNode, relationshipsXML);
      const imageSrc = svgContent ? null : await this.getImageFromPicture(picNode, slidePath, relationshipsXML);
      const rotation = pptTextAllInfo.getRotation(picNode);

      const alphaValue = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0]?.["a:alphaModFix"]?.[0]?.["$"]?.amt || "100000";
      const opacity = parseInt(alphaValue, 10) / 100000;

      const borderWidth = picNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0]?.["a:lnW"]?.[0] || "0";
      const borderColor = picNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0]?.["a:solidFill"]?.[0]?.["a:srgbClr"]?.[0]?.["$"]?.val || "#000000";

      const shadowColor = picNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0]?.["a:srgbClr"]?.[0]?.["$"]?.val || "#000000";
      const shadowOffsetX = picNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0]?.["$"]?.dx || "0";
      const shadowOffsetY = picNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0]?.["$"]?.dy || "0";

      const flipH = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.flipH === "1" ? "scaleX(-1)" : "";
      const flipV = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.flipV === "1" ? "scaleY(-1)" : "";

      if (svgContent) {
        htmlContent += `<div class="sli-svg-container" 
                            style="position:absolute; 
                            left:${position.x}px; 
                            top:${position.y}px; 
                            width:${position.width}px; 
                            height:${position.height}px; 
                            overflow:hidden;">
                            <svg style="width:100%; height:100%;" xmlns="http://www.w3.org/2000/svg">${svgContent}</svg>
                        </div>`;
      } else if (imageSrc) {
        htmlContent += `<img src="${imageSrc}" alt="Image" 
                          style="position:absolute; 
                          left:${position.x}px; 
                          top:${position.y}px; 
                          width:${position.width}px; 
                          height:${position.height}px; 
                          object-fit:contain; 
                          transform:${flipH} ${flipV} rotate(${rotation}deg); 
                          opacity:${opacity}; 
                          border:${parseInt(borderWidth, 10) / 12700}px solid ${borderColor}; 
                          box-shadow:${parseInt(shadowOffsetX, 10) / 12700}px ${parseInt(shadowOffsetY, 10) / 12700}px ${shadowColor};" />`;
      } else {
        console.warn("No valid SVG or PNG/JPG image found for the node.");
      }
    }

    htmlContent += "</div>";
    return htmlContent;
  }


  getAllTextInformationFromShape(shapeNode, themeXML, clrMap) {
    return {
      position: pptTextAllInfo.getPositionFromShape(shapeNode),
      text: pptTextAllInfo.getTextFromShape(shapeNode),
      textAlign: pptTextAllInfo.getTextAlignment(shapeNode),
      fontSize: pptTextAllInfo.getFontSize(shapeNode),
      fontColor: pptTextAllInfo.getFontColor(shapeNode, themeXML, clrMap),
      rotation: pptTextAllInfo.getRotation(shapeNode),
      zIndex: shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.id || 0, // Use ID as stacking order
    };
  }

  getColorMappingFromMaster(masterXML) {

    const clrMap = masterXML?.["p:sldMaster"]?.["p:clrMap"]?.[0]?.["$"];
    if (!clrMap) {
      console.error("No color mapping found in the master slide.");
      return null;
    }

    // Extract color mappings
    return {
      bg1: clrMap.bg1,
      tx1: clrMap.tx1,
      bg2: clrMap.bg2,
      tx2: clrMap.tx2,
      accent1: clrMap.accent1,
      accent2: clrMap.accent2,
      accent3: clrMap.accent3,
      accent4: clrMap.accent4,
      accent5: clrMap.accent5,
      accent6: clrMap.accent6,
      hlink: clrMap.hlink,
      folHlink: clrMap.folHlink,
    };
  }

  // Get SVG content from <a:blip> node

  async getSVGContent(blipNode, relationshipsXML) {

    const svgEmbedId = blipNode?.["a:extLst"]?.[0]?.["a:ext"]?.[0]?.["asvg:svgBlip"]?.[0]?.["$"]?.["r:embed"];

    if (!svgEmbedId) {
      return null;
    }

    const relationship = relationshipsXML?.["Relationships"]?.["Relationship"]?.find((rel) => rel["$"]?.Id === svgEmbedId);

    if (!relationship) {
      return null;
    }

    const svgPath = path.posix.normalize(`ppt/media/${relationship["$"]?.Target}`);

    const svgFile = this.files[svgPath];
    if (!svgFile) {
      return null;
    }

    return await svgFile.async("string");
  }


  // Get image path from a picture node


  async getImageFromPicture(picNode, slidePath, relationshipsXML) {

    const svgBlip = picNode?.["p:blipFill"]?.[0]?.["a:extLst"]?.[0]?.["a:ext"]?.find(ext => ext?.["asvg:svgBlip"]);
    const blip = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0];

    // console.log("svgBlip  --------@@@@@@@@@@@>>>>>>>>>> ", svgBlip);

    let imageId = null;
    if (svgBlip) {
      imageId = svgBlip["asvg:svgBlip"]?.[0]?.["$"]?.["r:embed"];
    } else if (blip) {
      imageId = blip["$"]?.["r:embed"];
    }

    console.log("Processing image with ID:", imageId);

    if (!imageId) return null;

    const rel = relationshipsXML?.["Relationships"]?.["Relationship"]?.find(rel => rel["$"]?.Id === imageId);

    // console.log("rel  --------@@@@@@@@@@@>>>>>>>>>> ", rel);

    if (!rel) {
      return null;
    }

    const imagePath = path.posix.normalize(`ppt/media/${rel["$"]?.Target}`);

    console.log(" imagePath  --------@@@@@@@@@@@>>>>>>>>>> ", imagePath);

    return `/api/slides/images/${path.basename(imagePath)}`;
  }


  async resolveImagePath(slidePath, imageId) {
    const relPath = `ppt/slides/_rels/${slidePath.split("/").pop()}.rels`;
    const relationshipsXML = await this.parseXML(relPath);

    const relationship = relationshipsXML?.["Relationships"]?.["Relationship"]?.find(
      (rel) => rel["$"]?.Id === imageId
    );

    if (relationship) {
      let targetPath = `ppt/media/${relationship["$"].Target}`;
      // Normalize path to remove redundant ../
      targetPath = path.posix.normalize(targetPath);
      return targetPath;
    }

    return null;
  }


}

module.exports = pptxToHtml;
