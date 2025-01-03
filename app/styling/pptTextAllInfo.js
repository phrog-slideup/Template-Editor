const backgroundColor = require("./pptBackgroundColors");

function getPositionFromShape(shapeNode) {
  const x =
    (shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:off"]?.[0]?.["$"]?.x || 0) / 12700;
  const y =
    (shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:off"]?.[0]?.["$"]?.y || 0) / 12700;
  const width =
    (shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:ext"]?.[0]?.["$"]?.cx || 100) / 12700;
  const height =
    (shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:ext"]?.[0]?.["$"]?.cy || 100) / 12700;
  return { x, y, width, height };
}

function isTextShape(shapeNode) {
  return !!shapeNode?.["p:txBody"];
}

function getTextFromShape(shapeNode) {
  const paragraphs = shapeNode?.["p:txBody"]?.[0]?.["a:p"];
  if (!paragraphs) return "";

  const isBulletPoint = (paragraph) => {
    const pPr = paragraph?.["a:pPr"]?.[0];
    return !!pPr?.["a:buChar"] || !!pPr?.["a:buFont"];
  };

  let htmlContent = "";
  let insideUL = false;

  paragraphs.forEach((paragraph) => {
    if (!paragraph) return;

    const runs = paragraph["a:r"] || [];
    const textParts = runs.map((run) => {
      const textContent = run?.["a:t"]?.[0]?.trim() || "";
      const runProperties = run?.["a:rPr"]?.[0] || {};

      const isBold = runProperties?.["$"]?.b === "1" ? "bold" : "normal";
      const isItalic = runProperties?.["$"]?.i === "1" ? "italic" : "normal";
      const isUnderline = runProperties?.["$"]?.u === "sng" ? "underline" : "none";

      let fontFace = "inherit";
      const latinFont = runProperties?.["a:latin"]?.[0]?.["$"]?.typeface;
      if (latinFont) fontFace = latinFont;

      const fontSize = runProperties?.["$"]?.sz ? `${parseInt(runProperties["$"].sz, 10) / 100}px` : "inherit";

      if (textContent) {
        return `<span style="font-weight: ${isBold}; font-style: ${isItalic}; text-decoration: ${isUnderline}; font-family: ${fontFace}; font-size: ${fontSize};">${textContent}</span>`;
      }
      return "";
    });

    const paragraphText = textParts.join("").trim();
    if (!paragraphText) return;

    if (isBulletPoint(paragraph)) {
      if (!insideUL) {
        htmlContent +=
          '<ul style="margin: 0; padding: 0; list-style-position: outside;">';
        insideUL = true;
      }
      htmlContent += `<li style="margin-bottom: 10px; line-height: 1.5;">${paragraphText}</li>`;
    } else {
      if (insideUL) {
        htmlContent += "</ul>";
        insideUL = false;
      }
      htmlContent += `<p>${paragraphText}</p>`;
    }
  });

  if (insideUL) htmlContent += "</ul>";

  return htmlContent;
}

function getFontSize(shapeNode) {

  const fontSize = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:r"]?.[0]?.["a:rPr"]?.[0];
  return fontSize?.["$"]?.sz ? parseInt(fontSize["$"].sz) / 100 : 16;
  
}

function getFontColor(shapeNode, themeXML, clrMap) {

  const run = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:r"]?.[0]?.["a:rPr"]?.[0];
  if (!run) return "#000000"; // Default to black

  // Handle RGB color
  const colorCode = run?.["a:solidFill"]?.[0]?.["a:srgbClr"]?.[0]?.["$"]?.val;
  if (colorCode) return `#${colorCode}`;

  // Handle theme color
  const schemeClr = run?.["a:solidFill"]?.[0]?.["a:schemeClr"]?.[0]?.["$"]?.val;
  if (schemeClr) {
    const resolvedColor = resolveColor(schemeClr, themeXML, clrMap);
    if (resolvedColor) return resolvedColor;
  }

  return "#000000"; // Default fallback
}

function resolveColor(colorKey, themeXML, clrMap) {
  // Use the mapping to resolve the actual color key
  const mappedKey = clrMap[colorKey] || colorKey;

  const colorNode = getThemeColor(themeXML, mappedKey);
  if (!colorNode) {
    console.error(`Color key '${colorKey}' not found in theme or mapping.`);
    return "#000000"; // Default to black
  }

  return resolveColorFromNode(colorNode, themeXML);
}

function resolveColorFromNode(colorNode, themeXML) {

  // Handle direct RGB color
  if (colorNode?.["a:srgbClr"]) {
    return `#${colorNode["a:srgbClr"][0]["$"]?.val}`;
  }

  // Handle system color
  if (colorNode?.["a:sysClr"]) {
    return `#${colorNode["a:sysClr"][0]["$"]?.lastClr || "FFFFFF"}`;
  }

  // Handle theme color reference
  if (colorNode?.["a:schemeClr"]) {
    const schemeColorKey = colorNode["a:schemeClr"][0]["$"]?.val;

    
    const colorMap = {
      bg1: "lt1", // Light 1
      bg2: "lt2", // Light 2
      tx1: "dk1", // Dark 1
      tx2: "dk2", // Dark 2
    };

    const mappedColorKey = colorMap[schemeColorKey] || schemeColorKey;

    // Resolve luminance modifiers (optional brightness/darkness adjustments)
    const lumMod = colorNode["a:schemeClr"][0]?.["a:lumMod"]?.[0]?.["$"]?.val;
    const lumOff = colorNode["a:schemeClr"][0]?.["a:lumOff"]?.[0]?.["$"]?.val;

    // Resolve base color from the theme
    const baseColor = resolveThemeColor(mappedColorKey, themeXML);

    console.log("baseColor --------->>>>>>>baseColor-------> ", baseColor);

    if (baseColor) {
      return applyLuminanceModifier(baseColor, lumMod, lumOff);
    }
  }

  // Default fallback color
  return null;
}


function getThemeColor(themeXML, schemeKey) {
  const clrScheme =
    themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
  if (!clrScheme) return null;

  const colorNode = clrScheme[`a:${schemeKey}`]?.[0];
  return colorNode || null;
}


function getTextAlignment(shapeNode) {
  const bodyPr = shapeNode?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0];
  const anchor = bodyPr?.["$"]?.anchor;

  const alignmentMap = {
    ctr: "center",
    r: "right",
    t: "top",
    b: "flex-end",
  };

  return alignmentMap[anchor] || "left";
}

function getRotation(shapeNode) {
  const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"];
  return xfrm?.[0]?.["$"]?.rot ? parseInt(xfrm[0]["$"]?.rot, 10) / 60000 : 0;
}

module.exports = {
  getPositionFromShape,
  isTextShape,
  getTextFromShape,
  getFontSize,
  getFontColor,
  getTextAlignment,
  getRotation,
};
