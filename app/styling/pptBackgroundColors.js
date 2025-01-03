
const path = require("path");
const fs = require("fs");

async function getBackgroundColor(slideXML, masterXML, themeXML, relationshipsXML) {
  try {
    const cSld = slideXML?.["p:sld"]?.["p:cSld"]?.[0];
    const bgPr = cSld?.["p:bg"]?.[0]?.["p:bgPr"]?.[0];

    // Check for solid fill with direct RGB color
    const solidFill = bgPr?.["a:solidFill"]?.[0];

    if (solidFill?.["a:srgbClr"]) {
      return `#${solidFill["a:srgbClr"][0]["$"]?.val}`;
    }

    // Check for theme color reference
    const schemeClr = solidFill?.["a:schemeClr"]?.[0]?.["$"]?.val;

    if (schemeClr) {
      const lumMod = solidFill?.["a:schemeClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;
      const lumOff = solidFill?.["a:schemeClr"]?.[0]?.["a:lumOff"]?.[0]?.["$"]?.val;

      const baseColor = resolveColor(schemeClr, themeXML);

      if (baseColor) {
        return applyLuminanceModifier(baseColor, lumMod, lumOff);
      }
    }

    // Check for gradient fill
    const gradFill = bgPr?.["a:gradFill"]?.[0];
    
    console.log("gradFill ------->>>>>>>>> ", gradFill);
    if (gradFill) {
      const gradientCSS = getGradientFillColor(gradFill, themeXML);

      if (gradientCSS) {
        return gradientCSS; //// Return CSS gradient
      }
    }
    
    // Check for pattern fill
    const pattFill = bgPr?.["a:pattFill"]?.[0];
    if (pattFill) {
      const patternCSS = getPatternFillCSS(pattFill, themeXML);

      console.log("Check for pattern fill --------->> ",patternCSS);
      if (patternCSS) return patternCSS;
    }

     // // Check for picture or texture Fill
    const pictureTextureFill = bgPr?.["a:blipFill"]?.[0];
    if(pictureTextureFill)
    {
      const pictureTextureCSS = getPictureTextureFillCSS(pictureTextureFill, relationshipsXML);
      
      console.log("pictureTextureCSS -----------FINAL~~~~~========== >>>> ", pictureTextureCSS);

      if(pictureTextureCSS){
        return pictureTextureCSS; // Return Picture Or Texture Fill CSS.
      }
    }
    // console.log("pictureTextureCSS -----------FINAL~~~~~========== >>>> ", pictureTextureCSS);

    // Check for background reference in master slide
    const masterBgRef = masterXML?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:bg"]?.[0]?.["p:bgRef"]?.[0];

    if (masterBgRef?.["a:schemeClr"]) {
      const lumMod = masterBgRef?.["a:schemeClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;
      const lumOff = masterBgRef?.["a:schemeClr"]?.[0]?.["a:lumOff"]?.[0]?.["$"]?.val;

      const baseColor = resolveColor(masterBgRef["a:schemeClr"][0]["$"]?.val, themeXML);
      if (baseColor) {
        return applyLuminanceModifier(baseColor, lumMod, lumOff);
      }
    }

    // Default fallback color
    return "";
  } catch (error) {
    console.error("Error resolving background color:", error);
    return ""; // Fallback in case of error
  }
}

function resolveColor(colorKey, themeXML) {
  const colorMap = {
    bg1: "lt1", // Light 1
    bg2: "lt2", // Light 2
    tx1: "dk1", // Dark 1
    tx2: "dk2", // Dark 2
  };

  const mappedColorKey = colorMap[colorKey] || colorKey;
  const colorNode = getThemeColor(themeXML, mappedColorKey);
  if (!colorNode) return "";
  return resolveColorFromNode(colorNode);
}

function getThemeColor(themeXML, schemeKey) {
  const clrScheme =
    themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
  if (!clrScheme) return null;

  const colorNode = clrScheme[`a:${schemeKey}`]?.[0];
  return colorNode || null;
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


function applyLuminanceModifier(hexColor, lumMod, lumOff) {
  if (!hexColor) return "";

  let r = parseInt(hexColor.substring(1, 3), 16);
  let g = parseInt(hexColor.substring(3, 5), 16);
  let b = parseInt(hexColor.substring(5, 7), 16);

  if (lumMod) {
    const factor = lumMod / 100000;
    r = Math.round(r * factor);
    g = Math.round(g * factor);
    b = Math.round(b * factor);
  }

  if (lumOff) {
    const offset = lumOff / 100000;
    r = Math.min(255, Math.round(r + 255 * offset));
    g = Math.min(255, Math.round(g + 255 * offset));
    b = Math.min(255, Math.round(b + 255 * offset));
  }

  r = Math.max(0, Math.min(255, r));
  g = Math.max(0, Math.min(255, g));
  b = Math.max(0, Math.min(255, b));

  return `#${r.toString(16).padStart(2, "0")}${g
    .toString(16)
    .padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
}

function getGradientFillColor(gradFill, themeXML) {

  const stops = gradFill?.["a:gsLst"]?.[0]?.["a:gs"] || [];

  // Extract gradient angle or default to 90 degrees (vertical gradient)
  const angle = parseInt(gradFill?.["a:lin"]?.[0]?.["$"]?.ang, 10) || 5400000; // Default angle: 90 degrees
  const gradientAngle = angle / 60000; // Convert angle to degrees

  let gradientStops = [];

  stops.forEach((stop) => {
    const position = parseInt(stop["$"]?.pos, 10) / 1000; // Convert position to percentage
    const schemeClr = stop?.["a:schemeClr"]?.[0]?.["$"]?.val;
    const srgbClr = stop?.["a:srgbClr"]?.[0]?.["$"]?.val;

    let color = null;

    if (srgbClr) {
      color = `#${srgbClr}`; // Direct RGB color
    } else if (schemeClr) {
      let baseColor = resolveColor(schemeClr, themeXML);

      // Apply luminance modifiers
      const lumMod = stop?.["a:schemeClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;
      const lumOff = stop?.["a:schemeClr"]?.[0]?.["a:lumOff"]?.[0]?.["$"]?.val;
      color = applyLuminanceModifier(baseColor, lumMod, lumOff);
    }

    if (color) {
      gradientStops.push(`${color} ${position}%`);
    }
  });

  // Build CSS gradient
  if (gradientStops.length > 0) {
    const stopsCSS = gradientStops.join(", ");
    
    return `linear-gradient(${gradientAngle}deg, ${stopsCSS})`;
  }

  return null; // Default if no gradient stops
}

function getPictureTextureFillCSS (blipFill, relationshipsXML){

  if (!relationshipsXML || !relationshipsXML.Relationships) {
    console.error("Invalid or missing relationships XML.");
    return null;
  }

  const embedId = blipFill?.["a:blip"]?.[0]?.["$"]?.["r:embed"];
  if (!embedId) return null;

  const relationship = relationshipsXML.Relationships.Relationship.find(
    (rel) => rel["$"]?.Id === embedId
  );

  if (relationship) {
    const imagePath = relationship["$"]?.Target;
    console.log(`Resolved image path: ${imagePath}`);
    return `url('${imagePath}')`;
  }

  return null;
};


function getPatternFillCSS(pattFill, themeXML) {
  const fgClrNode = pattFill?.["a:fgClr"]?.[0];

  console.log("fgClrNode =--------- fgClrNode >>>>>>>>> ", fgClrNode);

  const fgColor = resolveColorFromNode(fgClrNode, themeXML);

  console.log("fgColor =--------- fgColor >>>>>>>>> ", fgColor);

  const bgClrNode = pattFill?.["a:bgClr"]?.[0];

  console.log("bgClrNode =--------- bgClrNode >>>>>>>>> ", bgClrNode);

  const bgColor = resolveColorFromNode(bgClrNode, themeXML);

  console.log("bgColor =--------- bgColor >>>>>>>>> ", bgColor);


  // Map preset patterns to CSS equivalents
  const patternMap = {
    // Percentage Patterns
    pct5: "5%",
    pct10: "10%",
    pct20: "20%",
    pct25: "25%",
    pct30: "30%",
    pct40: "40%",
    pct50: "50%",
    pct60: "60%",
    pct70: "70%",
    pct80: "80%",
    pct90: "90%",
  
    // Line-Based Patterns
    horzStripe: "horizontal-stripes",
    thinHorzStripe: "thin-horizontal-stripes",
    vertStripe: "vertical-stripes",
    thinVertStripe: "thin-vertical-stripes",
    diagStripe: "diagonal-stripes",
    thinDiagStripe: "thin-diagonal-stripes",
    zigZag: "zigzag-lines",
  
    // Grid Patterns
    dkGrid: "dark-grid",
    ltGrid: "light-grid",
    squareGrid: "square-grid",
    hexGrid: "hexagonal-grid",
    dottedGrid: "dotted-grid",
  
    // Cross Patterns
    dkDiagCross: "dark-diagonal-cross",
    ltDiagCross: "light-diagonal-cross",
    solidCross: "solid-cross",
  
    // Special Patterns
    weave: "weave-pattern",
    wave: "wave-pattern",
  };
  

  const pattern = pattFill?.["$"]?.prst || "pct5"; // Default to 5% pattern
  const cssPattern = patternMap[pattern] || "none";

  console.log("cssPattern ------->>>>>>>>>>>> ", cssPattern);

  // CSS generation for patterns
  if (fgColor && bgColor && cssPattern) {
    return `
      background-color: ${bgColor};
      background-image: repeating-linear-gradient(${cssPattern}, ${fgColor} 0%, ${fgColor} 50%, ${bgColor} 50%, ${bgColor} 100%);
    `;
  }

  return null;
}

function resolveThemeColor(colorKey, themeXML) {
  const clrScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
  if (!clrScheme) return null;

  const colorNode = clrScheme[`a:${colorKey}`]?.[0];
  if (!colorNode) return null;

  // Handle direct RGB color
  if (colorNode?.["a:srgbClr"]) {
    return `#${colorNode["a:srgbClr"][0]["$"]?.val}`;
  }

  // Handle system color
  if (colorNode?.["a:sysClr"]) {
    return `#${colorNode["a:sysClr"][0]["$"]?.lastClr || "FFFFFF"}`;
  }

  return null;
}





module.exports = {
  getBackgroundColor,
  resolveColor,
  getThemeColor,
  resolveColorFromNode,
  applyLuminanceModifier,
  getGradientFillColor,
  getPictureTextureFillCSS,
  resolveThemeColor
};
