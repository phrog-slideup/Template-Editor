"use strict";

const express = require("express");
const router = express.Router();
const multer = require("multer");
const path = require("path");
const pptxToHtmlConverter = require("../controllers/pptxToHtmlConverter.controller");
const htmlToPptxConverter = require('../controllers/htmlToPptxConverter.controller');


// Configure multer for file uploads
// const upload = multer({ dest: path.join(__dirname, "../../../uploads/") });
const upload = multer({ storage: multer.memoryStorage() });

// Define routes
router.post("/upload", upload.single("pptxFile"), pptxToHtmlConverter.uploadFile); // Upload PPTX file
router.get("/getSlides", pptxToHtmlConverter.getSlides); // Get slides as HTML
router.post("/saveSlides", pptxToHtmlConverter.saveSlides); // Save updated slides
router.get("/downloadHTML", pptxToHtmlConverter.downloadHTML); // Download converted HTML
router.post("/convertToPPTX", htmlToPptxConverter.convertToPPTX); // Convert HTML to PPTX

router.get("/images/:key", pptxToHtmlConverter.serveImage);


module.exports = router;

