const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");

const app = express();

// Middleware for parsing JSON and URL-encoded data
app.use(bodyParser.urlencoded({ limit: "100mb", extended: true }));
app.use(bodyParser.json({ limit: "100mb", extended: true }));

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
