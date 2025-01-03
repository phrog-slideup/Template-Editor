const { mergeQueue, fileConvertQueue } = require('../queue/nodeQueue'); // Ensure the path is correct
const { Worker } = require('bullmq');
const { spawn } = require('child_process');
const fs = require('fs');
const fileConvertController = require("../controller/fileConvertController");

// Merge Worker
const mergeWorker = new Worker(
  mergeQueue.name,
  async (job) => {
    const { files, mergedDir, exePath } = job.data;

    // Validate files
    if (!Array.isArray(files) || files.length < 2) {
      throw new Error("At least two files are required for merging.");
    }

    // Prepare arguments for the .NET process
    const args = [exePath, ...files, mergedDir];

    return new Promise((resolve, reject) => {
      const dotnetProcess = spawn("dotnet", args);

      let outputPath = "";
      let errorData = "";

      // Handle .NET process stdout
      dotnetProcess.stdout.on("data", (data) => {
        const dataStr = data.toString().trim();
        console.log(`[.NET stdout]: ${dataStr}`);

        if (dataStr.startsWith("Merged file successfully saved:")) {
          outputPath = dataStr.replace("Merged file successfully saved:", "").trim();
        }
      });

      // Handle .NET process stderr
      dotnetProcess.stderr.on("data", (data) => {
        console.error(`[.NET stderr]: ${data.toString()}`);
        errorData += data.toString();
      });

      // Handle process close event
      dotnetProcess.on("close", (code) => {
        console.log(`.NET process exited with code: ${code}`);
        if (code === 0 && outputPath && fs.existsSync(outputPath)) {
          resolve(outputPath);
        } else {
          reject(new Error(errorData || `Process exited with code ${code}`));
        }
      });

      // Handle process errors
      dotnetProcess.on("error", (err) => {
        reject(new Error(`Error spawning .NET process: ${err.message}`));
      });
    });
  },
  {
    connection: mergeQueue.opts.connection, // Reuse the queue connection
  }
);

// File Convert Worker
const fileConvertWorker = new Worker(
  fileConvertQueue.name,
  async (job) => {
    const { type, file } = job.data;

    if (type === "PPTX_TO_HTML") {
      await fileConvertController.convertPptxToHtml(file);
    } else if (type === "HTML_TO_PPTX") {
      await fileConvertController.convertHtmlToPptx(file);
    }
  },
  {
    connection: fileConvertQueue.opts.connection, // Reuse the queue connection
  }
);

// Event Listeners
// Merge Worker Event Listeners
mergeWorker.on("completed", (job, result) => {
  console.log(`Merge Job ${job.id} completed successfully. Result: ${result}`);
});

mergeWorker.on("failed", (job, err) => {
  console.error(`Merge Job ${job.id} failed. Error: ${err.message}`);
});

// File Convert Worker Event Listeners
fileConvertWorker.on("completed", (job) => {
  console.log(`File Convert Job ${job.id} completed successfully!`);
});

fileConvertWorker.on("failed", (job, err) => {
  console.error(`File Convert Job ${job.id} failed. Error: ${err.message}`);
});

module.exports = { mergeWorker, fileConvertWorker };
