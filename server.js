require('dotenv').config(); // Load .env variables
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const { createCanvas, loadImage, registerFont } = require("canvas");
const fs = require("fs");
const cors = require("cors");
const path = require("path");
const archiver = require("archiver");
const nodemailer = require("nodemailer");

const app = express();
app.use(cors());

// Register a custom font if available, otherwise fallback to Arial
const fontPath = path.join(__dirname, "fonts", "CustomFont.ttf");
let defaultFont = "Arial";
if (fs.existsSync(fontPath)) {
  try {
    registerFont(fontPath, { family: "CustomFont" });
    defaultFont = "CustomFont";
  } catch (err) {
    console.warn("⚠️ Failed to register font:", err.message);
  }
} else {
//   console.warn("⚠️ Font file missing. Using system default font.");
}

// Serve static files for certificates and downloads
app.use("/certificates", express.static(path.join(__dirname, "certificates")));
app.use("/downloads", express.static(path.join(__dirname, "downloads")));

// Configure multer for file uploads
const upload = multer({ dest: "uploads/" });

// Configure Nodemailer transporter
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL,
    pass: process.env.APP_PASSWORD
  }
});

// Helper: Ensure directory exists
function ensureDir(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

// (Optional) Helper for dynamic font sizing (currently not used)
// function getDynamicFontSize(ctx, text, maxWidth, baseFontSize, fontFamily) {
//   let fontSize = baseFontSize;
//   const minFontSize = 30;
//   do {
//     ctx.font = `${fontSize}px ${fontFamily}`;
//     if (ctx.measureText(text).width <= maxWidth || fontSize <= minFontSize) break;
//     fontSize -= 2;
//   } while (true);
//   return fontSize;
// }

// Main API endpoint: /upload
// Expects:
// 1. "excel" file: Excel file containing 'Name' and 'Email' columns.
// 2. "template" file: Image file (PNG/JPG) for certificate template.
// 3. Additional form fields: fontStyle, fontColor, fontSize, posX, posY.
app.post("/upload", upload.fields([{ name: "excel" }, { name: "template" }]), async (req, res) => {
  try {
    // Validate file upload
    if (!req.files || !req.files["excel"] || !req.files["template"]) {
      return res.status(400).json({ error: "Both Excel and Template files are required." });
    }

    // Retrieve style settings with fallback defaults
    const chosenFont = req.body.fontStyle || defaultFont;
    const chosenColor = req.body.fontColor || "#000000";
    // Use the fontSize directly from the request (no dynamic resizing)
    const baseFontSize = parseInt(req.body.fontSize) || 60;
    const posX = parseFloat(req.body.posX) || 50; // Percentage horizontally
    const posY = parseFloat(req.body.posY) || 52; // Percentage vertically

    // Get file paths from uploaded files
    const excelPath = req.files["excel"][0].path;
    const templatePath = req.files["template"][0].path;

    // Read Excel file
    const workbook = xlsx.readFile(excelPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    if (data.length === 0) {
      return res.status(400).json({ error: "Excel file is empty or incorrectly formatted." });
    }

    // Prepare the canvas using the template image
    const template = await loadImage(templatePath);
    const canvas = createCanvas(template.width, template.height);
    const ctx = canvas.getContext("2d");

    // Prepare directories for certificates and downloads
    const certificatesDir = path.join(__dirname, "certificates");
    const downloadsDir = path.join(__dirname, "downloads");
    ensureDir(certificatesDir);
    ensureDir(downloadsDir);

    // Clear previous certificates
    fs.readdirSync(certificatesDir).forEach(file => fs.unlinkSync(path.join(certificatesDir, file)));

    const generatedFiles = [];
    let successCount = 0;
    let failCount = 0;

    // Limit text width to 70% of canvas width (if needed for dynamic sizing)
    const maxTextWidth = canvas.width * 0.7;
    const sendMailPromises = [];

    for (const row of data) {
      // Use "Sample Name" if no name is provided
      const name = row["Name"]?.toString().trim() || "Sample Name";
      const email = row["Email"]?.toString().trim();
      const isValidEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
      if (!email || !isValidEmail) continue;

      // Draw the certificate
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(template, 0, 0, canvas.width, canvas.height);

      // Use fixed font size from input (baseFontSize)
      ctx.font = `${baseFontSize}px ${chosenFont}`;
      ctx.fillStyle = chosenColor;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      // Calculate coordinates based on percentages
      const x = (posX / 100) * canvas.width;
      const y = (posY / 100) * canvas.height;
      ctx.fillText(name, x, y);

      // Save certificate as PNG
      const safeName = name.substring(0, 50).replace(/[^a-zA-Z0-9]/g, "_") + ".png";
      const outputPath = path.join(certificatesDir, safeName);
      const out = fs.createWriteStream(outputPath);
      const stream = canvas.createPNGStream();
      stream.pipe(out);
      await new Promise(resolve => out.on("finish", resolve));
      generatedFiles.push(outputPath);

      // Prepare email options
      const mailOptions = {
        from: `"Certificate Generator" <${process.env.EMAIL}>`,
        to: email,
        subject: "Your Certificate",
        text: `Hi ${name},\n\nPlease find your certificate attached.\n\nRegards,\nTeam`,
        attachments: [{ filename: safeName, path: outputPath }]
      };

      // Send mail in parallel
      sendMailPromises.push(
        transporter.sendMail(mailOptions)
          .then(() => {
            console.log(`✅ Sent certificate to ${email}`);
            successCount++;
          })
          .catch(err => {
            console.error(`❌ Failed to send email to ${email}:`, err);
            failCount++;
          })
      );
    }

    // Wait for all emails to finish sending
    await Promise.allSettled(sendMailPromises);

    // Create a ZIP file containing all generated certificates
    const zipFileName = `certificates_${Date.now()}.zip`;
    const zipFilePath = path.join(downloadsDir, zipFileName);
    await new Promise((resolve, reject) => {
      const output = fs.createWriteStream(zipFilePath);
      const archive = archiver("zip", { zlib: { level: 9 } });
      output.on("close", resolve);
      archive.on("error", reject);

      archive.pipe(output);
      generatedFiles.forEach(file => {
        archive.file(file, { name: path.basename(file) });
      });
      archive.finalize();
    });

    // Clean up: delete uploaded Excel and template files
    fs.unlinkSync(excelPath);
    fs.unlinkSync(templatePath);

    // Send success response
    res.json({
      message: "Certificates generated & emails sent successfully!",
      total: generatedFiles.length,
      sent: successCount,
      failed: failCount,
      zipFile: `/downloads/${zipFileName}`
    });
  } catch (error) {
    console.error("Error generating certificates:", error);
    res.status(500).json({ error: "Internal Server Error" });
  }
});

const PORT = 5001;
app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));
