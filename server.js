const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const { createCanvas, loadImage, registerFont } = require("canvas");
const fs = require("fs");
const cors = require("cors");
const path = require("path");
const archiver = require("archiver");

const app = express();
app.use(cors());

// Register font (optional)
const fontPath = path.join(__dirname, "fonts", "CustomFont.ttf");
if (fs.existsSync(fontPath)) {
    try {
        registerFont(fontPath, { family: "CustomFont" });
    } catch (err) {
        console.warn("⚠️ Failed to register font:", err.message);
    }
} else {
    console.warn("⚠️ Font file missing. Using system default font.");
}

// Static files for certificates and downloads
app.use("/certificates", express.static(path.join(__dirname, "certificates")));
app.use("/downloads", express.static(path.join(__dirname, "downloads")));

const upload = multer({ dest: "uploads/" });

const ensureDir = (dir) => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
};

const getDynamicFontSize = (ctx, text, maxWidth, baseFontSize) => {
    let fontSize = baseFontSize;
    const minFontSize = 30;

    do {
        ctx.font = `${fontSize}px CustomFont`;
        if (ctx.measureText(text).width <= maxWidth || fontSize <= minFontSize) break;
        fontSize -= 2;
    } while (true);

    return fontSize;
};

app.post("/upload", upload.fields([{ name: "excel" }, { name: "template" }]), async (req, res) => {
    try {
        if (!req.files || !req.files["excel"] || !req.files["template"]) {
            return res.status(400).json({ error: "Both Excel and Template files are required." });
        }

        const excelPath = req.files["excel"][0].path;
        const templatePath = req.files["template"][0].path;

        const workbook = xlsx.readFile(excelPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);

        if (data.length === 0) {
            return res.status(400).json({ error: "Excel file is empty or incorrectly formatted." });
        }

        const template = await loadImage(templatePath);
        const canvas = createCanvas(template.width, template.height);
        const ctx = canvas.getContext("2d");

        const certificatesDir = path.join(__dirname, "certificates");
        const downloadsDir = path.join(__dirname, "downloads");

        ensureDir(certificatesDir);
        ensureDir(downloadsDir);

        const generatedFiles = [];
        const baseFontSize = 80;
        const maxTextWidth = canvas.width * 0.7;

        // Clean old files
        fs.readdirSync(certificatesDir).forEach(file => fs.unlinkSync(path.join(certificatesDir, file)));

        for (const row of data) {
            const name = row["Name"];
            if (!name) continue;

            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(template, 0, 0, canvas.width, canvas.height);

            const fontSize = getDynamicFontSize(ctx, name, maxTextWidth, baseFontSize);
            ctx.font = `${fontSize}px CustomFont`;
            ctx.fillStyle = "black";
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";

            ctx.fillText(name, canvas.width / 2, canvas.height * 0.52);

            const safeName = name.replace(/[^a-zA-Z0-9]/g, "_") + ".png";
            const outputPath = path.join(certificatesDir, safeName);
            const out = fs.createWriteStream(outputPath);
            const stream = canvas.createPNGStream();
            stream.pipe(out);

            generatedFiles.push(outputPath);
            await new Promise((resolve) => out.on("finish", resolve));
        }

        // Create ZIP
        const zipFileName = `certificates_${Date.now()}.zip`;
        const zipFilePath = path.join(downloadsDir, zipFileName);

        const output = fs.createWriteStream(zipFilePath);
        const archive = archiver("zip", { zlib: { level: 9 } });

        output.on("close", () => {
            console.log(`ZIP created: ${zipFilePath} (${archive.pointer()} bytes)`);

            fs.unlinkSync(excelPath);
            fs.unlinkSync(templatePath);

            res.json({
                message: "Certificates generated successfully!",
                zipFile: `/downloads/${zipFileName}`
            });
        });

        archive.on("error", (err) => {
            throw err;
        });

        archive.pipe(output);

        generatedFiles.forEach(file => {
            archive.file(file, { name: path.basename(file) });
        });

        archive.finalize();

    } catch (error) {
        console.error("Error generating certificates:", error);
        res.status(500).json({ error: "Internal Server Error" });
    }
});

const PORT = 5001;
app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));
