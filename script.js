/*********************************************
 *          Global Drag & Drop Setup
 *********************************************/
const excelTypes = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
];
const imageTypes = ["image/png", "image/jpeg"];

// The entire page is a drop zone:
document.addEventListener("dragover", (e) => {
  e.preventDefault();
});

document.addEventListener("drop", (e) => {
  e.preventDefault();
  if (!e.dataTransfer.files.length) return;

  const file = e.dataTransfer.files[0];
  if (excelTypes.includes(file.type)) {
    // It's an Excel file
    const excelInput = document.getElementById("excelFile");
    excelInput.files = e.dataTransfer.files;
    document.getElementById(
      "excelInfo"
    ).textContent = `Selected file: ${file.name}`;
  } else if (imageTypes.includes(file.type)) {
    // It's an image file
    const templateInput = document.getElementById("templateFile");
    templateInput.files = e.dataTransfer.files;
    document.getElementById(
      "templateInfo"
    ).textContent = `Selected file: ${file.name}`;
    previewTemplate(); // Show preview automatically
  } else {
    alert(
      "Unsupported file type! Only Excel (.xlsx/.xls) or PNG/JPEG allowed."
    );
  }
});

/*********************************************
 *          Preview & Drawing Functions
 *********************************************/
function previewTemplate() {
  const file = document.getElementById("templateFile").files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    templateImg = new Image();
    templateImg.onload = drawPreview;
    templateImg.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function drawPreview() {
  if (!templateImg) return;
  const canvas = document.getElementById("previewCanvas");
  const ctx = canvas.getContext("2d");
  canvas.width = templateImg.width;
  canvas.height = templateImg.height;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(templateImg, 0, 0, canvas.width, canvas.height);

  // Use input value or fallback to "Sample Name"
  const name = document.getElementById("singleName").value || "Sample Name";
  const fontStyle = document.getElementById("fontStyle").value;
  const fontSize = document.getElementById("fontSize").value;
  const fontColor = document.getElementById("fontColor").value;
  const xPercent = document.getElementById("posX").value;
  const yPercent = document.getElementById("posY").value;

  ctx.font = `${fontSize}px ${fontStyle}`;
  ctx.fillStyle = fontColor;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  const x = (xPercent / 100) * canvas.width;
  const y = (yPercent / 100) * canvas.height;
  ctx.fillText(name, x, y);
}

function downloadSingleCertificate() {
  const canvas = document.getElementById("previewCanvas");
  const link = document.createElement("a");
  const name =
    document.getElementById("singleName").value.trim() || "Certificate";
  link.download = `${name.replace(/[^a-zA-Z0-9]/g, "_")}.png`;
  link.href = canvas.toDataURL("image/png");
  link.click();
}

/*********************************************
 *         Bulk Generation Functions
 *********************************************/
async function generateCertificates() {
  const excelFile = document.getElementById("excelFile").files[0];
  const templateFile = document.getElementById("templateFile").files[0];
  const fontStyle = document.getElementById("fontStyle").value;
  const fontSize = document.getElementById("fontSize").value;
  const fontColor = document.getElementById("fontColor").value;
  const xPercent = document.getElementById("posX").value;
  const yPercent = document.getElementById("posY").value;

  if (!excelFile || !templateFile) {
    alert("Please upload both files!");
    return;
  }

  // Show progress UI
  const progressContainer = document.getElementById("progressContainer");
  const progressText = document.getElementById("progressText");
  const progressBar = document.getElementById("progressBar");
  progressContainer.style.display = "block";
  const generateButton = document.querySelector("button");
  generateButton.disabled = true;
  generateButton.innerText = "Generating...";

  // Read Excel file
  const reader = new FileReader();
  reader.readAsArrayBuffer(excelFile);
  reader.onload = async function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    if (jsonData.length === 0) {
      alert("Excel file is empty or incorrectly formatted.");
      return;
    }

    // Read template image
    const templateReader = new FileReader();
    templateReader.readAsDataURL(templateFile);
    templateReader.onload = async function (e) {
      const templateSrc = e.target.result;
      const img = new Image();
      img.src = templateSrc;
      img.onload = async function () {
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        canvas.width = img.width;
        canvas.height = img.height;
        const zip = new JSZip();
        let count = 0;

        for (const row of jsonData) {
          const name = row["Name"];
          if (!name) continue;

          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

          ctx.font = `${fontSize}px ${fontStyle}`;
          ctx.fillStyle = fontColor;
          ctx.textAlign = "center";
          ctx.textBaseline = "middle";

          const x = (xPercent / 100) * canvas.width;
          const y = (yPercent / 100) * canvas.height;
          ctx.fillText(name, x, y);

          const imgData = canvas.toDataURL("image/png");
          const response = await fetch(imgData);
          const blob = await response.blob();
          const safeName = name.replace(/[^a-zA-Z0-9]/g, "_") + ".png";
          zip.file(safeName, blob);

          count++;
          const percent = Math.round((count / jsonData.length) * 100);
          progressBar.style.width = percent + "%";
          progressText.innerText = `Generating ${count}/${jsonData.length} certificates...`;
        }

        zip.generateAsync({ type: "blob" }).then((content) => {
          saveAs(content, "Certificates.zip");
          progressText.innerText = "✅ Completed!";
          generateButton.disabled = false;
          generateButton.innerText = "Generate and Download ZIP";
        });
      };
    };
  };
}

/*********************************************
 *      Server Email Submission
 *********************************************/
document.getElementById("uploadForm").addEventListener("submit", async function (e) {
  e.preventDefault(); // Prevent page reload
  document.getElementById("status").innerText = "Processing...";

  const formData = new FormData();
  const excelFile = document.getElementById("excelFile").files[0];
  const templateFile = document.getElementById("templateFile").files[0];
  if (!excelFile || !templateFile) {
    document.getElementById("status").innerText = "Please upload both files first.";
    return;
  }
  formData.append("excel", excelFile);
  formData.append("template", templateFile);
  // Append style settings
  formData.append("fontStyle", document.getElementById("fontStyle").value);
  formData.append("fontColor", document.getElementById("fontColor").value);
  formData.append("fontSize", document.getElementById("fontSize").value);
  formData.append("posX", document.getElementById("posX").value);
  formData.append("posY", document.getElementById("posY").value);

  try {
    const response = await fetch("http://localhost:5001/upload", {
      method: "POST",
      body: formData,
    });
    const result = await response.json();
    if (response.ok) {
      // Redirect to the email sent page
      window.location.href = "emailSent.html";
    } else {
      document.getElementById("status").innerText = result.error || "Something went wrong.";
    }
  } catch (err) {
    document.getElementById("status").innerText = "Error: " + err.message;
  }
});
 


/*********************************************
 *            Drag & Drop Setup
 *********************************************/
function setupDropZone(dropZoneId, inputId, onChangeCallback) {
  const dropZone = document.getElementById(dropZoneId);
  const input = document.getElementById(inputId);
  const dropText = dropZone.querySelector("p");

  dropZone.addEventListener("click", () => input.click());
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      dropText.innerHTML = `✅ <strong>${input.files[0].name}</strong> uploaded`;
      if (onChangeCallback) onChangeCallback();
    }
  });
  input.addEventListener("change", () => {
    if (input.files.length) {
      dropText.innerHTML = `✅ <strong>${input.files[0].name}</strong> uploaded`;
      if (onChangeCallback) onChangeCallback();
    }
  });
}

// Initialize drop zones for Excel and Template
setupDropZone("excelDropZone", "excelFile");
setupDropZone("templateDropZone", "templateFile", previewTemplate);

// Global drag & drop handler (prevent default behavior)
document.addEventListener("dragover", (e) => e.preventDefault());
document.addEventListener("drop", (e) => e.preventDefault());
