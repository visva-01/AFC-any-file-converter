const fileInput = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const fileSummary = document.getElementById("fileSummary");
const fileName = document.getElementById("fileName");
const fileMeta = document.getElementById("fileMeta");
const targetFormat = document.getElementById("targetFormat");
const backendChip = document.getElementById("backendChip");
const qualityWrap = document.getElementById("qualityWrap");
const qualityRange = document.getElementById("qualityRange");
const qualityLabel = document.getElementById("qualityLabel");
const convertButton = document.getElementById("convertButton");
const downloadLink = document.getElementById("downloadLink");
const statusBox = document.getElementById("statusBox");
const previewArea = document.getElementById("previewArea");
const imagePreviewTemplate = document.getElementById("imagePreviewTemplate");
const textPreviewTemplate = document.getElementById("textPreviewTemplate");

const IMAGE_FORMATS = ["png", "jpg", "webp"];
const TEXT_EXTENSIONS = ["txt", "md", "csv", "html", "css", "js", "xml"];
const JSON_EXTENSIONS = ["json"];
const PREVIEWABLE_TEXT_EXTENSIONS = [...TEXT_EXTENSIONS, ...JSON_EXTENSIONS];
const OFFICE_INPUTS = ["ppt", "pptx", "doc", "docx", "xls", "xlsx"];

let currentFile = null;
let currentObjectUrl = null;
let backendStatus = { available: false };

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) {
    handleSelectedFile(file);
  }
});

convertButton.addEventListener("click", () => {
  if (currentFile) {
    void convertCurrentFile();
  }
});

targetFormat.addEventListener("change", () => {
  syncQualityVisibility();
});

qualityRange.addEventListener("input", () => {
  qualityLabel.textContent = `${Math.round(Number(qualityRange.value) * 100)}%`;
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragging");
  });
});

dropzone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (file) {
    handleSelectedFile(file);
  }
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(() => {
      setStatus("Offline cache could not be installed, but conversion still works while this tab is open.");
    });
  });
}

void loadBackendStatus();

function handleSelectedFile(file) {
  revokeDownloadUrl();
  currentFile = file;

  const extension = getExtension(file.name);
  const category = detectCategory(file, extension);
  const targets = getTargetFormats(category, extension);

  fileSummary.hidden = false;
  fileName.textContent = file.name;
  fileMeta.textContent = `${formatBytes(file.size)} | ${category.label}`;

  populateTargetFormats(targets);
  renderPreview(file, extension, category);

  syncQualityVisibility();
  convertButton.disabled = targets.length === 0;

  if (targets.length === 0) {
    setStatus("This file type is not supported in the current offline build.");
  } else {
    setStatus(`Ready to convert ${extension ? `.${extension}` : "this file"} into ${targets.join(", ")}.`);
  }
}

function populateTargetFormats(targets) {
  targetFormat.innerHTML = "";

  if (!targets.length) {
    const option = new Option("No available conversions", "");
    targetFormat.add(option);
    targetFormat.disabled = true;
    return;
  }

  targetFormat.disabled = false;
  targets.forEach((format) => {
    targetFormat.add(new Option(format.toUpperCase(), format));
  });
}

function syncQualityVisibility() {
  qualityWrap.hidden = !["jpg", "webp"].includes(targetFormat.value);
}

function detectCategory(file, extension) {
  if (OFFICE_INPUTS.includes(extension)) {
    return { kind: "office", label: "Office document" };
  }

  if (file.type.startsWith("image/") || extension === "svg") {
    return { kind: "image", label: "Image file" };
  }

  if (JSON_EXTENSIONS.includes(extension)) {
    return { kind: "json", label: "JSON file" };
  }

  if (TEXT_EXTENSIONS.includes(extension) || file.type.startsWith("text/")) {
    return { kind: "text", label: "Text file" };
  }

  return { kind: "unsupported", label: "Unsupported in browser-only mode" };
}

function getTargetFormats(category, extension) {
  if (category.kind === "image") {
    return IMAGE_FORMATS.filter((format) => format !== normalizeExtension(extension));
  }

  if (category.kind === "office") {
    return ["pdf"];
  }

  if (category.kind === "json") {
    return ["json", "txt"];
  }

  if (category.kind === "text") {
    return ["txt", "md", "html"].filter((format) => format !== normalizeExtension(extension));
  }

  return [];
}

async function renderPreview(file, extension, category) {
  previewArea.innerHTML = "";

  if (category.kind === "office") {
    previewArea.innerHTML = `<p class="placeholder-copy">${file.name} will be converted to PDF by the local desktop helper.</p>`;
    return;
  }

  if (category.kind === "image") {
    const image = imagePreviewTemplate.content.firstElementChild.cloneNode(true);
    image.src = URL.createObjectURL(file);
    image.onload = () => URL.revokeObjectURL(image.src);
    previewArea.append(image);
    return;
  }

  if (PREVIEWABLE_TEXT_EXTENSIONS.includes(extension) || category.kind === "text" || category.kind === "json") {
    const text = await file.text();
    const textPreview = textPreviewTemplate.content.firstElementChild.cloneNode(true);
    textPreview.textContent = category.kind === "json" ? safeFormatJson(text) : text.slice(0, 30000);
    previewArea.append(textPreview);
    return;
  }

  previewArea.innerHTML = '<p class="placeholder-copy">Preview is not available for this file type yet.</p>';
}

async function convertCurrentFile() {
  const target = targetFormat.value;
  const extension = getExtension(currentFile.name);
  const category = detectCategory(currentFile, extension);

  if (!target) {
    setStatus("Pick an output format first.");
    return;
  }

  convertButton.disabled = true;
  setStatus("Converting locally...");

  try {
    let result;

    if (category.kind === "image") {
      result = await convertImage(currentFile, target);
    } else if (category.kind === "office") {
      result = await convertOffice(currentFile, target);
    } else if (category.kind === "json") {
      result = await convertJson(currentFile, target);
    } else if (category.kind === "text") {
      result = await convertText(currentFile, target);
    } else {
      throw new Error("This file type is not supported in the current build.");
    }

    prepareDownload(result.blob, result.filename || replaceExtension(currentFile.name, result.extension));
    setStatus(`Converted successfully to ${result.extension.toUpperCase()}.`);
  } catch (error) {
    setStatus(error.message || "Conversion failed.");
  } finally {
    convertButton.disabled = false;
  }
}

async function convertImage(file, targetFormatName) {
  const sourceExtension = getExtension(file.name);
  const imageBitmap = await loadImageBitmap(file, sourceExtension);
  const canvas = document.createElement("canvas");
  canvas.width = imageBitmap.width;
  canvas.height = imageBitmap.height;

  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(imageBitmap, 0, 0);

  const mimeType = imageFormatToMime(targetFormatName);
  const quality = Number(qualityRange.value);

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob((createdBlob) => {
      if (!createdBlob) {
        reject(new Error("The browser could not export this image format."));
        return;
      }
      resolve(createdBlob);
    }, mimeType, quality);
  });

  return { blob, extension: targetFormatName };
}

async function convertText(file, targetFormatName) {
  const text = await file.text();
  let output = text;

  if (targetFormatName === "html") {
    output = wrapAsHtml(text, file.name);
  } else if (targetFormatName === "md") {
    output = text;
  } else if (targetFormatName === "txt") {
    output = stripHtmlIfNeeded(text, getExtension(file.name));
  }

  return {
    blob: new Blob([output], { type: mimeForTextExtension(targetFormatName) }),
    extension: targetFormatName
  };
}

async function convertJson(file, targetFormatName) {
  const text = await file.text();
  const formatted = safeFormatJson(text, true);

  if (targetFormatName === "txt") {
    return {
      blob: new Blob([formatted], { type: "text/plain;charset=utf-8" }),
      extension: "txt"
    };
  }

  return {
    blob: new Blob([formatted], { type: "application/json;charset=utf-8" }),
    extension: "json"
  };
}

async function convertOffice(file, targetFormatName) {
  if (targetFormatName !== "pdf") {
    throw new Error("Office files currently export only to PDF.");
  }

  if (!backendStatus.available) {
    throw new Error("Desktop helper is offline. Start the Python server to convert Office files.");
  }

  const formData = new FormData();
  formData.append("file", file, file.name);
  formData.append("target", targetFormatName);

  const response = await fetch("/api/convert", {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    const payload = await safeReadJson(response);
    throw new Error(payload?.error || "Desktop conversion failed.");
  }

  const blob = await response.blob();
  const filename = getDownloadFilename(response.headers.get("Content-Disposition")) || replaceExtension(file.name, "pdf");

  return { blob, extension: getExtension(filename) || "pdf", filename };
}

function prepareDownload(blob, filename) {
  revokeDownloadUrl();
  currentObjectUrl = URL.createObjectURL(blob);
  downloadLink.href = currentObjectUrl;
  downloadLink.download = filename;
  downloadLink.hidden = false;
}

function revokeDownloadUrl() {
  if (currentObjectUrl) {
    URL.revokeObjectURL(currentObjectUrl);
    currentObjectUrl = null;
  }
  downloadLink.hidden = true;
  downloadLink.removeAttribute("href");
}

function setStatus(message) {
  statusBox.textContent = message;
}

function getExtension(filename) {
  const parts = filename.split(".");
  return parts.length > 1 ? normalizeExtension(parts.pop()) : "";
}

function normalizeExtension(extension) {
  return extension.toLowerCase().replace(/^\./, "");
}

function replaceExtension(filename, newExtension) {
  const index = filename.lastIndexOf(".");
  const baseName = index >= 0 ? filename.slice(0, index) : filename;
  return `${baseName}.${newExtension}`;
}

function getDownloadFilename(contentDisposition) {
  if (!contentDisposition) {
    return "";
  }

  const match = contentDisposition.match(/filename="?([^"]+)"?/i);
  return match ? match[1] : "";
}

function formatBytes(size) {
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}

function imageFormatToMime(format) {
  switch (format) {
    case "jpg":
      return "image/jpeg";
    case "webp":
      return "image/webp";
    default:
      return "image/png";
  }
}

async function loadImageBitmap(file, extension) {
  if (extension === "svg") {
    const svgText = await file.text();
    const encoded = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgText)}`;
    return await loadImageFromUrl(encoded);
  }

  const objectUrl = URL.createObjectURL(file);
  try {
    return await loadImageFromUrl(objectUrl);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function loadImageFromUrl(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("The image could not be decoded by this browser."));
    image.src = url;
  });
}

function mimeForTextExtension(extension) {
  switch (extension) {
    case "html":
      return "text/html;charset=utf-8";
    case "md":
      return "text/markdown;charset=utf-8";
    default:
      return "text/plain;charset=utf-8";
  }
}

function stripHtmlIfNeeded(text, extension) {
  if (extension !== "html") {
    return text;
  }

  const container = document.createElement("div");
  container.innerHTML = text;
  return container.textContent || container.innerText || "";
}

function wrapAsHtml(text, title) {
  const escaped = escapeHtml(text);
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Georgia, serif; margin: 2rem; line-height: 1.6; background: #fffaf0; color: #1b1f18; }
    main { max-width: 72ch; margin: 0 auto; }
    pre { white-space: pre-wrap; word-break: break-word; font-family: "Courier New", monospace; }
  </style>
</head>
<body>
  <main>
    <pre>${escaped}</pre>
  </main>
</body>
</html>`;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function safeFormatJson(text, strict = false) {
  try {
    return JSON.stringify(JSON.parse(text), null, 2);
  } catch (error) {
    if (strict) {
      throw new Error("Invalid JSON file.");
    }
    return text.slice(0, 30000);
  }
}

async function loadBackendStatus() {
  try {
    const response = await fetch("/api/status");
    if (!response.ok) {
      throw new Error("Status endpoint unavailable");
    }

    backendStatus = await response.json();
    backendChip.textContent = backendStatus.available
      ? "Desktop helper: online"
      : "Desktop helper: unavailable";
    backendChip.classList.toggle("is-warning", !backendStatus.available);
  } catch (error) {
    backendStatus = { available: false };
    backendChip.textContent = "Desktop helper: not running";
    backendChip.classList.add("is-warning");
  }
}

async function safeReadJson(response) {
  try {
    return await response.json();
  } catch (error) {
    return null;
  }
}
