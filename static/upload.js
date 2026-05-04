// Prevent the browser from opening dropped files anywhere on the page
document.addEventListener("dragover", (e) => e.preventDefault());
document.addEventListener("drop", (e) => e.preventDefault());

const input = document.getElementById("gradefile");
const fileLabel = document.getElementById("file-label");
const submitBtn = document.getElementById("submit-btn");
const dropZone = document.getElementById("drop-zone");

function applyFile(file) {
  fileLabel.textContent = "";
  const icon = document.createElement("i");
  icon.className = "bi bi-file-earmark-check me-1";
  icon.style.color = "var(--scdsb-dark)";
  const span = document.createElement("span");
  span.style.color = "var(--scdsb-dark)";
  span.style.fontWeight = "600";
  span.textContent = file.name;  // textContent never parsed as HTML
  fileLabel.appendChild(icon);
  fileLabel.appendChild(span);
  submitBtn.disabled = false;
}

input.addEventListener("change", () => {
  if (input.files.length) applyFile(input.files[0]);
});

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});
dropZone.addEventListener("dragleave", () =>
  dropZone.classList.remove("dragover")
);
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  const files = e.dataTransfer.files;
  if (!files.length) return;
  const dt = new DataTransfer();
  dt.items.add(files[0]);
  input.files = dt.files;
  applyFile(files[0]);
});

// Show spinner while generating
document.getElementById("upload-form").addEventListener("submit", () => {
  submitBtn.disabled = true;
  submitBtn.innerHTML =
    '<span class="spinner-border spinner-border-sm me-2" role="status"></span>Generating Reports\u2026';
});
