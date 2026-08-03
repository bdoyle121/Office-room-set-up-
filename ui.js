/* Drag-and-drop wiring + download trigger. Calls into app.js's convertPdfToWorkbook. */

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const statusEl = document.getElementById("status");

function setStatus(message, kind) {
  statusEl.textContent = message;
  statusEl.className = "status" + (kind ? ` status-${kind}` : "");
}

function slugForFilename(reportDate) {
  if (!reportDate) return "mcshane_setup.xlsx";
  const parsed = new Date(reportDate.replace(/^[A-Za-z]+,\s*/, ""));
  if (isNaN(parsed)) return "mcshane_setup.xlsx";
  const yyyy = parsed.getFullYear();
  const mm = String(parsed.getMonth() + 1).padStart(2, "0");
  const dd = String(parsed.getDate()).padStart(2, "0");
  return `mcshane_setup_${yyyy}-${mm}-${dd}.xlsx`;
}

async function handleFile(file) {
  if (!file) return;
  if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
    setStatus("That doesn't look like a PDF. Please drop the Daily Setup Report PDF.", "error");
    return;
  }

  setStatus("Reading PDF and building the spreadsheet…", "working");

  try {
    const arrayBuffer = await file.arrayBuffer();
    const { reportDate, rows, workbook } = await convertPdfToWorkbook(arrayBuffer);
    const filename = slugForFilename(reportDate);
    XLSX.writeFile(workbook, filename);
    setStatus(`Done — ${rows.length} event${rows.length === 1 ? "" : "s"} for ${reportDate || "this report"}. Downloaded as ${filename}.`, "success");
  } catch (err) {
    console.error(err);
    setStatus(`Couldn't process that PDF: ${err.message}`, "error");
  }
}

dropzone.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", (e) => {
  handleFile(e.target.files[0]);
  fileInput.value = "";
});

["dragenter", "dragover"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  })
);

["dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);

dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  handleFile(file);
});