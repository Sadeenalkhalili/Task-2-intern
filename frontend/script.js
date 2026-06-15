const fileInput = document.getElementById("fileInput");
const filename = document.getElementById("filename");
const translateBtn = document.getElementById("translateBtn");
const statusText = document.getElementById("status");
const downloadLink = document.getElementById("downloadLink");
const mainScreen = document.getElementById("mainScreen");
const loadingScreen = document.getElementById("loadingScreen");
const resultScreen = document.getElementById("resultScreen");
const uploadText = document.getElementById("uploadText");
const newFileBtn = document.getElementById("newFileBtn");
const dropArea = document.getElementById("dropArea");
const removeFileBtn = document.getElementById("removeFileBtn");

const API_URL = "https://task-2-intern.onrender.com/translate";

const appState = {
  selectedFile: null,
  isTranslating: false,
  error: null,
  downloadUrl: null,
};

function showScreen(screen) {
  mainScreen.classList.add("hidden");
  loadingScreen.classList.add("hidden");
  resultScreen.classList.add("hidden");
  screen.classList.remove("hidden");
}

function validateFile(file) {
  if (!file) {
    return "Please select a DOCX file first.";
  }

  if (!file.name.toLowerCase().endsWith(".docx")) {
    return "Only DOCX files are allowed.";
  }

  if (file.size > 10 * 1024 * 1024) {
    return "File must be less than 10MB.";
  }

  return null;
}

function handleSelectedFile(file) {

  appState.selectedFile = file;

  const error = validateFile(file);

  if (error) {
    appState.error = error;

    uploadText.innerHTML =
      "Drag & drop your DOCX here<br>or click to browse";

    filename.textContent = error;

    statusText.textContent = "";

    removeFileBtn.classList.add("hidden");

    return;
  }

  appState.error = null;

  uploadText.textContent = file.name;

  filename.textContent =
    "Press Translate to translate your DOCX";

  statusText.textContent = "";

  removeFileBtn.classList.remove("hidden");
}

function resetApp() {
  appState.selectedFile = null;
  appState.isTranslating = false;
  appState.error = null;
  appState.downloadUrl = null;

  fileInput.value = "";
  uploadText.innerHTML = "Drag & drop your DOCX here<br>or click to browse";
  filename.textContent = "No file selected";
  statusText.textContent = "";
  translateBtn.disabled = false;
  removeFileBtn.classList.add("hidden");
  showScreen(mainScreen);
}

function clearSelectedFile() {

  appState.selectedFile = null;

  fileInput.value = "";

  uploadText.innerHTML =
    "Drag & drop your DOCX here<br>or click to browse";

  filename.textContent = "No file selected";

  statusText.textContent = "";

  removeFileBtn.classList.add("hidden");
}

fileInput.addEventListener("change", () => {
  handleSelectedFile(fileInput.files[0]);
});

removeFileBtn.addEventListener("click", clearSelectedFile);

dropArea.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropArea.classList.add("drag-over");
});

dropArea.addEventListener("dragleave", () => {
  dropArea.classList.remove("drag-over");
});

dropArea.addEventListener("drop", (event) => {
  event.preventDefault();
  dropArea.classList.remove("drag-over");

  const droppedFile = event.dataTransfer.files[0];
  handleSelectedFile(droppedFile);
});

translateBtn.addEventListener("click", async () => {
  const error = validateFile(appState.selectedFile);

  if (error) {
    statusText.textContent = error;
    return;
  }

  appState.isTranslating = true;
  appState.error = null;

  translateBtn.disabled = true;
  showScreen(loadingScreen);

  const formData = new FormData();
  formData.append("file", appState.selectedFile);

  try {
    const response = await fetch(API_URL, {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      throw new Error("Translation failed. Please try again.");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    appState.downloadUrl = url;

    downloadLink.href = url;
    downloadLink.download = `translated_${appState.selectedFile.name}`;

    showScreen(resultScreen);
  } catch (error) {
    appState.error = error.message;
    showScreen(mainScreen);
    statusText.textContent = error.message;
  } finally {
    appState.isTranslating = false;
    translateBtn.disabled = false;
  }
});

newFileBtn.addEventListener("click", resetApp);