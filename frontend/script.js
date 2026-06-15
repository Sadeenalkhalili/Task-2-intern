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

const API_URL = "http://127.0.0.1:8000/translate";

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

function resetApp() {
  appState.selectedFile = null;
  appState.isTranslating = false;
  appState.error = null;
  appState.downloadUrl = null;

  fileInput.value = "";
  uploadText.textContent = "Choose DOCX File";
  filename.textContent = "No file selected";
  statusText.textContent = "";
  translateBtn.disabled = false;

  showScreen(mainScreen);
}

fileInput.addEventListener("change", () => {
  appState.selectedFile = fileInput.files[0];

  const error = validateFile(appState.selectedFile);

  if (error) {
    appState.error = error;
    uploadText.textContent = "Choose DOCX File";
    filename.textContent = error;
    statusText.textContent = "";
    return;
  }

  uploadText.textContent = appState.selectedFile.name;
  filename.textContent = "Press Translate to translate your DOCX";
  statusText.textContent = "";
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