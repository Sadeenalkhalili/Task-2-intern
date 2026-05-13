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

let selectedFile = null;

function showScreen(screen) {
  mainScreen.classList.add("hidden");
  loadingScreen.classList.add("hidden");
  resultScreen.classList.add("hidden");

  screen.classList.remove("hidden");
}

fileInput.addEventListener("change", () => {
  selectedFile = fileInput.files[0];

  if (selectedFile) {
    uploadText.textContent = selectedFile.name;
    filename.textContent = "Press Translate to translate your DOCX";
    statusText.textContent = "";
  } else {
    uploadText.textContent = "Choose DOCX File";
    filename.textContent = "No file selected";
    statusText.textContent = "";
  }
});

translateBtn.addEventListener("click", async () => {
  if (!selectedFile) {
    statusText.textContent = "Please select a DOCX file first.";
    return;
  }

  translateBtn.disabled = true;
  showScreen(loadingScreen);

  const formData = new FormData();
  formData.append("file", selectedFile);

  try {
    const response = await fetch("https://task-2-intern.onrender.com/translate", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      throw new Error("Translation failed. Please try again.");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    downloadLink.href = url;
    downloadLink.download = `translated_${selectedFile.name}`;

    showScreen(resultScreen);
  } catch (error) {
    showScreen(mainScreen);
    statusText.textContent = error.message;
  } finally {
    translateBtn.disabled = false;
  }
});

newFileBtn.addEventListener("click", () => {
  selectedFile = null;
  fileInput.value = "";
  uploadText.textContent = "Choose DOCX File";
  filename.textContent = "No file selected";
  statusText.textContent = "";
  showScreen(mainScreen);
});