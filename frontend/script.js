const fileInput = document.getElementById("fileInput");/*get the file from html*/
const fileName = document.getElementById("fileName");
const translateBtn = document.getElementById("translateBtn");
const statusText = document.getElementById("status");
const downloadLink = document.getElementById("downloadLink");
const mainScreen = document.getElementById("mainScreen");
const resultScreen = document.getElementById("resultScreen");
const uploadText = document.getElementById("uploadText");

let selectedFile = null;

fileInput.addEventListener("change", () => {
  selectedFile = fileInput.files[0];

  const uploadText = document.getElementById("uploadText");

  if (selectedFile) {
    uploadText.textContent = selectedFile.name;
    fileName.textContent = "Press Translate to translate your docx";
    statusText.textContent = "";
  } else {
    uploadText.textContent = "Upload File";
    fileName.textContent = "No file selected";
    statusText.textContent = "";
  }
});

translateBtn.addEventListener("click", async () => {
  if (!selectedFile) {
    alert("Please select a DOCX file first.");
    return;
  }

  statusText.textContent = "Translating...";

  const formData = new FormData();
  formData.append("file", selectedFile);

  try {
    const response = await fetch("https://task-2-intern.onrender.com/translate", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      throw new Error("Translation failed.");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    downloadLink.href = url;
    downloadLink.download = `translated_${selectedFile.name}`;

    mainScreen.classList.add("hidden");
    resultScreen.classList.remove("hidden");

  } catch (error) {
    statusText.textContent = "Error occurred.";
    alert(error.message);
  }
});