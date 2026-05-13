import shutil#upload file
from pathlib import Path
#backend web API
#fastapi Creates the web application/server
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware#enable frontend-backend communication

from models import translate_docx_xml


app = FastAPI(title="DOCX Translator API")#Create FastAPI app

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "translated_files"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/")
def home():
    return {"message": "DOCX Translator backend is running"}


@app.post("/translate")
async def translate_file(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are allowed")

    input_path = UPLOAD_DIR / file.filename
    output_path = OUTPUT_DIR / f"translated_{file.filename}"

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    translate_docx_xml(
        input_path=str(input_path),
        output_path=str(output_path),
        target_lang="EN",
        source_lang="AR",
        use_free_api=True,
    )

    return FileResponse(
        path=str(output_path),
        filename=f"translated_{file.filename}",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )