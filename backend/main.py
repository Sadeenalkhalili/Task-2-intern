import shutil
import logging
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from models import translate_docx_xml


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="DOCX Translator API")

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


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error(f"Unhandled error: {exc}")
    return JSONResponse(
        status_code=500,
        content={
            "error": "Internal server error",
            "message": "Something went wrong while processing the document.",
        },
    )


@app.get("/")
def home():
    return {"message": "DOCX Translator backend is running"}


@app.get("/health")
def health_check():
    return {
        "status": "ok",
        "service": "DOCX Translator API",
    }


@app.get("/supported-formats")
def supported_formats():
    return {
        "formats": [".docx"],
        "source_language": "Arabic",
        "target_language": "English",
    }


@app.post("/translate")
async def translate_file(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are allowed")

    input_path = UPLOAD_DIR / file.filename
    output_path = OUTPUT_DIR / f"translated_{file.filename}"

    logger.info(f"Received file: {file.filename}")

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    logger.info("Starting translation process...")

    translate_docx_xml(
        input_path=str(input_path),
        output_path=str(output_path),
        target_lang="EN",
        source_lang="AR",
        use_free_api=True,
    )

    logger.info(f"Translation completed: {output_path.name}")

    return FileResponse(
        path=str(output_path),
        filename=f"translated_{file.filename}",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )