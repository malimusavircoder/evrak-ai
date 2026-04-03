import os
import json
import base64
import shutil
import re
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from jinja2 import Environment, FileSystemLoader, select_autoescape
from openpyxl import Workbook
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
EXPORT_DIR = BASE_DIR / "exports"
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"

load_dotenv(BASE_DIR / ".env", override=True, encoding="utf-8-sig")

UPLOAD_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

CLAUDE_API_KEY = os.getenv("CLAUDE_API_KEY", "")
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-opus-4-6")
MODEL_ALIASES = {
    "claude-opus-4.6": "claude-opus-4-6",
}


def normalize_model_name(value: str) -> str:
    model_name = (value or "").strip()
    if "=" in model_name:
        model_name = model_name.split("=")[-1].strip()
    return MODEL_ALIASES.get(model_name, model_name or "claude-opus-4-6")

DATABASE_URL = f"sqlite:///{DATA_DIR / 'app.db'}"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

app = FastAPI(title="Evrak AI")

jinja_env = Environment(
    loader=FileSystemLoader(str(BASE_DIR / "templates")),
    autoescape=select_autoescape(["html", "xml"]),
    cache_size=0,
)

app.mount("/uploads", StaticFiles(directory=str(UPLOAD_DIR)), name="uploads")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class Record(Base):
    __tablename__ = "records"
    id = Column(Integer, primary_key=True, index=True)
    musteri = Column(String, default="")
    belge_tipi = Column(String, default="")
    firma = Column(String, default="")
    vergi_no = Column(String, default="")
    tarih = Column(String, default="")
    belge_no = Column(String, default="")
    matrah = Column(Float, default=0.0)
    kdv_orani = Column(String, default="")
    kdv_tutari = Column(Float, default=0.0)
    toplam = Column(Float, default=0.0)
    aciklama = Column(String, default="")
    dosya_adi = Column(String, default="")
    dosya_yolu = Column(String, default="")
    durum = Column(String, default="Kontrol")
    created_at = Column(DateTime, default=datetime.utcnow)


Base.metadata.create_all(bind=engine)


def safe_float(value) -> float:
    if value in (None, "", "-"):
        return 0.0
    try:
        text = str(value).replace("TL", "").replace("TRY", "").replace("%", "").strip()
        text = re.sub(r"[^\d,.\-]", "", text)
        if not text or text in {"-", ",", "."}:
            return 0.0

        last_dot = text.rfind(".")
        last_comma = text.rfind(",")

        if last_dot != -1 and last_comma != -1:
            if last_dot > last_comma:
                decimal_sep = "."
                thousands_sep = ","
            else:
                decimal_sep = ","
                thousands_sep = "."
            text = text.replace(thousands_sep, "")
            text = text.replace(decimal_sep, ".")
        elif last_comma != -1:
            digits_after = len(text) - last_comma - 1
            if digits_after == 3 and text.count(",") == 1:
                text = text.replace(",", "")
            else:
                text = text.replace(".", "")
                text = text.replace(",", ".")
        elif last_dot != -1:
            digits_after = len(text) - last_dot - 1
            if digits_after == 3 and text.count(".") == 1:
                text = text.replace(".", "")
            else:
                text = text.replace(",", "")

        return float(text)
    except Exception:
        return 0.0


def file_to_base64(filepath: Path) -> str:
    with open(filepath, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def get_mime_type(filename: str) -> str:
    lower = filename.lower()
    if lower.endswith(".png"):
        return "image/png"
    if lower.endswith(".jpg") or lower.endswith(".jpeg"):
        return "image/jpeg"
    if lower.endswith(".webp"):
        return "image/webp"
    if lower.endswith(".pdf"):
        return "application/pdf"
    return "application/octet-stream"


def build_claude_content(filepath: Path, prompt: str) -> list[dict]:
    mime_type = get_mime_type(filepath.name)
    b64 = file_to_base64(filepath)

    if mime_type == "application/pdf":
        file_block = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": mime_type,
                "data": b64,
            },
        }
    else:
        file_block = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": mime_type,
                "data": b64,
            },
        }

    return [
        file_block,
        {
            "type": "text",
            "text": prompt,
        },
    ]


def extract_with_claude(filepath: Path) -> dict:
    if not CLAUDE_API_KEY:
        raise HTTPException(status_code=500, detail="CLAUDE_API_KEY tanımlı değil.")
    model_name = normalize_model_name(CLAUDE_MODEL)

    prompt = """
Sen Türkiye'de çalışan deneyimli bir mali müşavir yardımcısısın.
Kullanıcının yüklediği belgeyi analiz et ve sadece aşağıdaki JSON yapısında cevap ver:
{
  "belge_tipi": "",
  "firma": "",
  "vergi_no": "",
  "tarih": "",
  "belge_no": "",
  "matrah": "",
  "kdv_orani": "",
  "kdv_tutari": "",
  "toplam": "",
  "aciklama": ""
}
Kurallar:
- Çıktıyı sadece JSON olarak ver
- Tarih formatı YYYY-MM-DD olsun
- Emin olmadığın alanları boş bırak
- Türkiye belgelerine göre yorumla
- Belge tipi şu seçeneklerden biri olsun:
  "fiş", "alış faturası", "satış faturası", "e-Arşiv", "e-Fatura", "serbest meslek makbuzu", "gider pusulası", "diğer"
- Sayısal alanlarda sadece sayı döndür
"""
    payload = {
        "model": model_name,
        "max_tokens": 1024,
        "messages": [
            {
                "role": "user",
                "content": build_claude_content(filepath, prompt),
            }
        ],
    }
    headers = {
        "x-api-key": CLAUDE_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers=headers,
        json=payload,
        timeout=120,
    )
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        detail = response.text[:500] if response.text else "AI isteği başarısız oldu."
        raise HTTPException(status_code=502, detail=detail) from exc

    data = response.json()
    text_output = ""
    for content in data.get("content", []):
        if content.get("type") == "text":
            text_output += content.get("text", "")
    if not text_output.strip():
        raise HTTPException(status_code=500, detail="AI cevabı alınamadı.")
    try:
        parsed = json.loads(text_output)
    except json.JSONDecodeError:
        cleaned = text_output.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        parsed = json.loads(cleaned)
    return {
        "belge_tipi": parsed.get("belge_tipi", ""),
        "firma": parsed.get("firma", ""),
        "vergi_no": parsed.get("vergi_no", ""),
        "tarih": parsed.get("tarih", ""),
        "belge_no": parsed.get("belge_no", ""),
        "matrah": parsed.get("matrah", ""),
        "kdv_orani": parsed.get("kdv_orani", ""),
        "kdv_tutari": parsed.get("kdv_tutari", ""),
        "toplam": parsed.get("toplam", ""),
        "aciklama": parsed.get("aciklama", ""),
    }


@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    template = jinja_env.get_template("index.html")
    return HTMLResponse(template.render(request=request))


@app.post("/upload", response_class=HTMLResponse)
async def upload_document(
    request: Request,
    musteri: str = Form(...),
    file: UploadFile = File(...),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Dosya seçilmedi.")

    mime_type = get_mime_type(file.filename)
    if mime_type == "application/octet-stream":
        raise HTTPException(status_code=400, detail="Sadece resim dosyaları ve PDF yükleyebilirsiniz.")

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    safe_name = f"{timestamp}_{file.filename.replace(' ', '_')}"
    file_path = UPLOAD_DIR / safe_name

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        ai_result = extract_with_claude(file_path)
    except Exception:
        if file_path.exists():
            file_path.unlink()
        raise

    context = {
        "request": request,
        "musteri": musteri,
        "dosya_adi": safe_name,
        "dosya_yolu": str(file_path),
        "result": ai_result,
        "is_pdf": mime_type == "application/pdf",
    }
    template = jinja_env.get_template("review.html")
    return HTMLResponse(template.render(**context))


@app.post("/save")
async def save_record(
    musteri: str = Form(...),
    belge_tipi: str = Form(""),
    firma: str = Form(""),
    vergi_no: str = Form(""),
    tarih: str = Form(""),
    belge_no: str = Form(""),
    matrah: str = Form(""),
    kdv_orani: str = Form(""),
    kdv_tutari: str = Form(""),
    toplam: str = Form(""),
    aciklama: str = Form(""),
    dosya_adi: str = Form(""),
    dosya_yolu: str = Form(""),
    durum: str = Form("Onaylandı"),
):
    db = SessionLocal()
    try:
        record = Record(
            musteri=musteri,
            belge_tipi=belge_tipi,
            firma=firma,
            vergi_no=vergi_no,
            tarih=tarih,
            belge_no=belge_no,
            matrah=safe_float(matrah),
            kdv_orani=kdv_orani,
            kdv_tutari=safe_float(kdv_tutari),
            toplam=safe_float(toplam),
            aciklama=aciklama,
            dosya_adi=dosya_adi,
            dosya_yolu=dosya_yolu,
            durum=durum,
        )
        db.add(record)
        db.commit()
    finally:
        db.close()
    return RedirectResponse(url="/records", status_code=303)


@app.get("/records", response_class=HTMLResponse)
def list_records(request: Request):
    db = SessionLocal()
    try:
        records = db.query(Record).order_by(Record.created_at.desc()).all()
    finally:
        db.close()
    template = jinja_env.get_template("records.html")
    return HTMLResponse(template.render(request=request, records=records))


@app.get("/export")
def export_excel():
    db = SessionLocal()
    try:
        records = db.query(Record).order_by(Record.created_at.desc()).all()
    finally:
        db.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Evraklar"
    headers = [
        "Müşteri",
        "Belge Türü",
        "Firma",
        "Vergi No",
        "Tarih",
        "Belge No",
        "Matrah",
        "KDV Oranı",
        "KDV Tutarı",
        "Toplam",
        "Açıklama",
        "Dosya Adı",
        "Durum",
        "Kayıt Tarihi",
    ]
    ws.append(headers)
    for r in records:
        ws.append(
            [
                r.musteri,
                r.belge_tipi,
                r.firma,
                r.vergi_no,
                r.tarih,
                r.belge_no,
                r.matrah,
                r.kdv_orani,
                r.kdv_tutari,
                r.toplam,
                r.aciklama,
                r.dosya_adi,
                r.durum,
                r.created_at.strftime("%Y-%m-%d %H:%M"),
            ]
        )
    export_path = EXPORT_DIR / f"evraklar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(export_path)
    return FileResponse(
        path=export_path,
        filename=export_path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
