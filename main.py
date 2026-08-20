import io
import requests
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
import pdfplumber
import openpyxl

app = FastAPI(title="Custom Doc Converter API")

@app.get("/")
def home():
    return {"status": "online", "message": "Doc Converter API Calisiyor"}

@app.post("/convert/pdf-to-excel")
async def pdf_to_excel(
    file: UploadFile = File(None),
    file_url: str = Form(None)
):
    pdf_bytes = None

    if file:
        pdf_bytes = await file.read()
    elif file_url:
        if file_url.startswith("//"):
            file_url = "https:" + file_url
        res = requests.get(file_url)
        if res.status_code == 200:
            pdf_bytes = res.content
        else:
            raise HTTPException(status_code=400, detail="URL'den dosya indirilemedi.")
    else:
        raise HTTPException(status_code=400, detail="Lutfen dosya veya file_url gonderin.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Donusturulen Tablo"

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row_data in table:
                    ws.append([cell if cell is not None else "" for cell in row_data])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        headers={"Content-Disposition": "attachment; filename=donusturulen_tablo.xlsx"},
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
