import io
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
import pdfplumber
import openpyxl

app = FastAPI(title="Custom Doc Converter API")

@app.get("/")
def home():
    return {"status": "online", "message": "Doc Converter API Calisiyor"}

@app.post("/convert/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Lutfen sadece PDF dosyasi yukleyin.")

    pdf_bytes = await file.read()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Donusturulen Tablo"

    current_row = 1

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row_data in table:
                    ws.append([cell if cell is not None else "" for cell in row_data])
                    current_row += 1
                current_row += 1  # Tablolar arasi 1 satir bosluk

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        headers={"Content-Disposition": f"attachment; filename=donusturulen_{file.filename.replace('.pdf', '.xlsx')}"},
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
