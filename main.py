"""
Kingdom Legacy — Backend de Procesamiento Financiero
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, HTMLResponse
import tempfile, os, re
from typing import Optional
import openpyxl
import pdfplumber

app = FastAPI(
    title="Kingdom Legacy — API Diagnóstico Financiero",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

MONTH_MAP = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12
}

def clean_num(s):
    try: return int(re.sub(r'[.,\s]','',str(s)))
    except: return 0

def detect_period(text):
    m = re.search(r'PERIODO[^\d]*(\d{6})', text, re.IGNORECASE)
    if m:
        raw = m.group(1)
        return f"{raw[:4]}-{raw[4:]}"
    for name, num in MONTH_MAP.items():
        m2 = re.search(rf'\b{name}\b[^\d]*(\d{{4}})', text, re.IGNORECASE)
        if m2:
            return f"{m2.group(1)}-{str(num).zfill(2)}"
    return None

def parse_f29_fields(text):
    def find(patterns):
        for p in patterns:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                return clean_num(m.group(1))
        return 0
    return {
        "base_imponible": find([r'BASE\s+IMPONIBLE\s+([\d.,]+)', r'\b563\b[^\d]+([\d.,]+)']),
        "debitos_iva":    find([r'DEBITOS\s+FACTURAS\s+EMITIDAS\s+([\d.,]+)', r'\b336\b[^\d]+([\d.,]+)']),
        "creditos_iva":   find([r'TOTAL\s+CR[EÉ]DITOS?\s+([\d.,]+)', r'\b337\b[^\d]+([\d.,]+)']),
        "iva_pagar":      find([r'TOTAL\s+A\s+PAGAR\s+DENTRO\s+DEL\s+PLAZO\s+LEGAL\s+([\d.,]+)',
                                r'TOTAL\s+DETERMINADO\s+([\d.,]+)']),
        "ppm":            find([r'PPM\s+NETO\s+DETERMINADO\s+([\d.,]+)', r'\b052\b[^\d]+([\d.,]+)']),
    }

@app.get("/", response_class=HTMLResponse)
async def dashboard():
    """Sirve el dashboard de Kingdom Legacy."""
    dashboard_path = os.path.join(os.path.dirname(__file__), "dashboard.html")
    if os.path.exists(dashboard_path):
        with open(dashboard_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>Kingdom Legacy API</h1><p>Dashboard no encontrado.</p>")

@app.get("/api")
def api_info():
    return {
        "sistema": "Kingdom Legacy — API Diagnóstico Financiero",
        "version": "1.0.0",
        "endpoints": {
            "POST /upload/f29": "Sube PDF de Carpeta Tributaria SII",
            "POST /upload/ventas": "Sube Excel de libro de ventas",
        }
    }

@app.post("/upload/f29")
async def upload_f29(file: UploadFile = File(...)):
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "Solo se aceptan archivos PDF")
    tmp = tempfile.mktemp(suffix='.pdf')
    try:
        with open(tmp, 'wb') as f:
            f.write(await file.read())
        resultados = []
        with pdfplumber.open(tmp) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if 'FORMULARIO 29' not in text.upper():
                    continue
                period = detect_period(text)
                if not period:
                    continue
                fields = parse_f29_fields(text)
                resultados.append({"period": period, **fields})
        resultados.sort(key=lambda x: x['period'])
        return {"ok": True, "archivo": file.filename,
                "periodos_extraidos": len(resultados), "datos": resultados}
    finally:
        try: os.remove(tmp)
        except: pass

@app.post("/upload/ventas")
async def upload_ventas(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.xlsx','.xls')):
        raise HTTPException(400, "Solo se aceptan archivos Excel")
    tmp = tempfile.mktemp(suffix='.xlsx')
    try:
        with open(tmp, 'wb') as f:
            f.write(await file.read())
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        resultado = {"archivo": file.filename, "hojas": wb.sheetnames, "resumen": {}}
        for name in wb.sheetnames:
            if 'VENTA' in name.upper():
                ws = wb[name]
                for row in ws.iter_rows(max_row=8, values_only=True):
                    if not row: continue
                    for i, cell in enumerate(row):
                        if not cell or not isinstance(cell, str): continue
                        val = row[i+1] if i+1 < len(row) else None
                        if 'VENTAS NETAS' in str(cell).upper() and val:
                            resultado['resumen']['ventas_netas'] = round(float(val),0)
                        elif 'COSTO DE VENTA' in str(cell).upper() and val:
                            resultado['resumen']['costo_venta'] = round(float(val),0)
                        elif '% RENTABILIDAD' in str(cell).upper() and val:
                            resultado['resumen']['rentabilidad_pct'] = round(float(val)*100,2)
                fechas = []
                for row in ws.iter_rows(min_row=9, max_row=50, values_only=True):
                    if row and row[1] and hasattr(row[1],'year'):
                        fechas.append(row[1])
                if fechas:
                    resultado['resumen']['period'] = f"{fechas[0].year}-{str(fechas[0].month).zfill(2)}"
                break
        return {"ok": True, **resultado}
    finally:
        try: os.remove(tmp)
        except: pass

@app.get("/health")
def health():
    return {"status": "ok"}
