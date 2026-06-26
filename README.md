# Kingdom Legacy — Backend de Diagnóstico Financiero

## ¿Qué hace este sistema?

Recibe archivos de clientes (PDF del SII y Excel de ventas), los procesa automáticamente y retorna datos estructurados que alimentan el dashboard financiero de Kingdom Legacy.

---

## Archivos incluidos

```
kl_backend/
├── main.py              ← API principal (FastAPI)
├── requirements.txt     ← Dependencias Python
└── README.md            ← Este archivo
```

---

## PASO 1 — Instalar en tu computador (para probar localmente)

### Requisitos previos
- Python 3.10 o superior → https://python.org/downloads
- Tesseract OCR → https://github.com/UB-Mannheim/tesseract/wiki (Windows) o `brew install tesseract` (Mac)
- Poppler → https://poppler.freedesktop.org (para pdftoppm)

### Comandos

```bash
# 1. Crear entorno virtual
python -m venv venv
source venv/bin/activate        # Mac/Linux
venv\Scripts\activate           # Windows

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Correr el servidor
uvicorn main:app --reload --port 8000
```

Abre en tu navegador: http://localhost:8000

---

## PASO 2 — Subir a producción gratuita en Render.com

### ¿Por qué Render?
- Gratuito para empezar
- No necesitas saber de servidores
- Se conecta directo a tu código

### Pasos en Render

1. **Crear cuenta** en https://render.com (es gratis)

2. **Subir el código a GitHub**
   - Crea una cuenta en https://github.com si no tienes
   - Crea un repositorio nuevo llamado `kl-backend`
   - Sube los archivos: `main.py` y `requirements.txt`

3. **Crear un Web Service en Render**
   - Clic en "New +" → "Web Service"
   - Conecta tu repositorio de GitHub
   - Configura:
     - **Name:** kl-diagnostico
     - **Runtime:** Python 3
     - **Build Command:** `pip install -r requirements.txt`
     - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - Clic en "Create Web Service"

4. **En 2–3 minutos** tendrás una URL como:
   `https://kl-diagnostico.onrender.com`

---

## PASO 3 — Conectar el dashboard

En el archivo `KL_DiagnosticoFinanciero.html`, actualiza la variable de URL:

```javascript
const API_URL = "https://kl-diagnostico.onrender.com";
```

---

## Endpoints disponibles

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| GET | `/` | Info del sistema |
| GET | `/health` | Verificar que está corriendo |
| POST | `/upload/f29` | Subir PDF Carpeta Tributaria SII |
| POST | `/upload/ventas` | Subir Excel de ventas |

---

## Nota sobre Tesseract en Render

En Render, Tesseract se instala automáticamente agregando esto al inicio del build:

**Build Command:**
```
apt-get install -y tesseract-ocr poppler-utils && pip install -r requirements.txt
```

---

## Soporte

Sistema desarrollado para Kingdom Legacy.
Para dudas de deployment, consultar con el equipo técnico.
