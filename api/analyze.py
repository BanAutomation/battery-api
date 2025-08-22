import os, io, json, base64, tempfile, requests
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import pandas as pd
import matplotlib
matplotlib.use("Agg")

from batt4 import run_pipeline_from_df  # uses the entrypoint you just added

app = FastAPI()

# Allow Bubble to call you from the browser
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://your-bubble-app.bubbleapps.io",   # <- replace with your Bubble app domain(s)
        "https://yourcustomdomain.com"
    ],
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

def _store_bytes(filename: str, content: bytes, content_type: str) -> str:
    store_url = os.environ.get("STORE_URL")
    if not store_url:
        raise RuntimeError("STORE_URL env var not set (should be https://<your-app>.vercel.app/api/store)")

    try:
        r = requests.post(
            store_url,
            headers={"Content-Type": "application/json"},
            data=json.dumps({
                "filename": filename,
                "content_type": content_type,
                "data_base64": base64.b64encode(content).decode("ascii")
            }),
            timeout=10  # fail fast instead of hanging 60s
        )
        r.raise_for_status()
        j = r.json()
        if "url" not in j:
            raise RuntimeError(f"Store did not return a URL: {j}")
        return j["url"]
    except requests.Timeout:
        raise RuntimeError("Timed out contacting /api/store (check STORE_URL or store route health).")
    except requests.RequestException as e:
        raise RuntimeError(f"Upload to store failed: {e}")

@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...), sheet_name: str = Form("Sheet")):
    try:
        print("A) receiving upload…")
        excel_bytes = await file.read()
        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name)
        print("B) parsed excel:", len(df), "rows")
    except Exception as e:
        return JSONResponse({"error": f"Failed to read Excel: {e}"}, status_code=400)

    try:
        print("C) running pipeline…")
        csv_path, pdf_path = run_pipeline_from_df(df)
        print("D) pipeline done:", csv_path, pdf_path)

        with open(csv_path, "rb") as f:
            csv_bytes = f.read()
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        print("E) uploading CSV to store…")
        csv_url = _store_bytes("result.csv", csv_bytes, "text/csv")
        print("F) uploading PDF to store…")
        pdf_url = _store_bytes("report.pdf", pdf_bytes, "application/pdf")

        print("G) done")
        return JSONResponse({"csv_url": csv_url, "pdf_url": pdf_url})
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)
