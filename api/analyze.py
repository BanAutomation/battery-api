# api/analyze.py
import os, io, json, base64, tempfile, requests
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import pandas as pd

import matplotlib
matplotlib.use("Agg")  # headless backend for serverless

# Later you'll import your real pipeline:
# from batt4 import run_pipeline_from_df

app = FastAPI()

# Lock CORS to your Bubble domain(s)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://your-bubble-app.bubbleapps.io", "https://yourcustomdomain.com"],
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

def _store_bytes(filename: str, content: bytes, content_type: str) -> str:
    store_url = os.environ.get("STORE_URL", "http://localhost:3000/api/store")
    r = requests.post(
        store_url,
        headers={"Content-Type": "application/json"},
        data=json.dumps({
            "filename": filename,
            "content_type": content_type,
            "data_base64": base64.b64encode(content).decode("ascii")
        }),
        timeout=60
    )
    r.raise_for_status()
    return r.json()["url"]

@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...), sheet_name: str = Form("Sheet")):
    # 1) Read Excel
    try:
        excel_bytes = await file.read()
        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name)
    except Exception as e:
        return JSONResponse({"error": f"Failed to read Excel: {e}"}, status_code=400)

    # 2) TEMP MOCK: write tiny CSV/PDF so the endpoint works now.
    #    Replace this block with: csv_path, pdf_path = run_pipeline_from_df(df)
    import csv, pathlib
    tmp = tempfile.gettempdir()
    csv_path = os.path.join(tmp, "result.csv")
    pdf_path = os.path.join(tmp, "report.pdf")
    with open(csv_path, "w", newline="") as f:
        w = csv.writer(f); w.writerow(["ok"]); w.writerow(["pipeline stub"])
    with open(pdf_path, "wb") as f:
        f.write(b"%PDF-1.4\n% stub pdf\n")  # placeholder

    # 3) Upload to storage and return URLs
    with open(csv_path, "rb") as f: csv_bytes = f.read()
    with open(pdf_path, "rb") as f: pdf_bytes = f.read()
    csv_url = _store_bytes("result.csv", csv_bytes, "text/csv")
    pdf_url = _store_bytes("report.pdf", pdf_bytes, "application/pdf")
    return JSONResponse({"csv_url": csv_url, "pdf_url": pdf_url})