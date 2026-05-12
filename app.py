"""
FastAPI backend for the GSRTC LED Screen proposal generator.
Serves the frontend UI and provides endpoints to generate PPT + Excel files.
"""

from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, Response, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List
from pathlib import Path
import re

from generator import (
    generate_pptx, generate_xlsx, sanitize_filename,
    LOCATIONS_IN_ORDER, LOCATION_EXCEL_ROW
)
from openpyxl import load_workbook

BASE_DIR = Path(__file__).parent
app = FastAPI(title="GSRTC Proposal Generator")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")


def get_location_data():
    """Read pricing from the rate card so the frontend can show prices."""
    wb = load_workbook(BASE_DIR / "assets" / "ratecard.xlsx", data_only=True)
    ws = wb["Sheet1"]

    locations = []
    for loc in LOCATIONS_IN_ORDER:
        row = LOCATION_EXCEL_ROW[loc]
        locations.append({
            "name": loc,
            "screens": int(ws.cell(row=row, column=5).value or 0),
            "unit_inch": str(ws.cell(row=row, column=3).value or ""),
            "grade": str(ws.cell(row=row, column=4).value or ""),
            "total_month": int(ws.cell(row=row, column=12).value or 0),
            "special_offer": int(ws.cell(row=row, column=13).value or 0),
        })
    return locations


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    locations = get_location_data()
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={"locations": locations},
    )


class ProposalRequest(BaseModel):
    company: str = ""
    client_name: str
    client_mobile: str = ""
    client_email: str = ""
    sender_name: str = ""
    sender_mobile: str = ""
    sender_email: str = ""
    locations: List[str]


@app.post("/generate/pptx")
async def gen_pptx(req: ProposalRequest):
    if not req.client_name.strip():
        raise HTTPException(400, "Client name is required")
    if not req.locations:
        raise HTTPException(400, "Select at least one location")

    try:
        data = generate_pptx(
            client_name=req.client_name,
            company=req.company,
            mobile=req.client_mobile,
            email=req.client_email,
            selected_locations=req.locations,
            sender_name=req.sender_name,
            sender_mobile=req.sender_mobile,
            sender_email=req.sender_email,
        )
    except Exception as e:
        raise HTTPException(500, f"Failed to generate PPT: {e}")

    base = req.company.strip() or req.client_name
    filename = f"GSRTC_Proposal_{sanitize_filename(base)}.pptx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.post("/generate/xlsx")
async def gen_xlsx(req: ProposalRequest):
    if not req.client_name.strip():
        raise HTTPException(400, "Client name is required")
    if not req.locations:
        raise HTTPException(400, "Select at least one location")

    try:
        data = generate_xlsx(
            client_name=req.client_name,
            company=req.company,
            mobile=req.client_mobile,
            email=req.client_email,
            selected_locations=req.locations,
            sender_name=req.sender_name,
            sender_mobile=req.sender_mobile,
            sender_email=req.sender_email,
        )
    except Exception as e:
        raise HTTPException(500, f"Failed to generate Excel: {e}")

    base = req.company.strip() or req.client_name
    filename = f"GSRTC_RateCard_{sanitize_filename(base)}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.get("/health")
async def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
