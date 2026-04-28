from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse
import uvicorn
import io
import re
from datetime import datetime

from extractor import extract_fields
from filler import fill_rebill_sheet
from data_loader import lookup_unit, lookup_customer, lookup_owner, lookup_email

app = FastAPI(title="Bill Back Generator")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/api/debug")
async def debug():
    import os
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    return {
        "key_present": bool(key),
        "key_length": len(key),
        "key_prefix": key[:14] if key else None,
        "has_spaces": " " in key,
        "has_quotes": '"' in key or "'" in key,
    }


@app.get("/", response_class=HTMLResponse)
async def root():
    with open("static/index.html") as f:
        return f.read()


@app.post("/api/extract")
async def extract(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are supported.")
    pdf_bytes = await file.read()
    try:
        fields = await extract_fields(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Extraction error: {e}")

    unit = fields.get("unit") or ""

    # Lease/Rental from unit lookup
    try:
        fields["lease_or_rental"] = lookup_unit(unit)
    except Exception:
        fields["lease_or_rental"] = "Lease"

    # Customer name + number from Customer Numbers spreadsheet
    try:
        cust_num, cust_name = lookup_customer(unit)
        if cust_num and cust_name:
            fields["customer"] = f"{cust_num} {cust_name}"
        elif cust_name:
            fields["customer"] = cust_name
        elif cust_num:
            fields["customer"] = cust_num
        else:
            fields["customer"] = ""
    except Exception:
        fields["customer"] = ""

    # Owner and email — append contact line to invoice wording
    try:
        owner = lookup_owner(unit)
        email = lookup_email(unit, owner)
        if owner or email:
            wording = fields.get("invoice_wording") or ""
            contact = f"Questions, contact {owner}" if owner else "Questions, contact"
            if email:
                contact += f" at {email}"
            fields["invoice_wording"] = f"{wording}  {contact}"
    except Exception:
        pass

    return fields


@app.post("/api/generate")
async def generate(body: dict):
    output = fill_rebill_sheet(body)

    unit     = re.sub(r"[^\w\s-]", "", body.get("unit") or "").strip()
    customer = re.sub(r"[^\w\s-]", "", body.get("customer") or "").strip()
    now      = datetime.now()
    date_str = f"{now.month}.{now.day}.{str(now.year)[2:]}"

    parts    = [p for p in [unit, customer, date_str] if p]
    filename = " ".join(parts) + ".xlsm"

    return StreamingResponse(
        io.BytesIO(output),
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
