"""Extracts repair order fields using Claude AI."""
import base64
import json
import anthropic
import os
from dotenv import load_dotenv

load_dotenv()

SYSTEM_PROMPT = """You are an expert at reading vehicle repair orders for a fleet leasing company.
Extract structured data to populate a bill-back rebill sheet. Return ONLY valid JSON with these exact keys.

For "vendor": the shop/dealer name (if Jackson Group Peterbilt, use only the city name, e.g. "Elko").
For "date": use YYYY-MM-DD format.
For "line_items": extract EVERY individual labor and parts charge as a separate item. For each:
  - "description": brief description of the work or part
  - "cost": the dollar amount (number only, no $ or commas). If a line has both parts and labor, split them into separate items.
  - "type": classify as "pm" (preventive maintenance — oil changes, filters, fluid services, greasing, A/B service, inspections, PMs) OR "repair" (fixing broken/worn components, replacements due to failure, diagnostics)
  - "category": classify as "parts" (physical parts, materials, supplies) OR "labor" (labor charges, labor hours, installation, diagnostic time)

Ignore courtesy inspections and campaigns entirely.

For "invoice_wording": write a brief rebill description in this format:
"[unit] - BB for [short plain-English summary of the main work done]"
Example: "1234 - BB for A service and brake repair"

For "misc": the dollar amount of any misc/shop supplies, environmental fees, or supply charges listed separately on the RO. Use null if none found.

{
  "vendor": "shop name",
  "invoice_number": "invoice or RO number",
  "date": "YYYY-MM-DD",
  "unit": "vehicle unit number",
  "invoice_wording": "unit - BB for ...",
  "misc": 12.50,
  "line_items": [
    {"description": "Oil & Filter Change Labor", "cost": 45.00, "type": "pm", "category": "labor"},
    {"description": "Oil Filter", "cost": 40.00, "type": "pm", "category": "parts"},
    {"description": "Front brake pads", "cost": 210.50, "type": "repair", "category": "parts"},
    {"description": "Brake Labor", "cost": 120.00, "type": "repair", "category": "labor"}
  ]
}"""


async def extract_fields(pdf_bytes: bytes) -> dict:
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", "").strip())
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64,
                    },
                },
                {"type": "text", "text": "Extract the repair order fields and return as JSON."},
            ],
        }],
    )

    raw = message.content[0].text.strip()

    # Extract JSON block — find outermost { ... }
    start = raw.find("{")
    end   = raw.rfind("}")
    if start != -1 and end != -1:
        raw = raw[start:end + 1]

    return json.loads(raw)
