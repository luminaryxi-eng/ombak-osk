"""
OSK Ombak - Document Automation Core
Fills all lawyer forms, DMC, and bank templates.
"""
import subprocess, sys, json, re, base64, shutil
from pathlib import Path

# ── Paths (adjusted for Colab environment) ──────────────────────────────────
SCRIPTS_DIR  = Path("/content/scripts/office")
TEMPLATES_DIR = Path("/content/templates")        # all template .docx files
BANK_DIR     = TEMPLATES_DIR / "bank"             # bank-specific templates
LAWYER_DIR   = TEMPLATES_DIR / "lawyer"           # 02-08 + DMC
OUT_DIR      = Path("/content/outputs")

def ensure_dirs():
    for d in [TEMPLATES_DIR, BANK_DIR, LAWYER_DIR, OUT_DIR]:
        d.mkdir(parents=True, exist_ok=True)

def unpack(src, work):
    subprocess.run(["python", str(SCRIPTS_DIR/"unpack.py"), str(src), str(work)],
                   check=True, capture_output=True)

def pack(work, dst, orig):
    subprocess.run(["python", str(SCRIPTS_DIR/"pack.py"), str(work), str(dst),
                    "--original", str(orig)], check=True, capture_output=True)

def read(p):     return p.read_text(encoding="utf-8")
def write(p, t): p.write_text(t, encoding="utf-8")

def replace_all(path, repl):
    txt = read(path)
    for old, new in repl.items():
        txt = txt.replace(old, new)
    write(path, txt)

def remove_highlights(path):
    txt = read(path)
    txt = txt.replace('<w:highlight w:val="yellow"/>', '')
    write(path, txt)

def xml_esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def arial_para(text, sz=20):
    return (f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{xml_esc(text)}</w:t></w:r></w:p>')

def inject_into_cell(txt, para_id, value):
    escaped_val = xml_esc(value)
    pattern = re.compile(
        r'(<w:p [^>]*w14:paraId="' + re.escape(para_id) + r'"[^>]*>)(.*?)(</w:p>)',
        re.DOTALL)
    def replacer(m):
        inner = m.group(2)
        inner = inner.replace('<w:highlight w:val="yellow"/>', '')
        inner = re.sub(r'<w:r>.*?<w:t[^>]*>\s*</w:t>\s*</w:r>', '', inner, flags=re.DOTALL)
        run = (f'<w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/>'
               f'</w:rPr><w:t xml:space="preserve">{escaped_val}</w:t></w:r>')
        return m.group(1) + inner + run + m.group(3)
    return pattern.sub(replacer, txt, count=1)


# ═══════════════════════════════════════════════════════════════════════════
# CLAUDE API  — vision extraction + generic template filling
# ═══════════════════════════════════════════════════════════════════════════

def call_claude(messages, system=None, max_tokens=1500):
    """Call Claude API via subprocess (works in Colab without anthropic package)."""
    import urllib.request, json as _json
    payload = {
        "model": "claude-opus-4-5",
        "max_tokens": max_tokens,
        "messages": messages,
    }
    if system:
        payload["system"] = system

    api_key = Path("/content/api_key.txt").read_text().strip()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=_json.dumps(payload).encode(),
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        return _json.loads(resp.read())["content"][0]["text"]


def extract_from_image(image_path: Path, doc_type: str) -> dict:
    """
    Use Claude vision to extract structured data from an uploaded image.
    doc_type: 'nric' | 'sales_form' | 'auto'
    Returns dict with extracted fields.
    """
    img_bytes = image_path.read_bytes()
    b64 = base64.b64encode(img_bytes).decode()
    suffix = image_path.suffix.lower()
    media_type = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".pdf": "application/pdf",
        ".webp": "image/webp"
    }.get(suffix, "image/jpeg")

    if doc_type == "nric":
        prompt = """This is a Malaysian MyKad (NRIC). Extract ONLY these fields and return as JSON:
{
  "name": "full name as printed in ALL CAPS",
  "nric": "12-digit IC number formatted as XXXXXX-XX-XXXX",
  "address": "full address as printed, single line",
  "dob": "date of birth if visible",
  "gender": "LELAKI or PEREMPUAN if visible"
}
Return ONLY the JSON object. Use empty string if field not visible."""

    elif doc_type == "sales_form":
        prompt = """This is a Malaysian property sales form / booking form / OTP.
Extract ALL available fields and return as JSON:
{
  "unit_no": "unit number e.g. A-24-10",
  "parcel_no": "parcel/lot number",
  "type_unit": "unit type e.g. A(M), B, C",
  "floor": "floor level",
  "tower": "block/tower e.g. Block A",
  "price": "SPA price as number only e.g. 318500",
  "purchasers": [
    {"name": "full name", "nric": "formatted IC", "address": "address"}
  ],
  "developer": "developer name",
  "project": "project name",
  "date": "sales date if visible"
}
Return ONLY the JSON. Use empty string or empty array if not visible."""

    else:  # auto-detect
        prompt = """Look at this document and determine what type it is, then extract all relevant fields.
Return JSON in this format:
{
  "doc_type": "nric | sales_form | bank_form | other",
  "unit_no": "",
  "parcel_no": "",
  "type_unit": "",
  "purchasers": [{"name": "", "nric": "", "address": ""}],
  "borrowers": [{"name": "", "nric": "", "address": ""}],
  "loan_amount": "",
  "bank_name": "",
  "date": "",
  "raw_text": "key text from document"
}
Return ONLY the JSON."""

    response = call_claude([{
        "role": "user",
        "content": [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
            {"type": "text", "text": prompt}
        ]
    }])

    match = re.search(r'\{[\s\S]*\}', response)
    if match:
        return json.loads(match[0])
    return {}


# ═══════════════════════════════════════════════════════════════════════════
# LAWYER FORMS (02–08 + DMC)
# ═══════════════════════════════════════════════════════════════════════════

def fill_lawyer_forms(data: dict) -> list[str]:
    """Fill all standard lawyer forms. Returns list of output file paths."""
    from fill_lawyers import fill_forms
    return fill_forms(data)


# ═══════════════════════════════════════════════════════════════════════════
# BANK TEMPLATE ENGINE — generic placeholder-based filling
# ═══════════════════════════════════════════════════════════════════════════

# Placeholder convention for bank templates:
# {{PURCHASER_NAME_1}}  {{PURCHASER_NRIC_1}}
# {{PURCHASER_NAME_2}}  {{PURCHASER_NRIC_2}}  (joint purchaser)
# {{BORROWER_NAME_1}}   {{BORROWER_NRIC_1}}
# {{BORROWER_NAME_2}}   {{BORROWER_NRIC_2}}   (second borrower)
# {{BORROWER_ADDRESS_1}} {{BORROWER_ADDRESS_2}}
# {{UNIT_NO}}  {{PARCEL_NO}}  {{PRICE}}  {{LOAN_AMOUNT}}
# {{DATE}}     {{BANK_NAME}}  {{PROJECT_NAME}}

PLACEHOLDER_MAP = {
    # Purchasers
    "{{PURCHASER_NAME_1}}":    lambda d: xml_esc(d.get("purchasers", [{}])[0].get("name", "")),
    "{{PURCHASER_NRIC_1}}":    lambda d: xml_esc(d.get("purchasers", [{}])[0].get("nric", "")),
    "{{PURCHASER_ADDRESS_1}}": lambda d: xml_esc(d.get("purchasers", [{}])[0].get("address", "")),
    "{{PURCHASER_NAME_2}}":    lambda d: xml_esc(d.get("purchasers", [{}, {}])[1].get("name", "") if len(d.get("purchasers", [])) > 1 else ""),
    "{{PURCHASER_NRIC_2}}":    lambda d: xml_esc(d.get("purchasers", [{}, {}])[1].get("nric", "") if len(d.get("purchasers", [])) > 1 else ""),
    "{{PURCHASER_ADDRESS_2}}": lambda d: xml_esc(d.get("purchasers", [{}, {}])[1].get("address", "") if len(d.get("purchasers", [])) > 1 else ""),
    # Combined purchaser (all names / NRICs joined)
    "{{PURCHASER_NAMES_ALL}}": lambda d: xml_esc(" & ".join(p.get("name","") for p in d.get("purchasers", []))),
    "{{PURCHASER_NRICS_ALL}}": lambda d: xml_esc(" / ".join(p.get("nric","") for p in d.get("purchasers", []))),
    # Borrowers (may differ from purchasers)
    "{{BORROWER_NAME_1}}":     lambda d: xml_esc(d.get("borrowers", [{}])[0].get("name", "")),
    "{{BORROWER_NRIC_1}}":     lambda d: xml_esc(d.get("borrowers", [{}])[0].get("nric", "")),
    "{{BORROWER_ADDRESS_1}}":  lambda d: xml_esc(d.get("borrowers", [{}])[0].get("address", "")),
    "{{BORROWER_NAME_2}}":     lambda d: xml_esc(d.get("borrowers", [{}, {}])[1].get("name", "") if len(d.get("borrowers", [])) > 1 else ""),
    "{{BORROWER_NRIC_2}}":     lambda d: xml_esc(d.get("borrowers", [{}, {}])[1].get("nric", "") if len(d.get("borrowers", [])) > 1 else ""),
    "{{BORROWER_ADDRESS_2}}":  lambda d: xml_esc(d.get("borrowers", [{}, {}])[1].get("address", "") if len(d.get("borrowers", [])) > 1 else ""),
    "{{BORROWER_NAMES_ALL}}":  lambda d: xml_esc(" & ".join(b.get("name","") for b in d.get("borrowers", []))),
    "{{BORROWER_NRICS_ALL}}":  lambda d: xml_esc(" / ".join(b.get("nric","") for b in d.get("borrowers", []))),
    # Unit / property
    "{{UNIT_NO}}":             lambda d: xml_esc(d.get("unit_no", "")),
    "{{PARCEL_NO}}":           lambda d: xml_esc(d.get("parcel_no", "")),
    "{{TYPE_UNIT}}":           lambda d: xml_esc(d.get("type_unit", "")),
    "{{PRICE}}":               lambda d: xml_esc(str(d.get("price", ""))),
    "{{LOAN_AMOUNT}}":         lambda d: xml_esc(str(d.get("loan_amount", ""))),
    "{{DATE}}":                lambda d: xml_esc(d.get("date", "")),
    "{{BANK_NAME}}":           lambda d: xml_esc(d.get("bank_name", "")),
    "{{PROJECT_NAME}}":        lambda d: xml_esc(d.get("project_name", "OSK OMBAK")),
    "{{DEVELOPER}}":           lambda d: xml_esc(d.get("developer", "PJD Sejahtera Sdn. Bhd.")),
    # Convenience: first borrower or first purchaser (whichever exists)
    "{{PRIMARY_NAME}}":        lambda d: xml_esc((d.get("borrowers") or d.get("purchasers") or [{}])[0].get("name", "")),
    "{{PRIMARY_NRIC}}":        lambda d: xml_esc((d.get("borrowers") or d.get("purchasers") or [{}])[0].get("nric", "")),
}

def fill_bank_template(template_path: Path, data: dict) -> str:
    """
    Fill a bank template using {{PLACEHOLDER}} convention.
    Returns output file path.
    """
    work = Path(f"/tmp/work_bank_{template_path.stem}")
    work.mkdir(parents=True, exist_ok=True)
    unpack(template_path, work)

    for xml_file in work.rglob("*.xml"):
        txt = read(xml_file)
        changed = False
        for placeholder, resolver in PLACEHOLDER_MAP.items():
            if placeholder in txt:
                txt = txt.replace(placeholder, resolver(data))
                changed = True
        if changed:
            write(xml_file, txt)
        remove_highlights(xml_file)

    suffix = data.get("unit_no", "output").replace("/","_").replace(" ","_")
    bank_name = data.get("bank_name", template_path.stem).replace(" ","_")
    out = OUT_DIR / f"BANK_{bank_name}_{suffix}.docx"
    pack(work, out, template_path)
    return str(out)


def fill_all_bank_templates(data: dict) -> list[str]:
    """Fill all templates found in the bank templates directory."""
    results = []
    bank_templates = list(BANK_DIR.glob("*.docx"))
    if not bank_templates:
        print("  No bank templates found in /content/templates/bank/")
        return results
    for tmpl in sorted(bank_templates):
        print(f"  Filling bank template: {tmpl.name}")
        try:
            out = fill_bank_template(tmpl, data)
            results.append(out)
        except Exception as e:
            print(f"  ⚠ Error filling {tmpl.name}: {e}")
    return results


# ═══════════════════════════════════════════════════════════════════════════
# MAIN RUNNER
# ═══════════════════════════════════════════════════════════════════════════

def run_all(data: dict, fill_lawyer=True, fill_bank=True) -> dict:
    """
    Master function. Fill all applicable forms given a data dict.
    Returns {"lawyer": [...], "bank": [...]}
    """
    ensure_dirs()
    results = {"lawyer": [], "bank": []}

    if fill_lawyer:
        print("📄 Filling lawyer forms...")
        try:
            results["lawyer"] = fill_lawyer_forms(data)
            print(f"  ✓ {len(results['lawyer'])} lawyer forms generated")
        except Exception as e:
            print(f"  ✗ Lawyer forms error: {e}")

    if fill_bank:
        print("🏦 Filling bank templates...")
        results["bank"] = fill_all_bank_templates(data)
        print(f"  ✓ {len(results['bank'])} bank templates generated")

    return results
