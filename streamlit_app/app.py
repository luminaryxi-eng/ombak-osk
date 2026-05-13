"""
OSK Ombak – Document Automation
Streamlit web app: upload NRIC / sales form → download all filled forms as ZIP.
"""
import streamlit as st
import sys, json, re, base64, zipfile, io, os, subprocess, shutil, tempfile
from pathlib import Path
import urllib.request

# ── App config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OSK Ombak – Form Automation",
    page_icon="🏠",
    layout="centered"
)

APP_DIR     = Path(__file__).parent
SCRIPTS_DIR = APP_DIR / "scripts" / "office"
LAWYER_DIR  = APP_DIR / "templates" / "lawyer"
BANK_DIR    = APP_DIR / "templates" / "bank"
OUT_DIR     = APP_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)
BANK_DIR.mkdir(exist_ok=True)

sys.path.insert(0, str(APP_DIR))

# ── Claude API ────────────────────────────────────────────────────────────────
def call_claude(messages, max_tokens=1500):
    api_key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY",""))
    if not api_key:
        st.error("ANTHROPIC_API_KEY not set. Add it to Streamlit secrets.")
        st.stop()
    payload = {
        "model": "claude-opus-4-5",
        "max_tokens": max_tokens,
        "messages": messages
    }
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(payload).encode(),
        headers={"Content-Type":"application/json","x-api-key":api_key,
                 "anthropic-version":"2023-06-01"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())["content"][0]["text"]

def extract_from_image(file_bytes: bytes, filename: str) -> dict:
    """Use Claude vision to extract details from NRIC or sales form."""
    ext = Path(filename).suffix.lower()
    mt = {".jpg":"image/jpeg",".jpeg":"image/jpeg",".png":"image/png",
          ".pdf":"application/pdf",".webp":"image/webp"}.get(ext,"image/jpeg")
    b64 = base64.b64encode(file_bytes).decode()

    prompt = """Examine this document. It may be a Malaysian MyKad (NRIC), property sales form, booking form, or OTP.
Extract ALL available fields and return ONLY a JSON object:
{
  "doc_type": "nric | sales_form | other",
  "name": "if NRIC: person full name in caps",
  "nric": "if NRIC: IC formatted XXXXXX-XX-XXXX",
  "address": "if NRIC or form: full address",
  "unit_no": "unit number e.g. A-24-10",
  "parcel_no": "parcel number",
  "type_unit": "unit type e.g. A(M) or B",
  "price": "SPA price digits only",
  "date": "document date",
  "purchasers": [{"name":"","nric":"","address":""}],
  "bank_name": "",
  "loan_amount": ""
}
Use empty string if a field is not visible. Return ONLY the JSON."""

    resp = call_claude([{"role":"user","content":[
        {"type":"image","source":{"type":"base64","media_type":mt,"data":b64}},
        {"type":"text","text":prompt}
    ]}])
    m = re.search(r'\{[\s\S]*\}', resp)
    return json.loads(m.group()) if m else {}

def merge_extractions(extractions: list) -> dict:
    merged = {"purchasers": [], "borrowers": []}
    for ex in extractions:
        dt = ex.get("doc_type","")
        if dt == "nric":
            p = {"name":ex.get("name",""),"nric":ex.get("nric",""),"address":ex.get("address","")}
            if p["name"] and p not in merged["purchasers"]:
                merged["purchasers"].append(p)
        elif dt == "sales_form":
            for f in ["unit_no","parcel_no","type_unit","price","date"]:
                if ex.get(f) and not merged.get(f):
                    merged[f] = ex[f]
            for p in ex.get("purchasers",[]):
                if p.get("name") and p not in merged["purchasers"]:
                    merged["purchasers"].append(p)
        for k, v in ex.items():
            if k not in ("doc_type","purchasers","borrowers") and v and k not in merged:
                merged[k] = v
    return merged

# ── Bank template filler ──────────────────────────────────────────────────────
def xml_esc(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def unpack(src, work):
    subprocess.run(["python", str(SCRIPTS_DIR/"unpack.py"), str(src), str(work)],
                   check=True, capture_output=True)

def pack(work, dst, orig):
    subprocess.run(["python", str(SCRIPTS_DIR/"pack.py"), str(work), str(dst),
                    "--original", str(orig)], check=True, capture_output=True)

PLACEHOLDERS = {
    "{{PURCHASER_NAME_1}}":    lambda d: xml_esc((d.get("purchasers") or [{}])[0].get("name","")),
    "{{PURCHASER_NRIC_1}}":    lambda d: xml_esc((d.get("purchasers") or [{}])[0].get("nric","")),
    "{{PURCHASER_ADDRESS_1}}": lambda d: xml_esc((d.get("purchasers") or [{}])[0].get("address","")),
    "{{PURCHASER_NAME_2}}":    lambda d: xml_esc((d.get("purchasers") or [{},{}])[1].get("name","") if len(d.get("purchasers",[]))>1 else ""),
    "{{PURCHASER_NRIC_2}}":    lambda d: xml_esc((d.get("purchasers") or [{},{}])[1].get("nric","") if len(d.get("purchasers",[]))>1 else ""),
    "{{PURCHASER_NAMES_ALL}}": lambda d: xml_esc(" & ".join(p.get("name","") for p in d.get("purchasers",[]))),
    "{{PURCHASER_NRICS_ALL}}": lambda d: xml_esc(" / ".join(p.get("nric","") for p in d.get("purchasers",[]))),
    "{{BORROWER_NAME_1}}":     lambda d: xml_esc((d.get("borrowers") or [{}])[0].get("name","")),
    "{{BORROWER_NRIC_1}}":     lambda d: xml_esc((d.get("borrowers") or [{}])[0].get("nric","")),
    "{{BORROWER_ADDRESS_1}}":  lambda d: xml_esc((d.get("borrowers") or [{}])[0].get("address","")),
    "{{BORROWER_NAME_2}}":     lambda d: xml_esc((d.get("borrowers") or [{},{}])[1].get("name","") if len(d.get("borrowers",[]))>1 else ""),
    "{{BORROWER_NRIC_2}}":     lambda d: xml_esc((d.get("borrowers") or [{},{}])[1].get("nric","") if len(d.get("borrowers",[]))>1 else ""),
    "{{BORROWER_NAMES_ALL}}":  lambda d: xml_esc(" & ".join(b.get("name","") for b in d.get("borrowers",[]))),
    "{{BORROWER_NRICS_ALL}}":  lambda d: xml_esc(" / ".join(b.get("nric","") for b in d.get("borrowers",[]))),
    "{{UNIT_NO}}":             lambda d: xml_esc(d.get("unit_no","")),
    "{{PARCEL_NO}}":           lambda d: xml_esc(d.get("parcel_no","")),
    "{{PRICE}}":               lambda d: xml_esc(str(d.get("price",""))),
    "{{LOAN_AMOUNT}}":         lambda d: xml_esc(str(d.get("loan_amount",""))),
    "{{DATE}}":                lambda d: xml_esc(d.get("date","")),
    "{{BANK_NAME}}":           lambda d: xml_esc(d.get("bank_name","")),
    "{{PROJECT_NAME}}":        lambda d: xml_esc(d.get("project_name","OSK OMBAK")),
    "{{DEVELOPER}}":           lambda d: xml_esc(d.get("developer","PJD Sejahtera Sdn. Bhd.")),
}

def fill_bank_template(tmpl_path: Path, data: dict, work_base: Path) -> Path:
    work = work_base / f"bank_{tmpl_path.stem}"
    work.mkdir(parents=True, exist_ok=True)
    unpack(tmpl_path, work)
    for xml_file in work.rglob("*.xml"):
        try:
            txt = xml_file.read_text(encoding="utf-8")
            changed = False
            for ph, fn in PLACEHOLDERS.items():
                if ph in txt:
                    txt = txt.replace(ph, fn(data))
                    changed = True
            if changed:
                xml_file.write_text(txt, encoding="utf-8")
        except Exception:
            pass
    suffix = data.get("unit_no","out").replace("/","_")
    bank   = data.get("bank_name", tmpl_path.stem).replace(" ","_")
    out    = work_base / f"BANK_{bank}_{suffix}.docx"
    pack(work, out, tmpl_path)
    return out

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("🏠 OSK Ombak – Document Automation")
st.caption("Upload NRIC photo(s) and/or sales form → get all filled forms as a ZIP")

# ── Section 1: Upload documents ───────────────────────────────────────────────
st.header("1 · Upload documents")
st.info("Upload the purchaser's NRIC photo(s) and/or the developer's sales form. Claude will extract all details automatically.", icon="📷")

uploaded_files = st.file_uploader(
    "NRIC photo(s) / Sales form / Any combination",
    accept_multiple_files=True,
    type=["jpg","jpeg","png","pdf","webp"],
    help="Accepts photos or scans. For multiple purchasers, upload each NRIC separately."
)

extracted_data = {}
if uploaded_files:
    with st.spinner("🔍 Reading documents with Claude AI..."):
        extractions = []
        for f in uploaded_files:
            ex = extract_from_image(f.read(), f.name)
            extractions.append(ex)
            st.success(f"✓ {f.name} — {ex.get('doc_type','?').replace('_',' ').title()}")
        extracted_data = merge_extractions(extractions)

# ── Section 2: Review & fill in details ──────────────────────────────────────
st.header("2 · Review & confirm details")

col1, col2 = st.columns(2)
with col1:
    unit_no  = st.text_input("Unit No.", value=extracted_data.get("unit_no",""), placeholder="A-24-10")
    parcel_no = st.text_input("Parcel No.", value=extracted_data.get("parcel_no",""), placeholder="472")
with col2:
    type_unit = st.text_input("Unit Type", value=extracted_data.get("type_unit",""), placeholder="A(M)")
    date_str  = st.text_input("Document Date", value=extracted_data.get("date",""), placeholder="13 May 2026")

# Purchasers
st.subheader("Purchaser(s)")
num_purchasers = st.number_input("Number of purchasers", min_value=1, max_value=3, value=max(1, len(extracted_data.get("purchasers",[]))))

purchasers = []
for i in range(int(num_purchasers)):
    with st.expander(f"Purchaser {i+1}" + (" (Primary)" if i==0 else " (Joint)"), expanded=True):
        pre = extracted_data.get("purchasers",[{}]*3)[i] if i < len(extracted_data.get("purchasers",[])) else {}
        c1, c2 = st.columns(2)
        with c1:
            pname = st.text_input("Full name", value=pre.get("name",""), key=f"pname_{i}", placeholder="KEVIN GOH")
        with c2:
            pnric = st.text_input("NRIC No.", value=pre.get("nric",""), key=f"pnric_{i}", placeholder="890715-06-5121")
        paddr = st.text_input("Address", value=pre.get("address",""), key=f"paddr_{i}", placeholder="No. 4, Lorong ...")
        purchasers.append({"name":pname,"nric":pnric,"address":paddr})

# Bank documents
st.subheader("Bank documents (optional)")
bank_templates = sorted(BANK_DIR.glob("*.docx"))

if not bank_templates:
    st.caption("No bank templates uploaded yet. See the sidebar to add them.")
else:
    st.caption(f"{len(bank_templates)} bank template(s) found: {', '.join(t.stem for t in bank_templates)}")

include_bank = st.toggle("Include bank documents", value=bool(bank_templates), disabled=not bank_templates)

borrowers = purchasers.copy()
loan_amount = ""
bank_name   = ""

if include_bank and bank_templates:
    same_as_purchaser = st.checkbox("Borrowers are the same as purchasers", value=True)
    if not same_as_purchaser:
        num_borrowers = st.number_input("Number of borrowers", min_value=1, max_value=3, value=1)
        borrowers = []
        for i in range(int(num_borrowers)):
            with st.expander(f"Borrower {i+1}", expanded=True):
                bc1, bc2 = st.columns(2)
                with bc1:
                    bname = st.text_input("Full name", key=f"bname_{i}", placeholder="KEVIN GOH")
                with bc2:
                    bnric = st.text_input("NRIC No.",  key=f"bnric_{i}", placeholder="890715-06-5121")
                baddr = st.text_input("Address",       key=f"baddr_{i}", placeholder="No. 4, Lorong ...")
                borrowers.append({"name":bname,"nric":bnric,"address":baddr})
    lc1, lc2 = st.columns(2)
    with lc1:
        loan_amount = st.text_input("Loan amount (RM)", placeholder="250000")
    with lc2:
        bank_name = st.text_input("Bank name", placeholder="Maybank / CIMB / Public Bank")

# Optional extras
with st.expander("Optional — TIN / TNB contact / NOK details"):
    oc1, oc2 = st.columns(2)
    with oc1:
        tin   = st.text_input("TIN (Tax ID)", placeholder="C1234567890")
        phone = st.text_input("Phone", placeholder="012-3456789")
        mothers_name = st.text_input("Mother's name (TNB)", placeholder="Tan Ah Moi")
        nok_name = st.text_input("NOK name", placeholder="Goh Wei Ming")
        nok_ic   = st.text_input("NOK NRIC", placeholder="920101-06-1234")
    with oc2:
        email = st.text_input("Email", placeholder="purchaser@email.com")
        nok_mobile = st.text_input("NOK mobile", placeholder="012-9876543")
        nok_home   = st.text_input("NOK home/office no.", placeholder="09-5551234")
        nok_email  = st.text_input("NOK email", placeholder="nok@email.com")
        nok_relationship = st.text_input("NOK relationship", placeholder="Son / Spouse")

# ── Section 3: Generate ───────────────────────────────────────────────────────
st.header("3 · Generate")

if st.button("⚡ Generate all documents", type="primary", use_container_width=True):
    if not unit_no:
        st.error("Please enter a Unit No. before generating.")
        st.stop()
    if not purchasers[0]["name"] or not purchasers[0]["nric"]:
        st.error("Please enter at least the primary purchaser's name and NRIC.")
        st.stop()

    data = {
        "purchasers":  [p for p in purchasers if p["name"]],
        "borrowers":   [b for b in borrowers  if b["name"]],
        "unit_no":     unit_no.upper().strip(),
        "parcel_no":   parcel_no.strip(),
        "type_unit":   type_unit.strip(),
        "date":        date_str.strip(),
        "tin":         tin if 'tin' in dir() else "",
        "email":       email if 'email' in dir() else "",
        "phone":       phone if 'phone' in dir() else "",
        "mothers_name": mothers_name if 'mothers_name' in dir() else "",
        "nok_name":    nok_name if 'nok_name' in dir() else "",
        "nok_ic":      nok_ic if 'nok_ic' in dir() else "",
        "nok_mobile":  nok_mobile if 'nok_mobile' in dir() else "",
        "nok_home":    nok_home if 'nok_home' in dir() else "",
        "nok_email":   nok_email if 'nok_email' in dir() else "",
        "nok_relationship": nok_relationship if 'nok_relationship' in dir() else "",
        "loan_amount": loan_amount,
        "bank_name":   bank_name,
    }

    all_outputs = []
    errors = []

    with st.status("Generating documents...", expanded=True) as status:
        # Lawyer forms
        st.write("📄 Filling lawyer forms & DMC...")
        try:
            import fill_forms
            import importlib; importlib.reload(fill_forms)
            lawyer_files = fill_forms.fill_forms(data)
            all_outputs.extend(lawyer_files)
            st.write(f"   ✓ {len(lawyer_files)} lawyer/DMC forms")
        except Exception as e:
            errors.append(f"Lawyer forms: {e}")
            st.write(f"   ✗ Error: {e}")

        # Bank templates
        if include_bank and bank_templates:
            st.write("🏦 Filling bank templates...")
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp = Path(tmpdir)
                for tmpl in bank_templates:
                    try:
                        out = fill_bank_template(tmpl, data, tmp)
                        # Copy to OUT_DIR
                        dest = OUT_DIR / out.name
                        shutil.copy2(out, dest)
                        all_outputs.append(str(dest))
                        st.write(f"   ✓ {out.name}")
                    except Exception as e:
                        errors.append(f"{tmpl.name}: {e}")
                        st.write(f"   ✗ {tmpl.name}: {e}")

        # Package ZIP
        st.write("📦 Creating ZIP...")
        unit_tag = unit_no.replace("/","_").replace(" ","_")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in all_outputs:
                p = Path(fp)
                if p.exists():
                    zf.write(p, p.name)
        zip_buf.seek(0)

        status.update(label=f"✅ Done — {len(all_outputs)} documents generated", state="complete")

    if errors:
        st.warning("Some errors occurred:\n" + "\n".join(errors))

    st.success(f"✅ {len(all_outputs)} documents ready!")

    # List files
    with st.expander("Files in this package", expanded=True):
        for fp in all_outputs:
            st.markdown(f"• `{Path(fp).name}`")

    # Download button
    st.download_button(
        label=f"⬇ Download OSK_Ombak_{unit_tag}.zip",
        data=zip_buf.getvalue(),
        file_name=f"OSK_Ombak_{unit_tag}.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary"
    )

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Templates")

    st.subheader("Lawyer forms")
    lawyer_templates = sorted(LAWYER_DIR.glob("*.docx"))
    for t in lawyer_templates:
        st.caption(f"• {t.name}")

    st.divider()

    st.subheader("Bank templates")
    bank_list = sorted(BANK_DIR.glob("*.docx"))
    if bank_list:
        for t in bank_list:
            st.caption(f"• {t.name}")
    else:
        st.caption("None yet")

    uploaded_bank = st.file_uploader(
        "Upload bank template (.docx)",
        type=["docx"],
        help="Template must use {{PLACEHOLDER}} syntax. See docs."
    )
    if uploaded_bank:
        dest = BANK_DIR / uploaded_bank.name
        dest.write_bytes(uploaded_bank.read())
        st.success(f"✓ Saved: {uploaded_bank.name}")
        st.rerun()

    st.divider()
    st.caption("**Placeholder reference** (for bank templates)")
    placeholders = [
        ("{{BORROWER_NAME_1}}", "Borrower 1 name"),
        ("{{BORROWER_NRIC_1}}", "Borrower 1 NRIC"),
        ("{{BORROWER_NAME_2}}", "Borrower 2 name"),
        ("{{BORROWER_NRIC_2}}", "Borrower 2 NRIC"),
        ("{{BORROWER_NAMES_ALL}}", "All borrowers joined"),
        ("{{PURCHASER_NAME_1}}", "Purchaser 1 name"),
        ("{{PURCHASER_NRIC_1}}", "Purchaser 1 NRIC"),
        ("{{PURCHASER_NAME_2}}", "Purchaser 2 name"),
        ("{{PURCHASER_NRIC_2}}", "Purchaser 2 NRIC"),
        ("{{UNIT_NO}}", "Unit number"),
        ("{{LOAN_AMOUNT}}", "Loan amount"),
        ("{{BANK_NAME}}", "Bank name"),
        ("{{DATE}}", "Document date"),
        ("{{PRICE}}", "SPA price"),
    ]
    for ph, desc in placeholders:
        st.caption(f"`{ph}` — {desc}")
