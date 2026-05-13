import subprocess, sys, json, re
from pathlib import Path

UPLOADS = Path(__file__).parent / "templates" / "lawyer"
SCRIPTS = Path(__file__).parent / "scripts" / "office"
OUT_DIR = Path(__file__).parent / "outputs"
OUT_DIR.mkdir(exist_ok=True)

SOURCES = {
    "02": Path(__file__).parent / "templates" / "lawyer" / "02_template_fixed.docx",
    "03": UPLOADS / "03__OMBAK_-_Letter_of_Consent_to_Charge_-_Final.docx",
    "04": UPLOADS / "04__OMBAK_-_Letter_of_Consent_to_amend_Building_Plans__General_.docx",
    "05": UPLOADS / "05__OMBAK_-_Consent_to_Letter_to_Approved_Plans__Specific__-_Final__Type_B_and_D_.docx",
    "06": UPLOADS / "06__OMBAK_-_Consent_to_Letter_to_Approved_Plans__Specific__-_Final__Type_C_.docx",
    "07": UPLOADS / "07__OMBAK_-_Declaration_Form_on_TIN.docx",
    "08": UPLOADS / "08__OMBAK_-_TNB_Contact_Details.docx",
    "dmc": UPLOADS / "DRAFT_DMC_20260506.docx",
}

def unpack(src, work):
    subprocess.run(["python", str(SCRIPTS/"unpack.py"), str(src), str(work)],
                   check=True, capture_output=True)

def pack(work, dst, orig):
    subprocess.run(["python", str(SCRIPTS/"pack.py"), str(work), str(dst),
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

def inject_into_cell(txt, para_id, value):
    """Inject a text run into a table cell identified by its first paragraph's paraId."""
    escaped_val = xml_esc(value)
    pattern = re.compile(
        r'(<w:p [^>]*w14:paraId="' + re.escape(para_id) + r'"[^>]*>)(.*?)(</w:p>)',
        re.DOTALL
    )
    def replacer(m):
        inner = m.group(2)
        inner = inner.replace('<w:highlight w:val="yellow"/>', '')
        inner = re.sub(r'<w:r>.*?<w:t[^>]*>\s*</w:t>\s*</w:r>', '', inner, flags=re.DOTALL)
        run = (f'<w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/>'
               f'</w:rPr><w:t xml:space="preserve">{escaped_val}</w:t></w:r>')
        return m.group(1) + inner + run + m.group(3)
    return pattern.sub(replacer, txt, count=1)

# ── XML paragraph builders ────────────────────────────────────────────────────
def arial_para(text, bold=False, sz=20):
    """Build a plain Arial paragraph run."""
    b = "<w:b/>" if bold else ""
    return (
        f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{xml_esc(text)}</w:t></w:r></w:p>'
    )

def sig_cell_content(name, nric, sz=20):
    """Paragraph block for one signature cell: dots / Name / NRIC."""
    return (
        f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t>&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;&#x2026;</w:t></w:r></w:p>'
        f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:b/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">Name: </w:t></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{xml_esc(name)}</w:t></w:r></w:p>'
        f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:b/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">NRIC No.: </w:t></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{xml_esc(nric)}</w:t></w:r></w:p>'
    )

# ── Main fill function ────────────────────────────────────────────────────────
def fill_forms(data):
    # Support list of purchasers: [{name, nric, address}]
    purchasers = data.get("purchasers", [])
    if not purchasers:
        # Fallback: legacy single/dual fields
        p1 = {"name": data.get("name1",""), "nric": data.get("nric1",""), "address": data.get("address","")}
        purchasers = [p1]
        if data.get("name2"):
            purchasers.append({"name": data.get("name2",""), "nric": data.get("nric2",""), "address": data.get("address","")})

    unit_no  = data.get("unit_no","")
    parcel   = data.get("parcel_no","")
    typ      = data.get("type_unit","")
    date_str = data.get("date","")
    tin      = data.get("tin","")
    email    = data.get("email","")
    phone    = data.get("phone","")
    mothers  = data.get("mothers_name","")
    nok_name = data.get("nok_name","")
    nok_ic   = data.get("nok_ic","")
    nok_mob  = data.get("nok_mobile","")
    nok_home = data.get("nok_home","")
    nok_email= data.get("nok_email","")
    nok_rel  = data.get("nok_relationship","")

    single = len(purchasers) == 1
    p1 = purchasers[0]
    p2 = purchasers[1] if not single else None

    n1, ic1 = p1["name"], p1["nric"]
    n2, ic2 = (p2["name"], p2["nric"]) if p2 else ("", "")
    addr = p1.get("address","")

    # Combined strings for forms that use a single merged line
    names_combined = n1 if single else f"{n1} & {n2}"
    nric_combined  = ic1 if single else f"{ic1} / {ic2}"
    names_x = xml_esc(names_combined)
    nric_x  = xml_esc(nric_combined)
    n1_x    = xml_esc(n1)
    n2_x    = xml_esc(n2)
    ic1_x   = xml_esc(ic1)
    ic2_x   = xml_esc(ic2)
    addr_x  = xml_esc(addr)

    # Heading block for forms 03–06:
    # Name only — one line per purchaser, then address
    def heading_lines_03to06():
        lines = []
        for p in purchasers:
            lines.append(p['name'])
        return lines

    suffix = unit_no.replace("/","_").replace(" ","_") if unit_no else "purchaser"
    results = []

    # ── FORM 02 ──────────────────────────────────────────────────────────────
    work = Path(f"/tmp/work_02_{suffix}"); work.mkdir(parents=True, exist_ok=True)
    unpack(SOURCES["02"], work)
    doc = work/"word"/"document.xml"
    txt = read(doc)

    # Heading: name + NRIC line, then Unit No line (insert after [Insert address])
    unit_line = f"Unit No. {unit_no}"
    heading_name = names_combined  # name only, no NRIC
    txt = txt.replace("[Insert name of Purchaser]", xml_esc(heading_name))
    # That approach is messy; do it cleanly
    txt = read(doc)  # re-read fresh
    txt = txt.replace("[Insert name of Purchaser]", xml_esc(heading_name))
    txt = txt.replace("[Insert address]", addr_x)
    write(doc, txt)

    # Unit ref in table
    replace_all(doc, {
        "Unit No\u2026., Type \u2026. in the development":
            f"Unit No. {unit_no} in the development",
        "Date: ": f"Date: {date_str}",
        # Sig cell 1 (Purchaser 1)
        " name of Purchaser]":        f" {n1_x}",
        " Purchaser\u2019s NRIC":     f" {ic1_x}",
        " Purchaser&#x2019;s NRIC":   f" {ic1_x}",
        # Sig cell 2 (Purchaser 2 for joint, blank for single)
        " [ name of Purchaser2]":     f" {n2_x if not single else ''}",
        "[ name of Purchaser2]":      n2_x if not single else "",
        " [Purchaser\u2019s NRIC]":   f" {ic2_x if not single else ''}",
        " [Purchaser&#x2019;s NRIC]": f" {ic2_x if not single else ''}",
    })

    remove_highlights(doc)
    out = OUT_DIR / f"02_Appointment_Sol_{suffix}.docx"
    pack(work, out, SOURCES["02"]); results.append(str(out))

    # ── FORMS 03 / 04 / 05 / 06 — shared heading logic ──────────────────────
    def fill_03to06(num, src, out_name, extra_repl=None):
        work = Path(f"/tmp/work_{num}_{suffix}"); work.mkdir(parents=True, exist_ok=True)
        unpack(src, work)
        doc = work/"word"/"document.xml"
        txt = read(doc)
        _f03_p2_injector = None  # only set for form 03 joint purchasers

        head_lines = heading_lines_03to06()  # e.g. ["Kevin Goh  (890715-06-5121)", "Unit No. A-24-10, Type A(M)"]

        # Forms 03: heading is [Insert name of Purchaser] / [Insert address]
        # head_lines = [name1] or [name1, name2] — name only, no NRIC, no unit line
        if num == "03":
            txt = txt.replace("[Insert name of Purchaser]", n1_x)
            if not single:
                name_para_end = txt.find("</w:p>", txt.find(n1_x))
                if name_para_end != -1:
                    txt = txt[:name_para_end+len("</w:p>")] + arial_para(n2_x, sz=20) + txt[name_para_end+len("</w:p>"):]
            txt = txt.replace("[Insert address]", addr_x)
        else:
            txt = txt.replace("(Purchaser(s) Name)", n1_x)
            txt = txt.replace("(Purchaser(s) NRIC)", n2_x if not single else "")
            txt = txt.replace("(Purchaser(s) Address)", addr_x)

        write(doc, txt)

        # Common replacements for body + sig block
        repl = {
            # Body parcel reference
            "Purchaser(s) Name ":            names_x + " ",
            "Purchaser(s) Name":             names_x,
            ": Purchaser(s) Name ":          f": {names_x} ",
            "Date:":                         f"Date: {date_str}",
        }
        if num == "03":
            repl.update({
                "\u2026\u2026\u2026\u2026\u2026.": f"Unit No. {unit_no}",
                # The "Parcel No. " label immediately before the ellipsis in the value cell
                'xml:space="preserve">Parcel No. </w:t>\n            </w:r>\n            <w:r>\n              <w:rPr>\n                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>\n                <w:sz w:val="20"/>\n                <w:szCs w:val="20"/>\n                <w:highlight w:val="yellow"/>\n              </w:rPr>\n              <w:t>':
                f'xml:space="preserve"></w:t>\n            </w:r>\n            <w:r>\n              <w:rPr>\n                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>\n                <w:sz w:val="20"/>\n                <w:szCs w:val="20"/>\n                <w:highlight w:val="yellow"/>\n              </w:rPr>\n              <w:t>',
                "[Insert name of Purchaser]":     n1_x,
                " [Insert Purchaser&#x2019;s NRIC]": f" {ic1_x}",
                "[Insert Purchaser&#x2019;s NRIC]":  ic1_x,
                "Insert Purchaser&#x2019;s NRIC]":   ic1_x,
                " [Insert name of Purchaser]":    f" {n1_x}",
                "Date: ":                         f"Date: {date_str}",
            })
            repl.pop("Date:", None)  # form 03 uses "Date: " not "Date:"

            # After replacements: inject P2 into sig right cell for joint purchasers
            if not single:
                def _after_repl_03(txt_in):
                    # Right sig cell paraId is 6F4347C8 — currently empty, inject P2 content
                    p2_sig = (
                        f'''<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>{xml_esc("…………………………………")}</w:t></w:r></w:p>'''
                        f'''<w:p><w:pPr><w:keepNext/><w:outlineLvl w:val="8"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'''
                        f'''<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>Name:</w:t></w:r>'''
                        f'''<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve"> {n2_x}</w:t></w:r></w:p>'''
                        f'''<w:p><w:pPr><w:keepNext/><w:outlineLvl w:val="8"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'''
                        f'''<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>NRIC No.:</w:t></w:r>'''
                        f'''<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve"> {ic2_x}</w:t></w:r>'''
                    )
                    return inject_into_cell(txt_in, "6F4347C8", "").replace(
                        'w14:paraId="6F4347C8"' + txt_in[txt_in.find('w14:paraId="6F4347C8"') + len('w14:paraId="6F4347C8"'):txt_in.find('</w:p>', txt_in.find('w14:paraId="6F4347C8"')) + len('</w:p>')],
                        'w14:paraId="6F4347C8"' + '><w:pPr><w:keepNext/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>' + p2_sig
                    )
                # Simpler: direct XML injection using paraId pattern
                import re as _re
                def _inject_03_p2(txt_in):
                    t = txt_in
                    # Find the empty right cell paragraph and replace it
                    old_p = _re.search(
                        r'<w:p w14:paraId="6F4347C8"[^>]*>.*?</w:p>', t, _re.DOTALL
                    )
                    if old_p:
                        new_p = (
                            f'<w:p w14:paraId="6F4347C8" w14:textId="77777777" w:rsidR="009D791B" w:rsidRDefault="009D791B" w:rsidP="009A1B26">'
                            f'<w:pPr><w:keepNext/><w:outlineLvl w:val="8"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
                            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>{xml_esc("…………………………………")}</w:t></w:r>'
                            f'</w:p>'
                            f'<w:p><w:pPr><w:keepNext/><w:outlineLvl w:val="8"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
                            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>Name:</w:t></w:r>'
                            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve"> {n2_x}</w:t></w:r></w:p>'
                            f'<w:p><w:pPr><w:keepNext/><w:outlineLvl w:val="8"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
                            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>NRIC No.:</w:t></w:r>'
                            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve"> {ic2_x}</w:t></w:r></w:p>'
                        )
                        t = t[:old_p.start()] + new_p + t[old_p.end():]
                    return t
                # Store for use after replace_all
                _f03_p2_injector = _inject_03_p2
            else:
                _f03_p2_injector = None
        else:
            # Parcel No. inline fix
            txt2 = read(doc)
            txt2 = txt2.replace(
                '<w:t xml:space="preserve">: Parcel No. </w:t></w:r>',
                f'<w:t xml:space="preserve">: Unit No. {unit_no}</w:t></w:r>'
            )
            write(doc, txt2)

        replace_all(doc, repl)

        # For form 03 joint purchaser: inject P2 into right sig cell
        if num == "03" and _f03_p2_injector is not None:
            write(doc, _f03_p2_injector(read(doc)))

        # Sig block: replace "Purchaser Name" occurrences in order (1st=P1, 2nd=P2/blank)
        # and "Purchaser NRIC" (1st=IC1, 2nd=IC2/blank)
        if num != "03":
            txt3 = read(doc)
            # Replace first "Purchaser Name" → n1, second → n2 or ""
            txt3 = txt3.replace("Purchaser Name", n1_x, 1)
            txt3 = txt3.replace("Purchaser Name", n2_x if not single else "", 1)
            # Replace "Purchaser NRIC" occurrences
            txt3 = txt3.replace("Purchaser NRIC", ic1_x, 1)
            txt3 = txt3.replace("Purchaser NRIC", ic2_x if not single else "", 1)
            # For single purchaser: remove right-column sig elements
            if single:
                txt3 = txt3.replace(
                    "<w:tab/><w:tab/><w:tab/><w:tab/><w:t>\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026.</w:t></w:r>",
                    "</w:r>"
                )
                txt3 = txt3.replace(
                    "<w:tab/><w:tab/><w:tab/><w:tab/><w:tab/><w:t>Name:</w:t><w:tab/></w:r>",
                    "</w:r>"
                )
                txt3 = txt3.replace(
                    '<w:tab/><w:tab/><w:tab/><w:t xml:space=\"preserve\">NRIC No: </w:t></w:r>',
                    "</w:r>"
                )
            write(doc, txt3)
        remove_highlights(doc)
        out = OUT_DIR / out_name
        pack(work, out, src); results.append(str(out))

    fill_03to06("03", SOURCES["03"], f"03_Letter_Consent_Charge_{suffix}.docx")
    fill_03to06("04", SOURCES["04"], f"04_Letter_Consent_Amend_Plans_{suffix}.docx")
    fill_03to06("05", SOURCES["05"], f"05_Consent_Approved_Plans_TypeBD_{suffix}.docx")
    fill_03to06("06", SOURCES["06"], f"06_Consent_Approved_Plans_TypeC_{suffix}.docx")

    # ── FORM 07 ──────────────────────────────────────────────────────────────
    cell_ids_07 = [
        ("04487996", names_x),
        ("3837E410", nric_x),
        ("293B1D4A", xml_esc(tin)),
        ("187B9391", addr_x),
        ("03EE7687", xml_esc(phone)),
        ("07802AC5", xml_esc(email)),
    ]
    work = Path(f"/tmp/work_07_{suffix}"); work.mkdir(parents=True, exist_ok=True)
    unpack(SOURCES["07"], work)
    doc = work/"word"/"document.xml"
    txt = read(doc)
    txt = txt.replace(
        "<w:t>Unit No</w:t>\n      <w:tab/><w:t>:</w:t></w:r>",
        f"<w:t>Unit No</w:t>\n      <w:tab/><w:t xml:space=\"preserve\">:\t{unit_no}</w:t></w:r>"
    )
    txt = txt.replace(
        "<w:t>Date</w:t>\n      <w:tab/><w:tab/><w:t>:</w:t></w:r>",
        f"<w:t>Date</w:t>\n      <w:tab/><w:tab/><w:t xml:space=\"preserve\">:\t{date_str}</w:t></w:r>"
    )
    for para_id, value in cell_ids_07:
        txt = inject_into_cell(txt, para_id, value)
    txt = txt.replace(
        "<w:t>Name</w:t>\n      <w:tab/><w:tab/><w:t>_____________________________________</w:t><w:tab/></w:r>",
        f"<w:t xml:space=\"preserve\">Name\t\t{names_x}</w:t></w:r>"
    )
    write(doc, txt)
    remove_highlights(doc)
    out = OUT_DIR / f"07_Declaration_TIN_{suffix}.docx"
    pack(work, out, SOURCES["07"]); results.append(str(out))

    # ── FORM 08 ──────────────────────────────────────────────────────────────
    work = Path(f"/tmp/work_08_{suffix}"); work.mkdir(parents=True, exist_ok=True)
    unpack(SOURCES["08"], work)
    doc = work/"word"/"document.xml"
    txt = read(doc)
    txt = txt.replace(
        "<w:t>Applicant name</w:t>\n      <w:tab/><w:tab/><w:t>:</w:t><w:tab/><w:t>________________________________</w:t></w:r>",
        f"<w:t>Applicant name</w:t>\n      <w:tab/><w:tab/><w:t>:</w:t><w:tab/><w:t xml:space=\"preserve\">{names_x}  ({nric_x})</w:t></w:r>"
    )
    txt = txt.replace(
        "<w:t>Parcel No</w:t>\n      <w:tab/><w:tab/><w:tab/><w:t>:</w:t><w:tab/><w:t>________________________________</w:t></w:r>",
        f"<w:t>Parcel No</w:t>\n      <w:tab/><w:tab/><w:tab/><w:t>:</w:t><w:tab/><w:t xml:space=\"preserve\">{unit_no}</w:t></w:r>"
    )
    txt = re.sub(
        r"(Mother&#x2019;s name[^<]*)</w:t>\n(\s*)<w:tab/><w:t>_+</w:t></w:r>",
        lambda m: f"{m.group(1)}</w:t>\n{m.group(2)}<w:tab/><w:t xml:space=\"preserve\">{xml_esc(mothers)}</w:t></w:r>",
        txt
    )
    txt = re.sub(
        r"(1\.\s+Name[^<]*)</w:t>\n(\s*)<w:tab/><w:t>_+</w:t></w:r>",
        lambda m: f"{m.group(1)}</w:t>\n{m.group(2)}<w:tab/><w:t xml:space=\"preserve\">{xml_esc(nok_name)}</w:t></w:r>",
        txt
    )
    write(doc, txt)
    remove_highlights(doc)
    out = OUT_DIR / f"08_TNB_Contact_{suffix}.docx"
    pack(work, out, SOURCES["08"]); results.append(str(out))

    # ── DMC (Deed of Mutual Covenants) ──────────────────────────────────────
    # Inject name + NRIC after the dotted line in the purchaser signature text box.
    # The text box appears in BOTH the modern wps:wsp and legacy v:shape fallback —
    # we must update both so the document renders correctly in all Word versions.
    work = Path(f"/tmp/work_dmc_{suffix}"); work.mkdir(parents=True, exist_ok=True)
    unpack(SOURCES["dmc"], work)
    doc = work/"word"/"document.xml"
    txt = read(doc)

    # Build the name + NRIC lines to inject as paragraphs inside the text box
    # One line per purchaser: "Name (NRIC No.)"
    def dmc_name_paras():
        paras = ""
        for p in purchasers:
            pn = xml_esc(p["name"])
            pic = xml_esc(p["nric"])
            paras += (
                f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial Narrow" w:hAnsi="Arial Narrow"/>' 
                f'<w:szCs w:val="22"/></w:rPr></w:pPr>' 
                f'<w:r><w:rPr><w:rFonts w:ascii="Arial Narrow" w:hAnsi="Arial Narrow"/>' 
                f'<w:szCs w:val="22"/></w:rPr>' 
                f'<w:t xml:space="preserve">{pn}  ({pic})</w:t></w:r></w:p>'
            )
        return paras

    name_paras = dmc_name_paras()

    # Target: the first empty paragraph AFTER the dotted line in the text box
    # paraId 1CF4F090 is the first blank para after the dots (in both modern and legacy)
    DOTS = "……………………………………………………………"
    # After the dots closing </w:p>, inject name paragraphs
    dots_end = '</w:t>\n                            </w:r>\n                          </w:p>'
    dots_end_legacy = '</w:t>\n                      </w:r>\n                    </w:p>'

    txt = txt.replace(
        dots_end,
        dots_end + name_paras,
        1  # modern text box only (first occurrence)
    )
    txt = txt.replace(
        dots_end_legacy,
        dots_end_legacy + name_paras,
        1  # legacy fallback (first occurrence)
    )

    write(doc, txt)
    remove_highlights(doc)
    out = OUT_DIR / f"DMC_{suffix}.docx"
    pack(work, out, SOURCES["dmc"]); results.append(str(out))

    return results


if __name__ == "__main__":
    data = json.loads(sys.argv[1])
    files = fill_forms(data)
    print(json.dumps(files))
