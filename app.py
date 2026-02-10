import json
import math
import uuid
import copy
from pathlib import Path
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

"""
Bid Manager with PDF export
===========================

This module provides a complete bid management GUI application using
Tkinter.  It closely follows the original specification but adds a
professional PDF export feature for each major tab (Job Info, Cost
Codes, Quotes and Frame Schedule).  When the user clicks the Export
PDF button on a tab, a nicely formatted report is generated using
reportlab (if available) and saved to a location chosen via a file
dialog.  If reportlab is not installed, the application gracefully
informs the user that PDF export is unavailable.

All existing functionality—data entry, undo support, autosave,
normalisation of quotes and schedules, etc.—is preserved.  The added
report generation functions gather the current state of the job and
render it into tables and headings appropriate for each tab.
"""

# Attempt to import reportlab for PDF generation.  If it fails
# gracefully, PDF export functions will alert the user that they are
# unavailable.  When generating PDFs we use the platypus API to
# assemble paragraphs and tables.
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# ======================================================
# Optional date picker (tkcalendar). If not installed,
# falls back to Entry with YYYY-MM-DD.
# ======================================================
try:
    from tkcalendar import DateEntry  # pip install tkcalendar
    HAS_TKCALENDAR = True
except Exception:
    DateEntry = None
    HAS_TKCALENDAR = False


# ======================================================
# Storage
# ======================================================

BASE_DIR = Path(__file__).parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

def job_path(job_id: str) -> Path:
    return JOBS_DIR / f"{job_id}.json"

def save_job(job: dict) -> None:
    job_path(job["id"]).write_text(json.dumps(job, indent=2), encoding="utf-8")

def load_jobs():
    jobs = []
    for f in JOBS_DIR.glob("*.json"):
        try:
            jobs.append(json.loads(f.read_text(encoding="utf-8")))
        except Exception:
            pass
    return jobs


# ======================================================
# Helpers
# ======================================================

def today_str() -> str:
    return date.today().isoformat()

def roundup(v: float) -> int:
    return int(math.ceil(v))

def safe_float(x) -> float:
    try:
        s = str(x).strip()
        if s == "":
            return 0.0
        return float(s)
    except Exception:
        return 0.0

def safe_int(x) -> int:
    try:
        return int(float(str(x).strip()))
    except Exception:
        return 0

def parse_money(txt: str) -> int:
    if not txt:
        return 0
    s = str(txt).replace("$", "").replace(",", "").strip()
    digits = "".join(c for c in s if c.isdigit())
    return int(digits) if digits else 0

def money_fmt(v) -> str:
    try:
        iv = int(float(v))
    except Exception:
        iv = 0
    return f"${iv:,}" if iv else "$0"

def parse_pct(txt: str) -> float:
    if not txt:
        return 0.0
    s = str(txt).replace("%", "").replace(",", "").strip()
    if not s:
        return 0.0
    cleaned = []
    dot_used = False
    for ch in s:
        if ch.isdigit():
            cleaned.append(ch)
        elif ch == "." and not dot_used:
            cleaned.append(ch)
            dot_used = True
    try:
        return float("".join(cleaned)) if cleaned else 0.0
    except Exception:
        return 0.0

def pct_fmt(v: float) -> str:
    try:
        fv = float(v)
    except Exception:
        fv = 0.0
    s = f"{fv:.4f}".rstrip("0").rstrip(".")
    return f"{s}%"

def calc_cost(price: int, surcharge_pct: float) -> int:
    p = safe_int(price)
    s = safe_float(surcharge_pct)
    return p + int(p * (s / 100.0))

def parse_alts(text: str):
    t = (text or "").strip()
    if not t:
        return []
    parts = [p.strip() for p in t.split(",") if p.strip()]
    out = []
    for p in parts:
        if not p.isdigit():
            continue
        n = int(p)
        if 1 <= n <= 25 and n not in out:
            out.append(n)
    return sorted(out)

def variants_for_cc(alts):
    alts = alts or []
    out = ["BASE"]
    for n in alts:
        out.append(f"ALT{n}")
    return out

def frame_spec_id(base_code: str, variant: str) -> str:
    return f"{base_code}||{variant}"

def parse_frame_spec_id(spec_id: str):
    if "||" in (spec_id or ""):
        base, var = spec_id.split("||", 1)
        return base.strip(), (var.strip() or "BASE")
    return (spec_id or "").strip(), "BASE"

def frame_spec_label(base_code: str, variant: str) -> str:
    return base_code if variant == "BASE" else f"{variant} {base_code}"


# ======================================================
# Date widgets
# ======================================================

def make_date_widget(parent, initial: str):
    if HAS_TKCALENDAR:
        w = DateEntry(parent, width=12, date_pattern="yyyy-mm-dd")
        if (initial or "").strip():
            try:
                w.set_date(initial)
            except Exception:
                w.set_date(today_str())
        return w
    e = ttk.Entry(parent, width=12)
    e.insert(0, initial or "")
    return e

def get_date_value(widget) -> str:
    if HAS_TKCALENDAR and isinstance(widget, DateEntry):
        try:
            return widget.get_date().isoformat()
        except Exception:
            return ""
    return (widget.get() or "").strip()

def set_date_value(widget, value: str):
    if HAS_TKCALENDAR and isinstance(widget, DateEntry):
        try:
            widget.set_date(value)
        except Exception:
            pass
    else:
        widget.delete(0, tk.END)
        widget.insert(0, value)


# ======================================================
# Scroll frame
# ======================================================

class ScrollFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.window_id, width=e.width))

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")


# ======================================================
# Data defaults
# ======================================================

JOB_FIELDS = [
    ("Job Name", "job_name", "text"),
    ("Bid Due Date", "bid_due_date", "date"),
    ("Project Number", "project_number", "text"),
    ("Walkthrough", "walkthrough", "bool"),
    ("MWD PO", "mwd_po", "text"),
    ("Addenda Count", "addenda_count", "text"),
    ("Job Address", "job_address", "text"),
    ("Plan Source", "plan_source", "text"),
    ("Owner Name", "owner_name", "text"),
    ("Start Date", "start_date", "date"),
    ("Owner Address", "owner_address", "text"),
    ("Completion Date", "completion_date", "date"),
    ("Architect(s)", "architects", "text"),
    ("Fab Start Date", "fab_start_date", "date"),
    ("Project Type", "project_type", "text"),
    ("Fab Due Date", "fab_due_date", "date"),
    ("Building Type", "building_type", "text"),
    ("Estimator", "estimator", "text"),
    ("Sales Person", "sales_person", "text"),
    ("Project Manager", "project_manager", "text"),
    ("Contact Name", "contact_name", "text"),
    ("Contact Email", "contact_email", "text"),
    ("Contact Phone", "contact_phone", "text"),
    ("Status", "status", "text"),
]

DEFAULT_STATUS = "Bid"

def default_config():
    return {
        "materials": [
            {"key": "bracing",     "label": "Bracing and Anchoring",              "basis": "perim_subtotal",     "factor": "1.00",   "rate": "1.50", "qty": "", "unit": "Linear Foot"},
            {"key": "sheet_metal", "label": "Sheet Metal Membrane Air Barriers",  "basis": "perim_subtotal",     "factor": "1.00",   "rate": "1.00", "qty": "", "unit": "Linear Foot"},
            {"key": "flashing",    "label": "Flashing and Sheet Metal",           "basis": "head_sill_subtotal", "factor": "1.00",   "rate": "8.00", "qty": "", "unit": "Linear Foot"},
            {"key": "backer_rods", "label": "Backer Rods",                        "basis": "caulk_lf_subtotal",  "factor": "1.00",   "rate": "0.50", "qty": "", "unit": "Linear Foot"},
            {"key": "sealants",    "label": "Joint Sealants",                     "basis": "caulk_lf_subtotal",  "factor": "0.0833", "rate": "12.00","qty": "", "unit": "Sausage"},
        ],
        "sheet_type_options": ["HOLLOW METAL", "ALUM", "WOOD", "STAINLESS", "OTHER"],
        "product_type_options": ["FRAME", "DOOR", "LOUVER", "WINDOW", "CURTAINWALL", "OTHER"],
    }

def new_job_template(name: str):
    jid = str(uuid.uuid4())
    j = {"id": jid}
    for _, key, t in JOB_FIELDS:
        if t == "bool":
            j[key] = False
        elif t == "date":
            j[key] = today_str() if "date" in key else ""
        else:
            j[key] = ""
    j["job_name"] = name.strip()
    j["status"] = DEFAULT_STATUS
    j["cost_codes"] = []
    j["quotes"] = {}
    j["bid_sheet"] = {}
    j["frame_schedules"] = {}
    j["frame_schedule_rollups"] = {}
    j["config"] = default_config()
    return j

def build_valid_frame_spec_ids(job: dict):
    ids = []
    for cc in job.get("cost_codes", []):
        code = (cc.get("code") or "").strip()
        if not code:
            continue
        for v in variants_for_cc(cc.get("alts", [])):
            ids.append(frame_spec_id(code, v))
    return ids

def ensure_job_defaults(job: dict) -> None:
    job.setdefault("id", str(uuid.uuid4()))
    for _, key, t in JOB_FIELDS:
        if key not in job:
            job[key] = False if t == "bool" else ""
    job.setdefault("cost_codes", [])
    job.setdefault("quotes", {})
    job.setdefault("bid_sheet", {})
    job.setdefault("frame_schedules", {})
    job.setdefault("frame_schedule_rollups", {})
    job.setdefault("config", default_config())

    for cc in job["cost_codes"]:
        cc.setdefault("id", str(uuid.uuid4()))
        cc.setdefault("code", "00 00 00")
        cc.setdefault("description", "")
        cc.setdefault("alts", [])

    normalize_quotes(job)
    normalize_bid_sheet(job)
    normalize_frame_schedules(job)


# ======================================================
# Quotes normalization
# ======================================================

def normalize_quotes(job: dict) -> None:
    job.setdefault("quotes", {})
    codes_present = set()

    for cc in job.get("cost_codes", []):
        code = (cc.get("code") or "").strip()
        if not code:
            continue
        codes_present.add(code)
        job["quotes"].setdefault(code, {})
        needed_variants = variants_for_cc(cc.get("alts", []))

        if isinstance(job["quotes"][code], list):
            job["quotes"][code] = {"BASE": job["quotes"][code]}
        if not isinstance(job["quotes"][code], dict):
            job["quotes"][code] = {}

        for v in needed_variants:
            job["quotes"][code].setdefault(v, [])
            if len(job["quotes"][code][v]) == 0:
                job["quotes"][code][v].append({})

        for v in list(job["quotes"][code].keys()):
            if v not in needed_variants:
                del job["quotes"][code][v]

    for code in list(job["quotes"].keys()):
        if code not in codes_present:
            del job["quotes"][code]


# ======================================================
# Bid sheet normalization
# ======================================================

def normalize_bid_sheet(job: dict) -> None:
    job.setdefault("bid_sheet", {})
    valid_specs = build_valid_frame_spec_ids(job)
    for spec_id in valid_specs:
        row = job["bid_sheet"].setdefault(spec_id, {})
        row.setdefault("markup_pct", "")
        row.setdefault("markup_amt", "")
        row.setdefault("markup_source", "pct")
        row.setdefault("notes", "")
        row.setdefault("color", "None")
    for spec_id in list(job["bid_sheet"].keys()):
        if spec_id not in valid_specs:
            del job["bid_sheet"][spec_id]

def bid_sheet_direct_cost(job: dict, code: str, variant: str) -> int:
    quotes_list = job.get("quotes", {}).get(code, {}).get(variant, [])
    return sum(
        safe_int(q.get("cost", 0) or 0)
        for q in quotes_list
        if safe_int(q.get("cost", 0) or 0) > 0
    )

def bid_sheet_install_material_total(job: dict, spec_id: str) -> int:
    rollup = job.get("frame_schedule_rollups", {}).get(spec_id, {})
    return safe_int(rollup.get("install_material_total", 0))

def bid_sheet_material_breakdown(job: dict, spec_id: str):
    totals = {}
    for section in job.get("frame_schedules", {}).get(spec_id, []) or []:
        if not isinstance(section, dict):
            continue
        recompute_section_totals(section)
        rows = section.get("rows", []) or []
        subs = schedule_subtotals(rows)
        basis_subs = {
            "perim": safe_int(subs.get("perim", 0)),
            "caulk_lf": safe_int(subs.get("caulk_lf", 0)),
            "head_sill": safe_int(subs.get("head_sill", 0)),
        }
        for m in section.get("materials", []) or []:
            if not isinstance(m, dict):
                continue
            label = (m.get("label") or "").strip() or "Install Material"
            qty_val = material_qty_for_row(m, basis_subs)
            rate_val = safe_float(m.get("rate", ""))
            cost_val = roundup(qty_val * rate_val) if qty_val * rate_val != 0 else 0
            totals[label] = totals.get(label, 0) + cost_val
    return sorted(totals.items(), key=lambda x: x[0].lower())

def compute_bid_sheet_total(job: dict) -> int:
    normalize_quotes(job)
    normalize_bid_sheet(job)
    normalize_frame_schedules(job)
    compute_frame_schedule_rollups(job)
    total = 0
    for spec_id, data in job.get("bid_sheet", {}).items():
        base, variant = parse_frame_spec_id(spec_id)
        base_cost = bid_sheet_direct_cost(job, base, variant)
        install_total = bid_sheet_install_material_total(job, spec_id)
        cost_value = base_cost + install_total
        pct_val = parse_pct(data.get("markup_pct", ""))
        amt_val = parse_money(data.get("markup_amt", ""))
        source = (data.get("markup_source") or "pct").strip().lower()
        if source == "amt":
            total += cost_value + (amt_val if amt_val > 0 else 0)
        elif pct_val > 0:
            total += cost_value + roundup(cost_value * (pct_val / 100.0))
        else:
            total += cost_value
    return total


# ======================================================
# Frame schedule model normalization + deterministic totals
# ======================================================

FS_HEADERS = [
    ("SPEC / MARK", "spec_mark", 16, "text"),
    ("QTY", "qty", 6, "num"),
    ("WIDTH", "width", 8, "num"),
    ("HEIGHT", "height", 8, "num"),
    ("SQFT", "sqft", 8, "calc"),
    ("PERIM", "perim", 8, "calc"),
    ("CAULK LF", "caulk_lf", 10, "calc"),
    ("HEAD/SILL", "head_sill", 10, "calc"),
]

SUBTOTAL_KEYS = ["qty", "sqft", "perim", "caulk_lf", "head_sill"]

def blank_frame_schedule_row() -> dict:
    return {
        "spec_mark": "",
        "qty": "",
        "width": "",
        "height": "",
        "sqft": 0,
        "perim": 0,
        "caulk_lf": 0,
        "head_sill": 0,
    }

def normalize_frame_schedules(job: dict) -> None:
    job.setdefault("frame_schedules", {})
    valid_specs = build_valid_frame_spec_ids(job)

    for spec_id in valid_specs:
        sections = job["frame_schedules"].setdefault(spec_id, [])
        if not isinstance(sections, list):
            sections = []
            job["frame_schedules"][spec_id] = sections

        seen_section_ids = set()
        deduped_sections = []
        for section in sections:
            normalize_section(section, job)
            section_id = (section.get("id") or "").strip()
            while section_id in seen_section_ids:
                section_id = str(uuid.uuid4())
                section["id"] = section_id
            seen_section_ids.add(section_id)
            deduped_sections.append(section)
        job["frame_schedules"][spec_id] = deduped_sections

    for spec_id in list(job["frame_schedules"].keys()):
        if spec_id not in valid_specs:
            del job["frame_schedules"][spec_id]

def is_valid_frame_spec_id(job: dict, spec_id: str) -> bool:
    return spec_id in build_valid_frame_spec_ids(job)

def create_frame_schedule_section(job: dict, spec_id: str):
    if not spec_id or not is_valid_frame_spec_id(job, spec_id):
        return None, "Please select a valid cost code before adding a section."

    job.setdefault("frame_schedules", {})
    current_sections = job["frame_schedules"].get(spec_id, [])
    if not isinstance(current_sections, list):
        current_sections = []

    existing_ids = {
        (s.get("id") or "").strip()
        for s in current_sections
        if isinstance(s, dict)
    }
    section_id = str(uuid.uuid4())
    while section_id in existing_ids:
        section_id = str(uuid.uuid4())

    next_index = len(current_sections) + 1
    new_section = {
        "id": section_id,
        "name": f"Section {next_index}",
        "rows": [blank_frame_schedule_row()],
        "materials": copy.deepcopy(job.get("config", {}).get("materials", default_config()["materials"])),
    }
    normalize_section(new_section, job)

    # immutable append so UI/state never relies on list mutation side effects
    job["frame_schedules"][spec_id] = list(current_sections) + [new_section]
    return new_section, ""

def normalize_section(section: dict, job: dict) -> None:
    section.setdefault("id", str(uuid.uuid4()))
    section.setdefault("name", "Section")
    section.setdefault("rows", [])
    section.setdefault("materials", copy.deepcopy(job.get("config", {}).get("materials", default_config()["materials"])))
    if not isinstance(section["rows"], list):
        section["rows"] = []
    if not section["rows"]:
        section["rows"] = [blank_frame_schedule_row()]
    if not isinstance(section["materials"], list):
        section["materials"] = []

    for r in section["rows"]:
        normalize_fs_row(r)

    cfg_materials = job.get("config", {}).get("materials", default_config()["materials"])
    by_key = {m.get("key"): m for m in section["materials"] if isinstance(m, dict)}
    normalized = []
    for tmpl in cfg_materials:
        k = tmpl.get("key")
        existing = by_key.get(k, {})
        row = dict(tmpl)
        for fld in ("label", "basis", "factor", "rate", "qty", "unit"):
            if fld in existing:
                row[fld] = existing[fld]
        normalized.append(row)
    section["materials"] = normalized
    recompute_section_totals(section)

def normalize_fs_row(r: dict):
    r.setdefault("spec_mark", "")
    r.setdefault("qty", "")
    r.setdefault("width", "")
    r.setdefault("height", "")
    r.setdefault("sqft", 0)
    r.setdefault("perim", 0)
    r.setdefault("caulk_lf", 0)
    r.setdefault("head_sill", 0)

def recalc_row_fields(r: dict) -> None:
    qty = safe_int(r.get("qty"))
    w = safe_float(r.get("width"))
    h = safe_float(r.get("height"))
    sqft = qty * w * h / 144.0
    perim = qty * (2 * (w + h)) / 12.0
    caulk = perim
    hs = qty * (w / 12.0) * 2.0
    r["sqft"] = roundup(sqft)
    r["perim"] = roundup(perim)
    r["caulk_lf"] = roundup(caulk)
    r["head_sill"] = roundup(hs)

def schedule_subtotals(rows):
    out = {k: 0 for k in SUBTOTAL_KEYS}
    for r in rows:
        recalc_row_fields(r)
        out["qty"] += safe_int(r.get("qty"))
        out["sqft"] += safe_int(r.get("sqft"))
        out["perim"] += safe_int(r.get("perim"))
        out["caulk_lf"] += safe_int(r.get("caulk_lf"))
        out["head_sill"] += safe_int(r.get("head_sill"))
    return out

def material_qty_for_row(material: dict, subtotals: dict):
    basis = material.get("basis", "")
    factor = safe_float(material.get("factor", "1"))
    if basis == "perim_subtotal":
        return subtotals["perim"] * factor
    if basis == "head_sill_subtotal":
        return subtotals["head_sill"] * factor
    if basis == "caulk_lf_subtotal":
        return subtotals["caulk_lf"] * factor
    if basis == "manual":
        return safe_float(material.get("qty", ""))
    return 0.0

def recompute_section_totals(section: dict) -> None:
    rows = section.get("rows", [])
    subs = schedule_subtotals(rows)
    mats = section.get("materials", [])
    mrows = []
    install_total = 0
    for m in mats:
        qty = material_qty_for_row(m, subs)
        rate = safe_float(m.get("rate", ""))
        cost = roundup(qty * rate) if qty * rate != 0 else 0
        install_total += cost
        mrows.append({
            "key": m.get("key", ""),
            "label": m.get("label", ""),
            "basis": m.get("basis", ""),
            "factor": m.get("factor", ""),
            "rate": m.get("rate", ""),
            "unit": m.get("unit", ""),
            "qty_calc": qty,
            "cost_calc": cost,
        })
    section["_subtotals"] = subs
    section["_material_rows"] = mrows
    section["_install_total"] = install_total

def compute_frame_schedule_rollups(job: dict) -> None:
    normalize_frame_schedules(job)
    rollups = {}
    for spec_id, sections in job.get("frame_schedules", {}).items():
        overall_subs = {k: 0 for k in SUBTOTAL_KEYS}
        install_total = 0
        for section in sections:
            recompute_section_totals(section)
            subs = section.get("_subtotals", {})
            for k in SUBTOTAL_KEYS:
                overall_subs[k] += safe_int(subs.get(k, 0))
            install_total += safe_int(section.get("_install_total", 0))
        rollups[spec_id] = {
            "subtotals": overall_subs,
            "install_material_total": install_total,
        }
    job["frame_schedule_rollups"] = rollups


# ======================================================
# Undo manager
# ======================================================

class UndoManager:
    def __init__(self):
        self.stack = []

    def snapshot(self, job: dict):
        self.stack.append(copy.deepcopy(job))
        if len(self.stack) > 50:
            self.stack.pop(0)

    def undo(self):
        if not self.stack:
            return None
        return self.stack.pop()


# ======================================================
# PDF export helpers
# ======================================================

def _require_reportlab():
    if not REPORTLAB_AVAILABLE:
        messagebox.showwarning(
            "PDF Unavailable",
            "PDF export requires reportlab.\nInstall with:\n\npip install reportlab"
        )
        return False
    return True

def _ask_pdf_path(default_name: str):
    return filedialog.asksaveasfilename(
        title="Save PDF",
        defaultextension=".pdf",
        initialfile=default_name,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )

def _styles():
    return getSampleStyleSheet()

def export_job_info_pdf(job: dict):
    if not _require_reportlab():
        return
    path = _ask_pdf_path(f"{(job.get('job_name') or 'job')}_job_info.pdf")
    if not path:
        return

    doc = SimpleDocTemplate(path, pagesize=letter)
    styles = _styles()
    story = [
        Paragraph(f"<b>Job Info Report</b>", styles["Title"]),
        Spacer(1, 8),
        Paragraph(f"Generated: {today_str()}", styles["Normal"]),
        Spacer(1, 10),
    ]

    table_data = [["Field", "Value"]]
    for label, key, _t in JOB_FIELDS:
        val = job.get(key, "")
        if isinstance(val, bool):
            val = "Yes" if val else "No"
        table_data.append([label, str(val)])

    t = Table(table_data, repeatRows=1, colWidths=[180, 360])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t)

    doc.build(story)
    messagebox.showinfo("PDF Export", f"Saved:\n{path}")

def export_cost_codes_pdf(job: dict):
    if not _require_reportlab():
        return
    path = _ask_pdf_path(f"{(job.get('job_name') or 'job')}_cost_codes.pdf")
    if not path:
        return

    doc = SimpleDocTemplate(path, pagesize=letter)
    styles = _styles()
    story = [
        Paragraph("<b>Cost Codes Report</b>", styles["Title"]),
        Spacer(1, 8),
        Paragraph(f"Job: {job.get('job_name','')}", styles["Normal"]),
        Spacer(1, 10),
    ]

    rows = sorted(job.get("cost_codes", []), key=lambda x: (x.get("code") or ""))
    data = [["Code", "Description", "Alternates"]]
    for cc in rows:
        code = cc.get("code", "")
        desc = cc.get("description", "")
        alts = cc.get("alts", []) or []
        alts_txt = ", ".join(f"ALT{n}" for n in alts) if alts else "—"
        data.append([code, desc, alts_txt])

    t = Table(data, repeatRows=1, colWidths=[90, 340, 110])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2e6b3f")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t)

    doc.build(story)
    messagebox.showinfo("PDF Export", f"Saved:\n{path}")

def export_quotes_pdf(job: dict):
    if not _require_reportlab():
        return
    path = _ask_pdf_path(f"{(job.get('job_name') or 'job')}_quotes.pdf")
    if not path:
        return

    normalize_quotes(job)

    doc = SimpleDocTemplate(path, pagesize=letter)
    styles = _styles()
    story = [
        Paragraph("<b>Quotes Report</b>", styles["Title"]),
        Spacer(1, 8),
        Paragraph(f"Job: {job.get('job_name','')}", styles["Normal"]),
        Spacer(1, 12),
    ]

    by_code = sorted(job.get("quotes", {}).keys())
    for code in by_code:
        story.append(Paragraph(f"<b>{code}</b>", styles["Heading3"]))
        cc_desc = ""
        for cc in job.get("cost_codes", []):
            if (cc.get("code") or "").strip() == code:
                cc_desc = cc.get("description", "")
                break
        if cc_desc:
            story.append(Paragraph(cc_desc, styles["Normal"]))
        story.append(Spacer(1, 4))

        variants = job["quotes"].get(code, {})
        for variant in sorted(variants.keys(), key=lambda v: (v != "BASE", v)):
            qlist = variants.get(variant, [])
            heading = "BASE" if variant == "BASE" else variant
            story.append(Paragraph(f"<b>{heading}</b>", styles["Heading4"]))

            data = [["Date", "Vendor(s)", "Price", "Sur %", "Cost", "Notes"]]
            total = 0
            count = 0
            for q in qlist:
                ensure_quote_defaults(q)
                price = safe_int(q.get("price"))
                sur = safe_float(q.get("surcharge_pct"))
                cost = safe_int(q.get("cost"))
                total += cost
                count += 1 if cost > 0 else 0
                data.append([
                    q.get("date", ""),
                    q.get("vendors", ""),
                    money_fmt(price),
                    pct_fmt(sur) if sur else "",
                    money_fmt(cost),
                    q.get("notes", ""),
                ])
            avg = roundup(total / count) if count else 0
            data.append(["", "", "", "", f"Total {money_fmt(total)} · Avg {money_fmt(avg)}", ""])

            t = Table(data, repeatRows=1, colWidths=[60, 140, 70, 55, 70, 170])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#444444")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("ALIGN", (2, 1), (4, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))
            story.append(t)
            story.append(Spacer(1, 8))

    doc.build(story)
    messagebox.showinfo("PDF Export", f"Saved:\n{path}")

def export_frame_schedule_pdf(job: dict):
    if not _require_reportlab():
        return
    path = _ask_pdf_path(f"{(job.get('job_name') or 'job')}_frame_schedule.pdf")
    if not path:
        return

    normalize_frame_schedules(job)
    compute_frame_schedule_rollups(job)

    doc = SimpleDocTemplate(path, pagesize=letter)
    styles = _styles()
    story = [
        Paragraph("<b>Frame Schedule Report</b>", styles["Title"]),
        Spacer(1, 8),
        Paragraph(f"Job: {job.get('job_name','')}", styles["Normal"]),
        Spacer(1, 12),
    ]

    for cc in sorted(job.get("cost_codes", []), key=lambda x: (x.get("code") or "")):
        code = (cc.get("code") or "").strip()
        if not code:
            continue
        desc = cc.get("description", "")
        for variant in variants_for_cc(cc.get("alts", [])):
            spec_id = frame_spec_id(code, variant)
            sections = job.get("frame_schedules", {}).get(spec_id, []) or []
            if not sections:
                continue
            story.append(Paragraph(f"<b>{frame_spec_label(code, variant)} — {desc}</b>", styles["Heading3"]))
            story.append(Spacer(1, 4))

            for sec in sections:
                normalize_section(sec, job)
                recompute_section_totals(sec)
                story.append(Paragraph(f"<b>Section:</b> {sec.get('name','')}", styles["Heading4"]))

                headers = [h[0] for h in FS_HEADERS]
                table_data = [headers]
                for r in sec.get("rows", []):
                    recalc_row_fields(r)
                    table_data.append([
                        r.get("spec_mark", ""),
                        str(safe_int(r.get("qty", "")) if str(r.get("qty","")).strip() else ""),
                        str(safe_float(r.get("width", "")) if str(r.get("width","")).strip() else ""),
                        str(safe_float(r.get("height", "")) if str(r.get("height","")).strip() else ""),
                        str(safe_int(r.get("sqft", 0))),
                        str(safe_int(r.get("perim", 0))),
                        str(safe_int(r.get("caulk_lf", 0))),
                        str(safe_int(r.get("head_sill", 0))),
                    ])
                subs = schedule_subtotals(sec.get("rows", []))
                subtotal_row = []
                for _, k, *_ in FS_HEADERS:
                    if k == "spec_mark":
                        subtotal_row.append("Subtotal")
                    elif k in SUBTOTAL_KEYS:
                        subtotal_row.append(str(subs[k]))
                    else:
                        subtotal_row.append("")
                table_data.append(subtotal_row)

                t = Table(table_data, repeatRows=1, colWidths=[95, 35, 45, 45, 40, 40, 50, 50])
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#666666")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                    ("BACKGROUND", (0, len(table_data)-1), (-1, len(table_data)-1), colors.lightgrey),
                    ("FONTNAME", (0, len(table_data)-1), (-1, len(table_data)-1), "Helvetica-Bold"),
                ]))
                story.append(t)
                story.append(Spacer(1, 6))

                m_headers = ["Material", "Basis", "Factor", "Rate", "Qty", "Cost"]
                m_data = [m_headers]
                for m in sec.get("_material_rows", []):
                    m_data.append([
                        m.get("label", ""),
                        m.get("basis", ""),
                        str(m.get("factor", "")),
                        str(m.get("rate", "")),
                        f"{safe_float(m.get('qty_calc',0)):.2f}",
                        money_fmt(m.get("cost_calc", 0)),
                    ])
                m_data.append(["", "", "", "", "Install Total", money_fmt(sec.get("_install_total", 0))])

                mt = Table(m_data, repeatRows=1, colWidths=[170, 120, 55, 55, 60, 60])
                mt.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3b6ea5")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                    ("BACKGROUND", (0, len(m_data)-1), (-1, len(m_data)-1), colors.lightgrey),
                    ("FONTNAME", (0, len(m_data)-1), (-1, len(m_data)-1), "Helvetica-Bold"),
                ]))
                story.append(mt)
                story.append(Spacer(1, 10))

            roll = job.get("frame_schedule_rollups", {}).get(spec_id, {})
            install_total = safe_int(roll.get("install_material_total", 0))
            story.append(Paragraph(f"<b>Spec Install Material Total:</b> {money_fmt(install_total)}", styles["Normal"]))
            story.append(Spacer(1, 12))

    doc.build(story)
    messagebox.showinfo("PDF Export", f"Saved:\n{path}")


# ======================================================
# UI
# ======================================================

def open_job_tab(root, notebook, job: dict, refresh_main, close_tab, config_manager=None):
    ensure_job_defaults(job)
    undo = UndoManager()

    tab = ttk.Frame(notebook)
    notebook.add(tab, text=job.get("job_name", "Job"))
    notebook.select(tab)

    # Top bar
    top = ttk.Frame(tab)
    top.pack(fill="x", padx=8, pady=6)
    ttk.Button(top, text="Save", command=lambda: (save_job(job), refresh_main())).pack(side="left", padx=3)
    ttk.Button(top, text="Undo", command=lambda: do_undo()).pack(side="left", padx=3)
    ttk.Button(top, text="Close Tab", command=lambda: close_tab(tab)).pack(side="left", padx=3)

    nb = ttk.Notebook(tab)
    nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    # Shared state for cross-tab updates
    quotes_scroll = None
    quote_ui = {}
    spec_frames = {}
    quotes_initialized = False
    quotes_dirty = True

    bid_sheet_dirty = {"value": True}
    bid_sheet_tab_ref = {"tab": None}
    bid_sheet_refresh_fn = {"fn": None}
    bid_sheet_after_id = {"id": None}

    frame_initialized = False
    frame_dirty = True
    current_spec_id_var = tk.StringVar(value="")
    selected_spec_label_var = tk.StringVar(value="")
    frame_schedule_error_var = tk.StringVar(value="")
    sections_container_ref = {"frame": None}
    section_uis = []

    def do_undo():
        prev = undo.undo()
        if prev:
            job.clear()
            job.update(prev)
            ensure_job_defaults(job)
            rebuild_job_info_ui()
            rebuild_cost_code_ui()
            refresh_quotes_ui(full=True, force=True)
            maybe_refresh_frame_ui(force=True)
            schedule_bid_sheet_refresh()
            save_job(job)
            refresh_main()

    def snapshot():
        undo.snapshot(job)

    def is_bid_sheet_active():
        tab_bs = bid_sheet_tab_ref["tab"]
        return tab_bs is not None and nb.select() == str(tab_bs)

    def schedule_bid_sheet_refresh(delay_ms=50):
        if bid_sheet_refresh_fn["fn"] is None:
            return
        existing = bid_sheet_after_id["id"]
        if existing:
            try:
                tab.after_cancel(existing)
            except Exception:
                pass
        bid_sheet_after_id["id"] = tab.after(delay_ms, _run_bid_sheet_refresh)

    def _run_bid_sheet_refresh():
        bid_sheet_after_id["id"] = None
        if bid_sheet_refresh_fn["fn"] is None:
            return
        if not is_bid_sheet_active():
            return
        bid_sheet_refresh_fn["fn"]()

    def mark_frame_dirty():
        nonlocal frame_dirty
        frame_dirty = True
        bid_sheet_dirty["value"] = True
        if is_bid_sheet_active():
            schedule_bid_sheet_refresh()

    def maybe_refresh_frame_ui(force=False):
        nonlocal frame_dirty
        if not frame_initialized:
            return
        if force or frame_dirty:
            refresh_frame_schedule_ui()
            frame_dirty = False

    # --------------------------------------------------
    # Tab: Job Info
    # --------------------------------------------------
    tab_info = ttk.Frame(nb)
    nb.add(tab_info, text="Job Info")

    info_scroll = ScrollFrame(tab_info)
    info_scroll.pack(fill="both", expand=True, padx=10, pady=10)

    info_entries = {}

    def rebuild_job_info_ui():
        for w in info_scroll.inner.winfo_children():
            w.destroy()

        left = ttk.Frame(info_scroll.inner)
        right = ttk.Frame(info_scroll.inner)
        left.grid(row=0, column=0, sticky="nw", padx=(0, 20))
        right.grid(row=0, column=1, sticky="nw")
        info_scroll.inner.columnconfigure(0, weight=1)
        info_scroll.inner.columnconfigure(1, weight=1)

        half = (len(JOB_FIELDS) + 1) // 2
        left_fields = JOB_FIELDS[:half]
        right_fields = JOB_FIELDS[half:]

        info_entries.clear()

        def build_side(parent, fields, start_row=0):
            r = start_row
            for label, key, typ in fields:
                ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", pady=3)
                if typ == "bool":
                    var = tk.BooleanVar(value=bool(job.get(key, False)))
                    cb = ttk.Checkbutton(parent, variable=var, command=lambda k=key, v=var: set_field(k, v.get()))
                    cb.grid(row=r, column=1, sticky="w", pady=3)
                    info_entries[key] = ("bool", var, cb)
                elif typ == "date":
                    w = make_date_widget(parent, job.get(key, ""))
                    w.grid(row=r, column=1, sticky="w", pady=3)
                    info_entries[key] = ("date", w)
                else:
                    e = ttk.Entry(parent, width=34)
                    e.insert(0, job.get(key, ""))
                    e.grid(row=r, column=1, sticky="w", pady=3)
                    info_entries[key] = ("text", e)
                r += 1
            return r

        l_end = build_side(left, left_fields, 0)
        r_end = build_side(right, right_fields, 0)

        ttk.Button(left, text="Save Job Info", command=save_job_info).grid(
            row=l_end + 1, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )
        ttk.Button(left, text="Export PDF", command=lambda: export_job_info_pdf(job)).grid(
            row=l_end + 2, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

    def set_field(key, value):
        snapshot()
        job[key] = value
        save_job(job)
        refresh_main()
        bid_sheet_dirty["value"] = True

    def save_job_info():
        snapshot()
        for key, spec in info_entries.items():
            typ = spec[0]
            if typ == "bool":
                var = spec[1]
                job[key] = bool(var.get())
            elif typ == "date":
                widget = spec[1]
                job[key] = get_date_value(widget)
            else:
                e = spec[1]
                job[key] = e.get().strip()

        if not (job.get("status") or "").strip():
            job["status"] = DEFAULT_STATUS

        notebook.tab(tab, text=(job.get("job_name") or "Job"))
        save_job(job)
        refresh_main()
        bid_sheet_dirty["value"] = True
        messagebox.showinfo("Saved", "Job info saved.")

    rebuild_job_info_ui()

    # --------------------------------------------------
    # Tab: Cost Codes
    # --------------------------------------------------
    tab_cc = ttk.Frame(nb)
    nb.add(tab_cc, text="Cost Codes")

    cc_wrap = ttk.Frame(tab_cc)
    cc_wrap.pack(fill="both", expand=True, padx=10, pady=10)

    cc_tree = ttk.Treeview(cc_wrap, columns=("desc", "alts"), show="headings", height=14)
    cc_tree.heading("desc", text="Description")
    cc_tree.heading("alts", text="Alternates (comma list, 1..25)")
    cc_tree.column("desc", width=380, anchor="w")
    cc_tree.column("alts", width=240, anchor="w")
    cc_tree.pack(fill="both", expand=True)

    cc_controls = ttk.Frame(cc_wrap)
    cc_controls.pack(fill="x", pady=8)
    ttk.Button(cc_controls, text="+ Add Row", command=lambda: add_cc_row()).pack(side="left", padx=2)
    ttk.Button(cc_controls, text="- Delete Selected", command=lambda: del_cc_row()).pack(side="left", padx=2)
    ttk.Button(cc_controls, text="Save Cost Codes", command=lambda: save_cost_codes()).pack(side="left", padx=8)
    ttk.Button(cc_controls, text="Export PDF", command=lambda: export_cost_codes_pdf(job)).pack(side="left", padx=2)

    cc_editor = {"widget": None}

    def rebuild_cost_code_ui():
        cc_tree.delete(*cc_tree.get_children())
        for cc in sorted(job.get("cost_codes", []), key=lambda x: (x.get("code") or "")):
            code = cc.get("code", "00 00 00")
            desc = cc.get("description", "")
            alts = ",".join(str(n) for n in (cc.get("alts", []) or []))
            cc_tree.insert("", "end", iid=cc["id"], text="", values=(code, desc, alts))

    def add_cc_row():
        snapshot()
        cc = {"id": str(uuid.uuid4()), "code": "00 00 00", "description": "", "alts": []}
        job["cost_codes"].append(cc)
        normalize_quotes(job)
        normalize_bid_sheet(job)
        normalize_frame_schedules(job)
        save_job(job)
        rebuild_cost_code_ui()
        refresh_quotes_ui(full=True)
        mark_frame_dirty()
        refresh_main()

    def del_cc_row():
        sel = cc_tree.selection()
        if not sel:
            return
        snapshot()
        ids = set(sel)
        job["cost_codes"] = [cc for cc in job["cost_codes"] if cc.get("id") not in ids]
        normalize_quotes(job)
        normalize_bid_sheet(job)
        normalize_frame_schedules(job)
        save_job(job)
        rebuild_cost_code_ui()
        refresh_quotes_ui(full=True)
        mark_frame_dirty()
        refresh_main()

    def _cc_start_edit(item, col):
        if cc_editor["widget"]:
            cc_editor["widget"].destroy()
            cc_editor["widget"] = None
        bbox = cc_tree.bbox(item, col)
        if not bbox:
            return
        x, y, w, h = bbox
        e = ttk.Entry(cc_tree)
        e.place(x=x, y=y, width=w, height=h)
        current = cc_tree.set(item, col)
        e.insert(0, current)
        e.focus_set()
        cc_editor["widget"] = e

        def commit(_=None):
            if not e.winfo_exists():
                return
            newv = e.get().strip()
            snapshot()
            for cc in job["cost_codes"]:
                if cc.get("id") == item:
                    if col == "desc":
                        cc["description"] = newv
                    elif col == "alts":
                        cc["alts"] = parse_alts(newv)
                    else:  # code
                        cc["code"] = newv
                    break
            normalize_quotes(job)
            normalize_bid_sheet(job)
            normalize_frame_schedules(job)
            save_job(job)
            rebuild_cost_code_ui()
            refresh_quotes_ui(full=True)
            mark_frame_dirty()
            refresh_main()
            e.destroy()
            cc_editor["widget"] = None

        e.bind("<Return>", commit)
        e.bind("<FocusOut>", commit)

    def _cc_dclick(event):
        item = cc_tree.identify_row(event.y)
        col = cc_tree.identify_column(event.x)
        if not item:
            return
        col_map = {"#1": "code", "#2": "desc", "#3": "alts"}
        if col in col_map:
            _cc_start_edit(item, col_map[col])

    cc_tree.bind("<Double-Button-1>", _cc_dclick)

    def save_cost_codes():
        snapshot()
        normalize_quotes(job)
        normalize_bid_sheet(job)
        normalize_frame_schedules(job)
        save_job(job)
        refresh_quotes_ui(full=True)
        mark_frame_dirty()
        refresh_main()
        messagebox.showinfo("Saved", "Cost codes saved.")

    rebuild_cost_code_ui()

    # --------------------------------------------------
    # Tab: Quotes
    # --------------------------------------------------
    tab_q = ttk.Frame(nb)
    nb.add(tab_q, text="Quotes")

    q_top = ttk.Frame(tab_q)
    q_top.pack(fill="x", padx=10, pady=(8, 4))
    ttk.Button(q_top, text="Export PDF", command=lambda: export_quotes_pdf(job)).pack(side="left")

    quotes_scroll = ScrollFrame(tab_q)
    quotes_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def ensure_quote_defaults(q: dict):
        q.setdefault("id", str(uuid.uuid4()))
        q.setdefault("date", today_str())
        q.setdefault("vendors", "")
        q.setdefault("price", "")
        q.setdefault("surcharge_pct", "")
        q.setdefault("cost", "")
        q.setdefault("notes", "")

    def recalc_quote(q: dict):
        price = safe_int(q.get("price", ""))
        sur = safe_float(q.get("surcharge_pct", ""))
        q["cost"] = str(calc_cost(price, sur))

    def summary_total_avg(quotes_list):
        costs = [safe_int(q.get("cost", 0)) for q in quotes_list if safe_int(q.get("cost", 0)) > 0]
        total = sum(costs)
        avg = roundup(total / len(costs)) if costs else 0
        return total, avg

    def delete_quote_row(spec_key, row_id):
        snapshot()
        code, variant = spec_key
        qlist = job["quotes"][code][variant]
        qlist[:] = [q for q in qlist if q.get("id") != row_id]
        if len(qlist) == 0:
            qlist.append({})
        normalize_quotes(job)
        save_job(job)
        refresh_quotes_ui(full=False)
        handle_quote_change(code, variant)
        refresh_main()

    def create_quote_row(spec_key, parent, idx, qobj, code, variant):
        ensure_quote_defaults(qobj)

        date_w = make_date_widget(parent, qobj.get("date", today_str()))
        date_w.grid(row=idx + 1, column=0, padx=4, pady=2, sticky="w")

        vendor = ttk.Entry(parent, width=22)
        vendor.insert(0, qobj.get("vendors", ""))
        vendor.grid(row=idx + 1, column=1, padx=4, pady=2, sticky="w")

        price = ttk.Entry(parent, width=12)
        price.insert(0, qobj.get("price", ""))
        price.grid(row=idx + 1, column=2, padx=4, pady=2, sticky="w")

        sur = ttk.Entry(parent, width=10)
        sur.insert(0, qobj.get("surcharge_pct", ""))
        sur.grid(row=idx + 1, column=3, padx=4, pady=2, sticky="w")

        cost_var = tk.StringVar(value=money_fmt(qobj.get("cost", 0)))
        cost_lbl = ttk.Label(parent, textvariable=cost_var, width=12)
        cost_lbl.grid(row=idx + 1, column=4, padx=4, pady=2, sticky="w")

        notes = ttk.Entry(parent, width=36)
        notes.insert(0, qobj.get("notes", ""))
        notes.grid(row=idx + 1, column=5, padx=4, pady=2, sticky="w")

        del_btn = ttk.Button(parent, text="×", width=2,
                             command=lambda: delete_quote_row(spec_key, qobj["id"]))
        del_btn.grid(row=idx + 1, column=6, padx=2, pady=2)

        def commit(_=None):
            qobj["date"] = get_date_value(date_w)
            qobj["vendors"] = vendor.get().strip()
            qobj["price"] = str(safe_int(price.get().strip() or "0"))
            qobj["surcharge_pct"] = str(safe_float(sur.get().strip() or "0"))
            qobj["notes"] = notes.get().strip()
            recalc_quote(qobj)
            cost_var.set(money_fmt(qobj.get("cost", 0)))
            total, avg = summary_total_avg(job["quotes"][code][variant])
            ui = quote_ui.get(spec_key)
            if ui and ui.get("total_label"):
                ui["total_label"].configure(text=f"Total {money_fmt(total)} · Avg {money_fmt(avg)}")
            save_job(job)
            handle_quote_change(code, variant)
            refresh_main()

        for w in [date_w, vendor, price, sur, notes]:
            w.bind("<FocusOut>", commit)
            w.bind("<Return>", commit)

        return {
            "id": qobj["id"],
            "date_w": date_w,
            "vendor": vendor,
            "price": price,
            "sur": sur,
            "cost_var": cost_var,
            "notes": notes,
            "del_btn": del_btn,
            "commit": commit,
        }

    def add_quote_row(spec_key, code, variant):
        snapshot()
        q = {}
        ensure_quote_defaults(q)
        job["quotes"][code][variant].append(q)
        save_job(job)
        rebuild_spec_block(spec_key)
        handle_quote_change(code, variant)

    def toggle_spec_expand(spec_key):
        code, variant = spec_key
        # store state in-memory only
        expanded = quote_ui.get(spec_key, {}).get("expanded", True)
        new_state = not expanded
        quote_ui.setdefault(spec_key, {})["expanded"] = new_state
        rebuild_spec_block(spec_key)

    def build_spec_block(parent, code, desc, variant):
        spec_key = (code, variant)
        quotes_list = job["quotes"][code][variant]
        for q in quotes_list:
            ensure_quote_defaults(q)

        head = ttk.Frame(parent)
        head.pack(fill="x", pady=(8, 2))

        expanded = quote_ui.get(spec_key, {}).get("expanded", True)
        btn_txt = "▾" if expanded else "▸"
        ttk.Button(head, text=btn_txt, width=2, command=lambda: toggle_spec_expand(spec_key)).pack(side="left")

        title = frame_spec_label(code, variant)
        ttk.Label(head, text=f"{title} — {desc}", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(6, 8))
        ttk.Button(head, text="+ Add Quote", command=lambda: add_quote_row(spec_key, code, variant)).pack(side="left")

        is_collapsed = not expanded
        if is_collapsed:
            total, avg = summary_total_avg(quotes_list)
            ttk.Label(parent, text=f"Total {money_fmt(total)} · Avg {money_fmt(avg)}", foreground="gray")\
                .pack(anchor="w", padx=28, pady=(2, 6))
            quote_ui.pop(spec_key, None)
            return

        grid = ttk.Frame(parent)
        grid.pack(fill="x", padx=22, pady=(2, 6))

        headers = ["DATE", "VENDOR(S)", "PRICE", "SUR %", "COST", "NOTES", ""]
        widths = [12, 18, 12, 10, 12, 30, 2]
        for c, h in enumerate(headers):
            ttk.Label(grid, text=h).grid(row=0, column=c, padx=4, sticky="w")

        rows = []
        for idx, q in enumerate(quotes_list):
            ensure_quote_defaults(q)
            rows.append(create_quote_row(spec_key, grid, idx, q, code, variant))

        total, avg = summary_total_avg(quotes_list)
        total_label = ttk.Label(parent, text=f"Total {money_fmt(total)} · Avg {money_fmt(avg)}", foreground="gray")
        total_label.pack(anchor="w", padx=28, pady=(0, 2))

        quote_ui[spec_key] = {
            "grid": grid,
            "rows": rows,
            "total_label": total_label,
        }

    def rebuild_spec_block(spec_key):
        frame = spec_frames.get(spec_key)
        if not frame:
            refresh_quotes_ui(full=True)
            return
        for w in frame.winfo_children():
            w.destroy()
        code, variant = spec_key
        desc = ""
        for cc in job["cost_codes"]:
            if (cc.get("code") or "").strip() == code:
                desc = cc.get("description", "")
                break
        build_spec_block(frame, code, desc, variant)

    def ensure_quotes_initialized():
        nonlocal quotes_initialized
        if quotes_initialized:
            return
        quotes_initialized = True
        refresh_quotes_ui(full=True, force=True)

    def refresh_quotes_ui(full=True, force=False):
        nonlocal quotes_dirty
        if not quotes_initialized and not force:
            quotes_dirty = True
            return
        normalize_quotes(job)
        normalize_bid_sheet(job)
        if full:
            for w in quotes_scroll.inner.winfo_children():
                w.destroy()
            spec_frames.clear()

        ordered = []
        for cc in sorted(job["cost_codes"], key=lambda x: (x.get("code") or "")):
            code = (cc.get("code") or "").strip()
            if not code:
                continue
            desc = cc.get("description", "")
            for v in variants_for_cc(cc.get("alts", [])):
                ordered.append((code, desc, v))

        for code, desc, variant in ordered:
            spec_key = (code, variant)
            if spec_key not in spec_frames:
                frame = ttk.Frame(quotes_scroll.inner)
                frame.pack(fill="x", anchor="w")
                spec_frames[spec_key] = frame
                build_spec_block(frame, code, desc, variant)
        bid_sheet_dirty["value"] = True
        quotes_dirty = False
        if is_bid_sheet_active():
            schedule_bid_sheet_refresh()

    # --------------------------------------------------
    # Tab: Bid Sheet
    # --------------------------------------------------
    tab_bs = ttk.Frame(nb)
    nb.add(tab_bs, text="Bid Sheet")

    bs_scroll = ScrollFrame(tab_bs)
    bs_scroll.pack(fill="both", expand=True, padx=10, pady=10)

    bs_topbar = ttk.Frame(bs_scroll.inner)
    bs_topbar.pack(fill="x", pady=(0, 8))
    ttk.Label(
        bs_topbar,
        text="Bid Sheet overview and markup editing. Double-click markup, notes, or color to edit.",
    ).pack(side="left")

    bs_overview = ttk.LabelFrame(bs_scroll.inner, text="Job Overview")
    bs_overview.pack(fill="x", pady=(0, 8))
    overview_fields = [
        ("Job Name", "job_name"),
        ("Project Number", "project_number"),
        ("Estimator", "estimator"),
        ("Project Manager", "project_manager"),
        ("Bid Due Date", "bid_due_date"),
        ("Status", "status"),
        ("Client", "owner_name"),
        ("Address", "job_address"),
    ]
    overview_vars = {}
    for idx, (label, key) in enumerate(overview_fields):
        row = idx // 2
        col = (idx % 2) * 2
        ttk.Label(bs_overview, text=f"{label}:", width=16).grid(row=row, column=col, sticky="w", padx=(8, 4), pady=3)
        var = tk.StringVar(value="")
        ttk.Label(bs_overview, textvariable=var).grid(row=row, column=col + 1, sticky="w", padx=(0, 10), pady=3)
        overview_vars[key] = var

    bs_tree = ttk.Treeview(
        bs_scroll.inner,
        columns=("desc", "cost", "markup_pct", "markup_amt", "sov", "notes", "color"),
        show="tree headings",
        height=18,
    )
    bs_tree.pack(fill="both", expand=True)

    bs_tree.heading("#0", text="CODE")
    bs_tree.heading("desc", text="SPECIFICATION / SUBTOTAL")
    bs_tree.heading("cost", text="BASE COST")
    bs_tree.heading("markup_pct", text="MARKUP PERCENT (%)")
    bs_tree.heading("markup_amt", text="MARKUP DOLLARS ($)")
    bs_tree.heading("sov", text="TOTAL SOV")
    bs_tree.heading("notes", text="NOTES")
    bs_tree.heading("color", text="COLOR")

    bs_tree.column("#0", width=120, anchor="w")
    bs_tree.column("desc", width=360, anchor="w")
    bs_tree.column("cost", width=120, anchor="e")
    bs_tree.column("markup_pct", width=140, anchor="e")
    bs_tree.column("markup_amt", width=150, anchor="e")
    bs_tree.column("sov", width=130, anchor="e")
    bs_tree.column("notes", width=220, anchor="w")
    bs_tree.column("color", width=90, anchor="center")

    COLOR_OPTIONS = ["None", "Yellow", "Green", "Red", "Blue", "Gray"]
    COLOR_MAP = {
        "None": "",
        "Yellow": "#fff3a0",
        "Green": "#c8f7c5",
        "Red": "#f7c5c5",
        "Blue": "#c5d9f7",
        "Gray": "#e0e0e0",
    }

    bid_sheet_item_meta = {}
    bid_sheet_line_index = {}
    bs_sort_state = {"col": None, "descending": False}
    bs_edit_widget = {"widget": None}

    def _clear_bs_tree():
        for item in bs_tree.get_children():
            bs_tree.delete(item)
        bid_sheet_item_meta.clear()
        bid_sheet_line_index.clear()

    def _refresh_overview_values():
        for _label, key in overview_fields:
            overview_vars[key].set((job.get(key) or "").strip())

    def _apply_color_tag(item_id: str, color_name: str):
        color = COLOR_MAP.get(color_name or "None", "")
        if color:
            tag = f"color_{color_name}"
            if not bs_tree.tag_has(tag):
                bs_tree.tag_configure(tag, background=color)
            bs_tree.item(item_id, tags=(tag,))
        else:
            bs_tree.item(item_id, tags=())

    def _format_markup_values(pct_val: float, amt_val: int):
        pct_display = pct_fmt(pct_val) if pct_val > 0 else ""
        amt_display = money_fmt(amt_val) if amt_val > 0 else ""
        return pct_display, amt_display

    def _calc_markup_from_source(cost_value: int, data: dict):
        source = (data.get("markup_source") or "pct").strip().lower()
        pct_val = parse_pct(data.get("markup_pct", ""))
        amt_val = parse_money(data.get("markup_amt", ""))
        if source == "amt":
            amt_val = max(0, amt_val)
            pct_val = (amt_val / cost_value * 100.0) if cost_value and amt_val > 0 else 0.0
        else:
            pct_val = max(0.0, pct_val)
            amt_val = roundup(cost_value * (pct_val / 100.0)) if pct_val > 0 else 0
            source = "pct"
        data["markup_source"] = source
        return pct_val, amt_val

    def _sort_children(parent_id: str, col: str, descending: bool):
        children = list(bs_tree.get_children(parent_id))

        def sort_key(item_id):
            if col == "#0":
                return (bs_tree.item(item_id, "text") or "").lower()
            values = bs_tree.item(item_id, "values")
            idx_map = {
                "desc": 0,
                "cost": 1,
                "markup_pct": 2,
                "markup_amt": 3,
                "sov": 4,
                "notes": 5,
                "color": 6,
            }
            idx = idx_map.get(col, 0)
            val = values[idx] if idx < len(values) else ""
            if col in ("cost", "markup_amt", "sov"):
                return parse_money(val)
            if col == "markup_pct":
                return parse_pct(val)
            return str(val).lower()

        children.sort(key=sort_key, reverse=descending)
        for idx, item_id in enumerate(children):
            bs_tree.move(item_id, parent_id, idx)
            _sort_children(item_id, col, descending)

    def _sort_by(col: str):
        descending = bs_sort_state["descending"] if bs_sort_state["col"] == col else False
        bs_sort_state["col"] = col
        bs_sort_state["descending"] = not descending
        _sort_children("", col, not descending)

    for col_id in ("#0", "desc", "cost", "markup_pct", "markup_amt", "sov", "notes", "color"):
        bs_tree.heading(col_id, command=lambda c=col_id: _sort_by(c))

    def _is_valid_markup_input(value: str, allow_decimal: bool) -> bool:
        if value == "":
            return True
        if value.startswith("-"):
            return False
        stripped = value.replace("$", "").replace("%", "").replace(",", "").strip()
        if stripped == "":
            return True
        dot_count = stripped.count(".")
        if dot_count > 1:
            return False
        if not allow_decimal and dot_count > 0:
            return False
        return all(ch.isdigit() or (allow_decimal and ch == ".") for ch in stripped)

    def _start_edit(item_id: str, column: str):
        if item_id not in bid_sheet_item_meta:
            return
        meta = bid_sheet_item_meta[item_id]
        if meta.get("row_type") != "line_item":
            return
        if column not in ("markup_pct", "markup_amt", "notes", "color"):
            return

        if bs_edit_widget["widget"]:
            bs_edit_widget["widget"].destroy()
            bs_edit_widget["widget"] = None

        bbox = bs_tree.bbox(item_id, column)
        if not bbox:
            return
        x, y, w, h = bbox
        parent = bs_tree

        if column == "color":
            widget = ttk.Combobox(parent, values=COLOR_OPTIONS, state="readonly")
            widget.place(x=x, y=y, width=w, height=h)
            widget.set(meta["data"].get("color", "None"))
        else:
            widget = ttk.Entry(parent)
            widget.place(x=x, y=y, width=w, height=h)
            widget.insert(0, bs_tree.set(item_id, column))
            if column == "markup_pct":
                widget.configure(validate="key", validatecommand=(parent.register(lambda v: _is_valid_markup_input(v, True)), "%P"))
            elif column == "markup_amt":
                widget.configure(validate="key", validatecommand=(parent.register(lambda v: _is_valid_markup_input(v, False)), "%P"))

        bs_edit_widget["widget"] = widget
        widget.focus_set()

        def _commit_markup():
            cost_value = meta["cost_value"]
            pct_val, amt_val = _calc_markup_from_source(cost_value, meta["data"])
            pct_display, amt_display = _format_markup_values(pct_val, amt_val)
            meta["data"]["markup_pct"] = pct_display
            meta["data"]["markup_amt"] = amt_display
            bs_tree.set(item_id, "markup_pct", pct_display)
            bs_tree.set(item_id, "markup_amt", amt_display)
            bs_tree.set(item_id, "sov", money_fmt(cost_value + amt_val))

        def commit(_event=None):
            if not widget.winfo_exists():
                return
            new_val = widget.get().strip()
            if column == "color":
                meta["data"]["color"] = new_val if new_val in COLOR_OPTIONS else "None"
                bs_tree.set(item_id, "color", meta["data"]["color"])
                _apply_color_tag(item_id, meta["data"]["color"])
            elif column == "notes":
                meta["data"]["notes"] = new_val
                bs_tree.set(item_id, "notes", new_val)
            elif column == "markup_pct":
                meta["data"]["markup_source"] = "pct"
                meta["data"]["markup_pct"] = new_val
                _commit_markup()
            elif column == "markup_amt":
                meta["data"]["markup_source"] = "amt"
                meta["data"]["markup_amt"] = new_val
                _commit_markup()

            widget.destroy()
            bs_edit_widget["widget"] = None

        def realtime_update(_event=None):
            if column not in ("markup_pct", "markup_amt"):
                return
            new_val = widget.get().strip()
            if column == "markup_pct":
                meta["data"]["markup_source"] = "pct"
                meta["data"]["markup_pct"] = new_val
            else:
                meta["data"]["markup_source"] = "amt"
                meta["data"]["markup_amt"] = new_val
            _commit_markup()

        widget.bind("<Return>", commit)
        widget.bind("<FocusOut>", commit)
        if column in ("markup_pct", "markup_amt"):
            widget.bind("<KeyRelease>", realtime_update)

    def _on_double_click(event):
        item_id = bs_tree.identify_row(event.y)
        column = bs_tree.identify_column(event.x)
        if not item_id:
            return
        col_map = {
            "#2": "desc",
            "#3": "cost",
            "#4": "markup_pct",
            "#5": "markup_amt",
            "#6": "sov",
            "#7": "notes",
            "#8": "color",
        }
        col_id = col_map.get(column)
        if col_id:
            _start_edit(item_id, col_id)

    bs_tree.bind("<Double-Button-1>", _on_double_click)

    bid_sheet_tab_ref["tab"] = tab_bs

    def _on_tab_changed(_event=None):
        selected = nb.select()
        if selected == str(tab_q):
            ensure_quotes_initialized()
            if quotes_dirty:
                refresh_quotes_ui(full=True, force=True)
        if selected == str(tab_fs):
            ensure_frame_initialized()
            if frame_dirty:
                maybe_refresh_frame_ui()
        if is_bid_sheet_active() and bid_sheet_dirty["value"]:
            schedule_bid_sheet_refresh()

    nb.bind("<<NotebookTabChanged>>", _on_tab_changed)

    def refresh_bid_sheet_ui():
        normalize_quotes(job)
        normalize_bid_sheet(job)
        normalize_frame_schedules(job)
        compute_frame_schedule_rollups(job)
        _clear_bs_tree()
        _refresh_overview_values()

        for cc in sorted(job["cost_codes"], key=lambda x: (x.get("code") or "")):
            code = (cc.get("code") or "").strip()
            if not code:
                continue
            desc = cc.get("description", "")

            for variant in variants_for_cc(cc.get("alts", [])):
                spec_id = frame_spec_id(code, variant)
                data = job["bid_sheet"].setdefault(spec_id, {})
                data.setdefault("markup_source", "pct")
                base_cost = bid_sheet_direct_cost(job, code, variant)
                install_total = bid_sheet_install_material_total(job, spec_id)
                cost_value = base_cost + install_total

                pct_val, amt_val = _calc_markup_from_source(cost_value, data)
                pct_display, amt_display = _format_markup_values(pct_val, amt_val)
                data["markup_pct"] = pct_display
                data["markup_amt"] = amt_display

                spec_label = desc if variant == "BASE" else f"{desc} ({variant})"
                line_id = bs_tree.insert(
                    "",
                    "end",
                    text=code,
                    values=(
                        spec_label,
                        money_fmt(cost_value),
                        pct_display,
                        amt_display,
                        money_fmt(cost_value + amt_val),
                        data.get("notes", ""),
                        data.get("color", "None"),
                    ),
                    open=True,
                )
                bid_sheet_item_meta[line_id] = {
                    "row_type": "line_item",
                    "spec_id": spec_id,
                    "data": data,
                    "cost_value": cost_value,
                }
                bid_sheet_line_index[spec_id] = line_id
                _apply_color_tag(line_id, data.get("color", "None"))

                bs_tree.insert(
                    line_id,
                    "end",
                    text="",
                    values=("Base Product", money_fmt(base_cost), "", "", money_fmt(base_cost), "", ""),
                )
                install_id = bs_tree.insert(
                    line_id,
                    "end",
                    text="",
                    values=("Install Material", money_fmt(install_total), "", "", money_fmt(install_total), "", ""),
                    open=False,
                )
                bid_sheet_item_meta[line_id]["install_id"] = install_id

                for label, cost in bid_sheet_material_breakdown(job, spec_id):
                    bs_tree.insert(
                        install_id,
                        "end",
                        text="",
                        values=(label, money_fmt(cost), "", "", money_fmt(cost), "", ""),
                    )
        bid_sheet_dirty["value"] = False

    bid_sheet_refresh_fn["fn"] = refresh_bid_sheet_ui

    def update_bid_sheet_line_item(spec_id: str):
        line_id = bid_sheet_line_index.get(spec_id)
        if not line_id:
            return
        meta = bid_sheet_item_meta.get(line_id, {})
        data = meta.get("data", {})
        base, variant = parse_frame_spec_id(spec_id)
        base_cost = bid_sheet_direct_cost(job, base, variant)
        install_total = bid_sheet_install_material_total(job, spec_id)
        cost_value = base_cost + install_total
        meta["cost_value"] = cost_value
        pct_val, amt_val = _calc_markup_from_source(cost_value, data)
        pct_display, amt_display = _format_markup_values(pct_val, amt_val)
        data["markup_pct"] = pct_display
        data["markup_amt"] = amt_display
        bs_tree.set(line_id, "cost", money_fmt(cost_value))
        bs_tree.set(line_id, "markup_pct", pct_display)
        bs_tree.set(line_id, "markup_amt", amt_display)
        bs_tree.set(line_id, "sov", money_fmt(cost_value + amt_val))
        children = bs_tree.get_children(line_id)
        if len(children) >= 1:
            bs_tree.item(children[0], values=("Base Product", money_fmt(base_cost), "", "", money_fmt(base_cost), "", ""))
        if len(children) >= 2:
            bs_tree.item(children[1], values=("Install Material", money_fmt(install_total), "", "", money_fmt(install_total), "", ""))

    def handle_quote_change(code: str, variant: str):
        spec_id = frame_spec_id(code, variant)
        bid_sheet_dirty["value"] = True
        if not is_bid_sheet_active():
            return
        if bid_sheet_refresh_fn["fn"] is None:
            return
        if spec_id in bid_sheet_line_index:
            update_bid_sheet_line_item(spec_id)
            return
        schedule_bid_sheet_refresh()

    # --------------------------------------------------
    # Tab: Frame Schedule  (FIXED: add section never tears down existing UIs)
    # --------------------------------------------------
    tab_fs = ttk.Frame(nb)
    nb.add(tab_fs, text="Frame Schedule")

    fs_scroll = ScrollFrame(tab_fs)
    fs_scroll.pack(fill="both", expand=True, padx=10, pady=10)

    selector_frame = ttk.Frame(fs_scroll.inner)
    selector_frame.pack(anchor="w", pady=6, fill="x")

    ttk.Label(selector_frame, text="Select Spec:").pack(side="left")
    spec_combo = ttk.Combobox(selector_frame, state="readonly", width=40, textvariable=selected_spec_label_var)
    spec_combo.pack(side="left", padx=8)
    add_section_btn = ttk.Button(selector_frame, text="+ Add Section", command=lambda: add_section_ui())
    add_section_btn.pack(side="left", padx=4)
    ttk.Button(selector_frame, text="Export PDF", command=lambda: export_frame_schedule_pdf(job)).pack(side="left", padx=4)
    frame_schedule_error = ttk.Label(selector_frame, textvariable=frame_schedule_error_var, foreground="#b00020")
    frame_schedule_error.pack(side="left", padx=8)

    sections_wrap = ttk.Frame(fs_scroll.inner)
    sections_wrap.pack(fill="both", expand=True)
    sections_container_ref["frame"] = sections_wrap

    class SectionUI:
        def __init__(self, master, section: dict, spec_id: str):
            self.master = master
            self.section = section
            self.spec_id = spec_id
            self.row_widgets = []
            self.material_widgets = []
            self.subtotal_row_index = None

            self.outer = ttk.LabelFrame(master, text=section.get("name", "Section"))
            self.outer.pack(fill="x", pady=8)

            head = ttk.Frame(self.outer)
            head.pack(fill="x", padx=6, pady=(4, 2))
            ttk.Label(head, text="Section Name").pack(side="left")
            self.name_entry = ttk.Entry(head, width=30)
            self.name_entry.insert(0, section.get("name", "Section"))
            self.name_entry.pack(side="left", padx=6)
            ttk.Button(head, text="Delete Section", command=self.delete_section).pack(side="right")
            self.name_entry.bind("<FocusOut>", self.save_name)
            self.name_entry.bind("<Return>", self.save_name)

            self.grid = ttk.Frame(self.outer)
            self.grid.pack(fill="x", padx=6, pady=4)

            # headers
            for c, (h, key, w, kind) in enumerate(FS_HEADERS):
                ttk.Label(self.grid, text=h).grid(row=0, column=c, padx=3, sticky="w")

            ttk.Button(self.outer, text="+ Add Row", command=self.add_row).pack(anchor="w", padx=6, pady=(0, 6))

            # Build rows
            for r in self.section["rows"]:
                self._add_row_widgets(r)

            self._render_subtotal_row()

            # Material table
            sep = ttk.Separator(self.outer, orient="horizontal")
            sep.pack(fill="x", padx=6, pady=6)

            mat_head = ttk.Frame(self.outer)
            mat_head.pack(fill="x", padx=6)
            ttk.Label(mat_head, text="INSTALL MATERIAL BREAKDOWN", font=("Segoe UI", 9, "bold")).pack(anchor="w")

            self.mat_grid = ttk.Frame(self.outer)
            self.mat_grid.pack(fill="x", padx=6, pady=(4, 6))

            mh = ["Material", "Basis", "Factor", "Rate", "Qty", "Cost", "Unit"]
            for c, h in enumerate(mh):
                ttk.Label(self.mat_grid, text=h).grid(row=0, column=c, padx=3, sticky="w")

            self._build_material_rows()
            self.update_subtotals_and_materials()

        def save_name(self, _=None):
            new_name = self.name_entry.get().strip() or "Section"
            if new_name != self.section.get("name", "Section"):
                snapshot()
                self.section["name"] = new_name
                self.outer.configure(text=new_name)
                save_job(job)
                mark_frame_dirty()

        def delete_section(self):
            if not messagebox.askyesno("Delete Section", "Delete this section?"):
                return
            snapshot()
            sections = list(job["frame_schedules"].get(self.spec_id, []))
            job["frame_schedules"][self.spec_id] = [
                s for s in sections if s.get("id") != self.section.get("id")
            ]
            save_job(job)
            refresh_frame_schedule_ui()
            mark_frame_dirty()

        def _is_last_row(self, rdata: dict) -> bool:
            rows = self.section.get("rows", [])
            return bool(rows) and rows[-1] is rdata

        def _append_blank_row(self, persist: bool = True):
            r = blank_frame_schedule_row()
            self.section["rows"] = list(self.section.get("rows", [])) + [r]
            self._add_row_widgets(r)
            if persist:
                save_job(job)
                mark_frame_dirty()

        def _maybe_append_blank_row_for_qty(self, rdata: dict):
            qty_text = str(rdata.get("qty", "")).strip()
            if qty_text and self._is_last_row(rdata):
                self._append_blank_row(persist=False)

        def _row_commit(self, rdata: dict, fields: dict, persist: bool = True):
            rdata["spec_mark"] = fields["spec_mark"].get().strip()
            rdata["qty"] = fields["qty"].get().strip()
            rdata["width"] = fields["width"].get().strip()
            rdata["height"] = fields["height"].get().strip()
            recalc_row_fields(rdata)
            fields["sqft_var"].set(str(safe_int(rdata["sqft"])))
            fields["perim_var"].set(str(safe_int(rdata["perim"])))
            fields["caulk_var"].set(str(safe_int(rdata["caulk_lf"])))
            fields["hs_var"].set(str(safe_int(rdata["head_sill"])))

            self._maybe_append_blank_row_for_qty(rdata)

            if persist:
                save_job(job)
            self.update_subtotals_and_materials()
            if persist:
                mark_frame_dirty()

        def _remove_subtotal_row_widgets(self):
            if self.subtotal_row_index is None:
                return
            for w in list(self.grid.grid_slaves(row=self.subtotal_row_index)):
                w.destroy()
            self.subtotal_row_index = None

        def _render_subtotal_row(self):
            self._remove_subtotal_row_widgets()
            self.subtotal_row_index = len(self.row_widgets) + 1
            r = self.subtotal_row_index
            ttk.Label(self.grid, text="Subtotal", font=("Segoe UI", 9, "bold")).grid(row=r, column=0, padx=3, sticky="w")
            self.sub_vars = {
                "qty": tk.StringVar(value="0"),
                "sqft": tk.StringVar(value="0"),
                "perim": tk.StringVar(value="0"),
                "caulk_lf": tk.StringVar(value="0"),
                "head_sill": tk.StringVar(value="0"),
            }
            ttk.Label(self.grid, textvariable=self.sub_vars["qty"], font=("Segoe UI", 9, "bold")).grid(row=r, column=1, padx=3, sticky="e")
            ttk.Label(self.grid, text="").grid(row=r, column=2, padx=3, sticky="e")
            ttk.Label(self.grid, text="").grid(row=r, column=3, padx=3, sticky="e")
            ttk.Label(self.grid, textvariable=self.sub_vars["sqft"], font=("Segoe UI", 9, "bold")).grid(row=r, column=4, padx=3, sticky="e")
            ttk.Label(self.grid, textvariable=self.sub_vars["perim"], font=("Segoe UI", 9, "bold")).grid(row=r, column=5, padx=3, sticky="e")
            ttk.Label(self.grid, textvariable=self.sub_vars["caulk_lf"], font=("Segoe UI", 9, "bold")).grid(row=r, column=6, padx=3, sticky="e")
            ttk.Label(self.grid, textvariable=self.sub_vars["head_sill"], font=("Segoe UI", 9, "bold")).grid(row=r, column=7, padx=3, sticky="e")

        def _add_row_widgets(self, rdata: dict):
            normalize_fs_row(rdata)
            recalc_row_fields(rdata)

            self._remove_subtotal_row_widgets()

            rowi = len(self.row_widgets) + 1
            spec_mark = ttk.Entry(self.grid, width=18)
            qty_var_input = tk.StringVar(value=rdata.get("qty", ""))
            width_var_input = tk.StringVar(value=rdata.get("width", ""))
            height_var_input = tk.StringVar(value=rdata.get("height", ""))

            qty = ttk.Entry(self.grid, width=8, textvariable=qty_var_input)
            width = ttk.Entry(self.grid, width=8, textvariable=width_var_input)
            height = ttk.Entry(self.grid, width=8, textvariable=height_var_input)

            spec_mark.insert(0, rdata.get("spec_mark", ""))

            sqft_var = tk.StringVar(value=str(safe_int(rdata.get("sqft", 0))))
            perim_var = tk.StringVar(value=str(safe_int(rdata.get("perim", 0))))
            caulk_var = tk.StringVar(value=str(safe_int(rdata.get("caulk_lf", 0))))
            hs_var = tk.StringVar(value=str(safe_int(rdata.get("head_sill", 0))))

            spec_mark.grid(row=rowi, column=0, padx=3, pady=2, sticky="w")
            qty.grid(row=rowi, column=1, padx=3, pady=2, sticky="w")
            width.grid(row=rowi, column=2, padx=3, pady=2, sticky="w")
            height.grid(row=rowi, column=3, padx=3, pady=2, sticky="w")
            ttk.Label(self.grid, textvariable=sqft_var, width=8, anchor="e").grid(row=rowi, column=4, padx=3, pady=2, sticky="e")
            ttk.Label(self.grid, textvariable=perim_var, width=8, anchor="e").grid(row=rowi, column=5, padx=3, pady=2, sticky="e")
            ttk.Label(self.grid, textvariable=caulk_var, width=10, anchor="e").grid(row=rowi, column=6, padx=3, pady=2, sticky="e")
            ttk.Label(self.grid, textvariable=hs_var, width=10, anchor="e").grid(row=rowi, column=7, padx=3, pady=2, sticky="e")

            del_btn = ttk.Button(self.grid, text="×", width=2)
            del_btn.grid(row=rowi, column=8, padx=2, pady=2)

            fields = {
                "spec_mark": spec_mark,
                "qty": qty,
                "width": width,
                "height": height,
                "qty_var_input": qty_var_input,
                "width_var_input": width_var_input,
                "height_var_input": height_var_input,
                "sqft_var": sqft_var,
                "perim_var": perim_var,
                "caulk_var": caulk_var,
                "hs_var": hs_var,
                "del_btn": del_btn,
                "row_data": rdata,
            }

            def commit(_=None):
                self._row_commit(rdata, fields)

            def commit_live(_=None):
                self._row_commit(rdata, fields, persist=False)

            def live_trace_handler(*_):
                self._row_commit(rdata, fields, persist=False)

            for w in (spec_mark, qty, width, height):
                w.bind("<FocusOut>", commit)
                w.bind("<Return>", commit)
            for w in (qty, width, height):
                w.bind("<KeyRelease>", commit_live)

            qty_var_input.trace_add("write", live_trace_handler)
            width_var_input.trace_add("write", live_trace_handler)
            height_var_input.trace_add("write", live_trace_handler)

            def delete_row():
                snapshot()
                rows = self.section["rows"]
                remaining = [x for x in rows if x is not rdata]
                if not remaining:
                    remaining = [blank_frame_schedule_row()]
                self.section["rows"] = remaining
                save_job(job)
                self.rebuild_rows()
                self.update_subtotals_and_materials()
                mark_frame_dirty()

            del_btn.configure(command=delete_row)
            self.row_widgets.append(fields)
            self._render_subtotal_row()

        def rebuild_rows(self):
            # wipe existing row widgets
            for rw in self.row_widgets:
                for k in ("spec_mark", "qty", "width", "height", "del_btn"):
                    w = rw.get(k)
                    if w:
                        w.destroy()
                # calc labels are not stored as widgets; we can destroy row by row
                r_index = self.row_widgets.index(rw) + 1
                for w in list(self.grid.grid_slaves(row=r_index)):
                    w.destroy()
            self.row_widgets.clear()
            self._remove_subtotal_row_widgets()

            for r in self.section["rows"]:
                self._add_row_widgets(r)
            self._render_subtotal_row()

        def add_row(self):
            snapshot()
            r = blank_frame_schedule_row()
            self.section["rows"] = list(self.section.get("rows", [])) + [r]
            save_job(job)
            self._add_row_widgets(r)
            self.update_subtotals_and_materials()
            mark_frame_dirty()

        def _build_material_rows(self):
            # remove old
            for row in self.material_widgets:
                for w in row.values():
                    if hasattr(w, "destroy"):
                        w.destroy()
            self.material_widgets.clear()

            mats = self.section.get("materials", [])
            for i, m in enumerate(mats, start=1):
                label = ttk.Entry(self.mat_grid, width=34)
                label.insert(0, m.get("label", ""))

                basis = ttk.Combobox(self.mat_grid, state="readonly", width=20,
                                     values=["perim_subtotal", "head_sill_subtotal", "caulk_lf_subtotal", "manual"])
                basis.set(m.get("basis", "perim_subtotal"))

                factor = ttk.Entry(self.mat_grid, width=8)
                factor.insert(0, m.get("factor", "1.0"))

                rate = ttk.Entry(self.mat_grid, width=8)
                rate.insert(0, m.get("rate", "0"))

                qty_var = tk.StringVar(value="0.00")
                cost_var = tk.StringVar(value="$0")

                unit = ttk.Entry(self.mat_grid, width=12)
                unit.insert(0, m.get("unit", ""))

                label.grid(row=i, column=0, padx=3, pady=2, sticky="w")
                basis.grid(row=i, column=1, padx=3, pady=2, sticky="w")
                factor.grid(row=i, column=2, padx=3, pady=2, sticky="w")
                rate.grid(row=i, column=3, padx=3, pady=2, sticky="w")
                ttk.Label(self.mat_grid, textvariable=qty_var, width=10, anchor="e").grid(row=i, column=4, padx=3, pady=2, sticky="e")
                ttk.Label(self.mat_grid, textvariable=cost_var, width=10, anchor="e").grid(row=i, column=5, padx=3, pady=2, sticky="e")
                unit.grid(row=i, column=6, padx=3, pady=2, sticky="w")

                row = {
                    "label": label, "basis": basis, "factor": factor, "rate": rate,
                    "qty_var": qty_var, "cost_var": cost_var, "unit": unit, "mref": m
                }
                self.material_widgets.append(row)

                def commit_factory(rw=row):
                    def commit(_=None):
                        mref = rw["mref"]
                        mref["label"] = rw["label"].get().strip()
                        mref["basis"] = rw["basis"].get().strip()
                        mref["factor"] = rw["factor"].get().strip()
                        mref["rate"] = rw["rate"].get().strip()
                        mref["unit"] = rw["unit"].get().strip()
                        # keep qty only for manual basis
                        if mref.get("basis") == "manual":
                            # ask optional manual qty if empty remains as-is
                            pass
                        save_job(job)
                        self.update_subtotals_and_materials()
                        mark_frame_dirty()
                    return commit

                c = commit_factory()
                for w in (label, basis, factor, rate, unit):
                    w.bind("<FocusOut>", c)
                    w.bind("<Return>", c)

            self.install_total_var = tk.StringVar(value="$0")
            ttk.Label(self.mat_grid, text="Install Material Total", font=("Segoe UI", 9, "bold")).grid(
                row=len(mats)+1, column=4, padx=3, pady=(6, 2), sticky="e"
            )
            ttk.Label(self.mat_grid, textvariable=self.install_total_var, font=("Segoe UI", 9, "bold")).grid(
                row=len(mats)+1, column=5, padx=3, pady=(6, 2), sticky="e"
            )

        def update_subtotals_and_materials(self):
            subs = schedule_subtotals(self.section["rows"])
            self.sub_vars["qty"].set(str(subs["qty"]))
            self.sub_vars["sqft"].set(str(subs["sqft"]))
            self.sub_vars["perim"].set(str(subs["perim"]))
            self.sub_vars["caulk_lf"].set(str(subs["caulk_lf"]))
            self.sub_vars["head_sill"].set(str(subs["head_sill"]))

            install_total = 0
            for rw in self.material_widgets:
                m = rw["mref"]
                basis = rw["basis"].get().strip()
                m["basis"] = basis
                m["factor"] = rw["factor"].get().strip()
                m["rate"] = rw["rate"].get().strip()
                m["label"] = rw["label"].get().strip()
                m["unit"] = rw["unit"].get().strip()

                qty = material_qty_for_row(m, subs)
                rate = safe_float(m.get("rate"))
                cost = roundup(qty * rate) if qty * rate != 0 else 0
                rw["qty_var"].set(f"{qty:.2f}")
                rw["cost_var"].set(money_fmt(cost))
                install_total += cost

            self.install_total_var.set(money_fmt(install_total))

            recompute_section_totals(self.section)
            save_job(job)

    def frame_spec_options():
        opts = []
        seen_labels = set()
        for cc in sorted(job["cost_codes"], key=lambda x: (x.get("code") or "")):
            code = (cc.get("code") or "").strip()
            if not code:
                continue
            desc = cc.get("description", "")
            for v in variants_for_cc(cc.get("alts", [])):
                sid = frame_spec_id(code, v)
                lbl = f"{frame_spec_label(code, v)} — {desc}".strip()
                if lbl in seen_labels:
                    lbl = f"{lbl} [{sid}]"
                seen_labels.add(lbl)
                opts.append((sid, lbl))
        return opts

    def set_frame_schedule_error(message: str):
        frame_schedule_error_var.set(message)

    def update_add_section_button_state(spec_id: str):
        if is_valid_frame_spec_id(job, spec_id):
            add_section_btn.configure(state="normal")
        else:
            add_section_btn.configure(state="disabled")

    def resolve_current_spec_id(opts=None):
        options = opts if opts is not None else frame_spec_options()
        if not options:
            return ""

        id_to_label = {sid: lbl for sid, lbl in options}
        label_to_id = {lbl: sid for sid, lbl in options}

        current_id = current_spec_id_var.get().strip()
        if current_id in id_to_label:
            return current_id

        selected_label = selected_spec_label_var.get().strip() or spec_combo.get().strip()
        if selected_label in label_to_id:
            resolved_id = label_to_id[selected_label]
            current_spec_id_var.set(resolved_id)
            return resolved_id

        return ""

    def ensure_frame_initialized():
        nonlocal frame_initialized
        if frame_initialized:
            return
        frame_initialized = True
        refresh_frame_schedule_ui()

    def refresh_frame_schedule_ui():
        normalize_frame_schedules(job)
        compute_frame_schedule_rollups(job)

        for w in sections_wrap.winfo_children():
            w.destroy()
        section_uis.clear()

        opts = frame_spec_options()
        labels = [lbl for _, lbl in opts]
        ids = [sid for sid, _ in opts]
        spec_combo["values"] = labels

        # current selection
        cur_sid = resolve_current_spec_id(opts)
        if cur_sid not in ids:
            cur_sid = ids[0] if ids else ""
            current_spec_id_var.set(cur_sid)

        update_add_section_button_state(cur_sid)

        # set combobox display text by current sid
        display = ""
        for sid, lbl in opts:
            if sid == cur_sid:
                display = lbl
                break
        selected_spec_label_var.set(display)
        spec_combo.set(display)

        if not cur_sid:
            set_frame_schedule_error("Please add and select a valid cost code before adding a section.")
            ttk.Label(sections_wrap, text="No cost codes available. Add cost codes first.").pack(anchor="w", pady=8)
            return

        set_frame_schedule_error("")

        sections = job["frame_schedules"].setdefault(cur_sid, [])
        if not sections:
            ttk.Label(sections_wrap, text="No sections yet. Click '+ Add Section'.").pack(anchor="w", pady=8)
        else:
            for sec in sections:
                normalize_section(sec, job)
                ui = SectionUI(sections_wrap, sec, cur_sid)
                section_uis.append(ui)

    def on_spec_combo_selected(_=None):
        selected_label = spec_combo.get().strip()
        sid = ""
        for _sid, lbl in frame_spec_options():
            if lbl == selected_label:
                sid = _sid
                break
        if sid:
            set_frame_schedule_error("")
            current_spec_id_var.set(sid)
            selected_spec_label_var.set(selected_label)
            update_add_section_button_state(sid)
            refresh_frame_schedule_ui()
        else:
            current_spec_id_var.set("")
            set_frame_schedule_error("Please select a valid cost code before adding a section.")
            update_add_section_button_state("")

    spec_combo.bind("<<ComboboxSelected>>", on_spec_combo_selected)

    def add_section_ui():
        sid = resolve_current_spec_id()
        if not sid or not is_valid_frame_spec_id(job, sid):
            err = "Please select a valid cost code before adding a section."
            set_frame_schedule_error(err)
            messagebox.showerror("Invalid Cost Code", err)
            update_add_section_button_state("")
            return
        set_frame_schedule_error("")
        snapshot()
        _, err = create_frame_schedule_section(job, sid)
        if err:
            set_frame_schedule_error(err)
            messagebox.showerror("Unable to Add Section", err)
            return
        try:
            save_job(job)
            refresh_frame_schedule_ui()
            mark_frame_dirty()
        except Exception as ex:
            err_msg = f"Unable to persist new section: {ex}"
            set_frame_schedule_error(err_msg)
            messagebox.showerror("Unable to Add Section", err_msg)

    # --------------------------------------------------
    # Save-on-close / autosave hooks
    # --------------------------------------------------
    def on_tab_focus_out(_=None):
        # lightweight save; robust enough for this desktop app
        try:
            save_job(job)
            refresh_main()
        except Exception:
            pass

    tab.bind("<FocusOut>", on_tab_focus_out)

    # kick initial active tab state
    _on_tab_changed()

    return tab


class ConfigManagerDialog(tk.Toplevel):
    def __init__(self, master, on_saved=None):
        super().__init__(master)
        self.title("Global Configuration")
        self.geometry("840x540")
        self.resizable(True, True)
        self.on_saved = on_saved
        self.config_path = BASE_DIR / "global_config.json"

        self.cfg = self._load_cfg()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_mat = ttk.Frame(nb)
        self.tab_opts = ttk.Frame(nb)
        nb.add(self.tab_mat, text="Install Materials")
        nb.add(self.tab_opts, text="Dropdown Options")

        self._build_material_tab()
        self._build_options_tab()

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bar, text="Save", command=self.save).pack(side="left")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="left", padx=6)

    def _load_cfg(self):
        if self.config_path.exists():
            try:
                d = json.loads(self.config_path.read_text(encoding="utf-8"))
                if not isinstance(d, dict):
                    raise ValueError()
                # ensure defaults
                base = default_config()
                out = copy.deepcopy(base)
                out.update(d)
                # normalize materials by key
                by_key = {m.get("key"): m for m in d.get("materials", []) if isinstance(m, dict)}
                mats = []
                for tmpl in base["materials"]:
                    row = dict(tmpl)
                    old = by_key.get(tmpl["key"], {})
                    for k in ("label", "basis", "factor", "rate", "qty", "unit"):
                        if k in old:
                            row[k] = old[k]
                    mats.append(row)
                out["materials"] = mats
                return out
            except Exception:
                pass
        return default_config()

    def _build_material_tab(self):
        wrap = ttk.Frame(self.tab_mat)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        headers = ["Key", "Label", "Basis", "Factor", "Rate", "Qty (manual)", "Unit"]
        for c, h in enumerate(headers):
            ttk.Label(wrap, text=h).grid(row=0, column=c, padx=4, sticky="w")

        self.mat_rows = []
        for i, m in enumerate(self.cfg.get("materials", []), start=1):
            key_lbl = ttk.Label(wrap, text=m.get("key", ""))
            key_lbl.grid(row=i, column=0, padx=4, pady=2, sticky="w")

            label = ttk.Entry(wrap, width=28)
            label.insert(0, m.get("label", ""))
            label.grid(row=i, column=1, padx=4, pady=2, sticky="w")

            basis = ttk.Combobox(wrap, width=20, state="readonly",
                                 values=["perim_subtotal", "head_sill_subtotal", "caulk_lf_subtotal", "manual"])
            basis.set(m.get("basis", "perim_subtotal"))
            basis.grid(row=i, column=2, padx=4, pady=2, sticky="w")

            factor = ttk.Entry(wrap, width=8)
            factor.insert(0, m.get("factor", "1"))
            factor.grid(row=i, column=3, padx=4, pady=2, sticky="w")

            rate = ttk.Entry(wrap, width=8)
            rate.insert(0, m.get("rate", "0"))
            rate.grid(row=i, column=4, padx=4, pady=2, sticky="w")

            qty = ttk.Entry(wrap, width=10)
            qty.insert(0, m.get("qty", ""))
            qty.grid(row=i, column=5, padx=4, pady=2, sticky="w")

            unit = ttk.Entry(wrap, width=14)
            unit.insert(0, m.get("unit", ""))
            unit.grid(row=i, column=6, padx=4, pady=2, sticky="w")

            self.mat_rows.append({
                "key": m.get("key", ""),
                "label": label, "basis": basis, "factor": factor, "rate": rate, "qty": qty, "unit": unit
            })

    def _build_options_tab(self):
        wrap = ttk.Frame(self.tab_opts)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(wrap, text="Sheet Type Options (comma-separated)").grid(row=0, column=0, sticky="w")
        self.sheet_opts = ttk.Entry(wrap, width=90)
        self.sheet_opts.insert(0, ",".join(self.cfg.get("sheet_type_options", [])))
        self.sheet_opts.grid(row=1, column=0, sticky="w", pady=(2, 12))

        ttk.Label(wrap, text="Product Type Options (comma-separated)").grid(row=2, column=0, sticky="w")
        self.prod_opts = ttk.Entry(wrap, width=90)
        self.prod_opts.insert(0, ",".join(self.cfg.get("product_type_options", [])))
        self.prod_opts.grid(row=3, column=0, sticky="w", pady=(2, 0))

    def save(self):
        mats = []
        for r in self.mat_rows:
            mats.append({
                "key": r["key"],
                "label": r["label"].get().strip(),
                "basis": r["basis"].get().strip(),
                "factor": r["factor"].get().strip(),
                "rate": r["rate"].get().strip(),
                "qty": r["qty"].get().strip(),
                "unit": r["unit"].get().strip(),
            })

        sheet = [s.strip() for s in self.sheet_opts.get().split(",") if s.strip()]
        prod = [s.strip() for s in self.prod_opts.get().split(",") if s.strip()]

        self.cfg = {
            "materials": mats,
            "sheet_type_options": sheet or default_config()["sheet_type_options"],
            "product_type_options": prod or default_config()["product_type_options"],
        }
        self.config_path.write_text(json.dumps(self.cfg, indent=2), encoding="utf-8")
        if self.on_saved:
            self.on_saved(self.cfg)
        messagebox.showinfo("Saved", "Global configuration saved.")
        self.destroy()


class BidManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bid Manager")
        self.geometry("1280x820")
        self.minsize(1080, 700)

        self.global_cfg = self._load_global_cfg()

        # Notebook for Home + open jobs
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self.home_tab = ttk.Frame(self.nb)
        self.nb.add(self.home_tab, text="Home")

        self.open_job_tabs = {}

        self._build_home()

    def _load_global_cfg(self):
        p = BASE_DIR / "global_config.json"
        if p.exists():
            try:
                d = json.loads(p.read_text(encoding="utf-8"))
                ensure = default_config()
                ensure.update(d if isinstance(d, dict) else {})
                return ensure
            except Exception:
                pass
        return default_config()

    def _build_home(self):
        for w in self.home_tab.winfo_children():
            w.destroy()

        home_top = ttk.Frame(self.home_tab)
        home_top.pack(fill="x", padx=10, pady=10)

        ttk.Label(home_top, text="New Job Name").pack(anchor="w")
        self.new_job_entry = ttk.Entry(home_top, width=44)
        self.new_job_entry.pack(anchor="w", pady=(2, 6))

        btns = ttk.Frame(home_top)
        btns.pack(anchor="w")
        ttk.Button(btns, text="Create New Job", command=self.create_job).pack(side="left", padx=2)
        ttk.Button(btns, text="Refresh", command=self.refresh_jobs_table).pack(side="left", padx=2)
        ttk.Button(btns, text="Global Config", command=self.open_global_config).pack(side="left", padx=2)

        ttk.Separator(self.home_tab, orient="horizontal").pack(fill="x", padx=10, pady=(0, 8))

        # Summary tree
        self.summary_tree = ttk.Treeview(
            self.home_tab,
            columns=("job_name", "estimator", "pm", "status", "bid_total"),
            show="headings",
            height=8,
        )
        self.summary_tree.heading("job_name", text="Job Name")
        self.summary_tree.heading("estimator", text="Estimator")
        self.summary_tree.heading("pm", text="Project Manager")
        self.summary_tree.heading("status", text="Status")
        self.summary_tree.heading("bid_total", text="Bid Sheet Total")
        self.summary_tree.column("job_name", width=200, anchor="w")
        self.summary_tree.column("estimator", width=160, anchor="w")
        self.summary_tree.column("pm", width=180, anchor="w")
        self.summary_tree.column("status", width=140, anchor="w")
        self.summary_tree.column("bid_total", width=140, anchor="e")
        self.summary_tree.pack(fill="x", pady=(4, 10))

        self.summary_tree.bind("<Double-Button-1>", self.open_selected_job)

        # Full jobs list
        self.jobs_tree = ttk.Treeview(
            self.home_tab,
            columns=("name", "project", "due", "status", "est", "pm"),
            show="headings",
            height=14,
        )
        self.jobs_tree.heading("name", text="Job Name")
        self.jobs_tree.heading("project", text="Project Number")
        self.jobs_tree.heading("due", text="Bid Due Date")
        self.jobs_tree.heading("status", text="Status")
        self.jobs_tree.heading("est", text="Estimator")
        self.jobs_tree.heading("pm", text="Project Manager")
        self.jobs_tree.column("name", width=260, anchor="w")
        self.jobs_tree.column("project", width=130, anchor="w")
        self.jobs_tree.column("due", width=110, anchor="w")
        self.jobs_tree.column("status", width=120, anchor="w")
        self.jobs_tree.column("est", width=160, anchor="w")
        self.jobs_tree.column("pm", width=180, anchor="w")
        self.jobs_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.jobs_tree.bind("<Double-Button-1>", self.open_selected_job)

        bottom = ttk.Frame(self.home_tab)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bottom, text="Open Selected Job", command=self.open_selected_job).pack(side="left", padx=2)
        ttk.Button(bottom, text="Duplicate Selected Job", command=self.duplicate_selected_job).pack(side="left", padx=2)
        ttk.Button(bottom, text="Delete Selected Job", command=self.delete_selected_job).pack(side="left", padx=2)
        ttk.Button(bottom, text="Import Job JSON", command=self.import_job_json).pack(side="left", padx=8)
        ttk.Button(bottom, text="Export Selected Job JSON", command=self.export_selected_job_json).pack(side="left", padx=2)

        self.refresh_jobs_table()

    def open_global_config(self):
        def saved(cfg):
            self.global_cfg = cfg
            # no force migration across all jobs, but new/open job tabs can reference this if needed
        ConfigManagerDialog(self, on_saved=saved)

    def create_job(self):
        name = (self.new_job_entry.get() or "").strip()
        if not name:
            messagebox.showerror("Error", "Job Name is required.")
            return
        j = new_job_template(name)
        # seed with global cfg
        j["config"] = copy.deepcopy(self.global_cfg)
        save_job(j)
        self.new_job_entry.delete(0, tk.END)
        self.refresh_jobs_table()
        self.open_job(j)

    def _selected_job_id(self):
        sel = self.jobs_tree.selection()
        if not sel:
            return None
        return sel[0]

    def open_selected_job(self, _event=None):
        jid = self._selected_job_id()
        if not jid:
            # maybe summary tree selected
            sel = self.summary_tree.selection()
            if sel:
                jid = sel[0]
        if not jid:
            return
        p = job_path(jid)
        if not p.exists():
            messagebox.showerror("Missing", "Selected job file not found.")
            self.refresh_jobs_table()
            return
        try:
            j = json.loads(p.read_text(encoding="utf-8"))
            ensure_job_defaults(j)
        except Exception:
            messagebox.showerror("Error", "Failed to load job.")
            return
        self.open_job(j)

    def open_job(self, job):
        jid = job.get("id")
        if jid in self.open_job_tabs:
            tab = self.open_job_tabs[jid]
            self.nb.select(tab)
            return

        # merge in global cfg defaults for missing keys only
        cfg = copy.deepcopy(self.global_cfg)
        local = job.get("config", {})
        # merge by key for materials
        by_key = {m.get("key"): m for m in (local.get("materials", []) if isinstance(local.get("materials"), list) else [])}
        mats = []
        for tmpl in cfg["materials"]:
            row = dict(tmpl)
            old = by_key.get(tmpl["key"], {})
            for f in ("label", "basis", "factor", "rate", "qty", "unit"):
                if f in old:
                    row[f] = old[f]
            mats.append(row)
        cfg["materials"] = mats
        for k in ("sheet_type_options", "product_type_options"):
            if isinstance(local.get(k), list) and local.get(k):
                cfg[k] = local[k]
        job["config"] = cfg

        tab = open_job_tab(
            self, self.nb, job,
            refresh_main=self.refresh_jobs_table,
            close_tab=self.close_job_tab,
            config_manager=self.global_cfg,
        )
        self.open_job_tabs[jid] = tab

    def close_job_tab(self, tab):
        # find job id by tab object
        jid_to_remove = None
        for jid, t in self.open_job_tabs.items():
            if t == tab:
                jid_to_remove = jid
                break
        if jid_to_remove:
            self.open_job_tabs.pop(jid_to_remove, None)
        self.nb.forget(tab)

    def duplicate_selected_job(self):
        jid = self._selected_job_id()
        if not jid:
            sel = self.summary_tree.selection()
            if sel:
                jid = sel[0]
        if not jid:
            return
        p = job_path(jid)
        if not p.exists():
            return
        try:
            src = json.loads(p.read_text(encoding="utf-8"))
            ensure_job_defaults(src)
        except Exception:
            messagebox.showerror("Error", "Failed to load selected job.")
            return

        new_name = simpledialog.askstring("Duplicate Job", "New Job Name:", initialvalue=f"{src.get('job_name','Job')} Copy")
        if not new_name:
            return

        dup = copy.deepcopy(src)
        dup["id"] = str(uuid.uuid4())
        dup["job_name"] = new_name.strip()
        save_job(dup)
        self.refresh_jobs_table()
        self.open_job(dup)

    def delete_selected_job(self):
        jid = self._selected_job_id()
        if not jid:
            sel = self.summary_tree.selection()
            if sel:
                jid = sel[0]
        if not jid:
            return
        if not messagebox.askyesno("Delete Job", "Delete selected job permanently?"):
            return
        p = job_path(jid)
        try:
            if p.exists():
                p.unlink()
        except Exception:
            messagebox.showerror("Error", "Failed to delete file.")
            return

        if jid in self.open_job_tabs:
            tab = self.open_job_tabs.pop(jid)
            try:
                self.nb.forget(tab)
            except Exception:
                pass

        self.refresh_jobs_table()

    def import_job_json(self):
        fp = filedialog.askopenfilename(
            title="Import Job JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not fp:
            return
        try:
            data = json.loads(Path(fp).read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("Invalid JSON root.")
            ensure_job_defaults(data)
            if not data.get("id"):
                data["id"] = str(uuid.uuid4())
            # prevent collision
            if job_path(data["id"]).exists():
                data["id"] = str(uuid.uuid4())
            save_job(data)
            self.refresh_jobs_table()
            messagebox.showinfo("Imported", "Job imported successfully.")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def export_selected_job_json(self):
        jid = self._selected_job_id()
        if not jid:
            sel = self.summary_tree.selection()
            if sel:
                jid = sel[0]
        if not jid:
            return
        p = job_path(jid)
        if not p.exists():
            return
        out = filedialog.asksaveasfilename(
            title="Export Job JSON",
            defaultextension=".json",
            initialfile=f"{jid}.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not out:
            return
        try:
            Path(out).write_text(p.read_text(encoding="utf-8"), encoding="utf-8")
            messagebox.showinfo("Exported", f"Saved:\n{out}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def refresh_jobs_table(self):
        jobs = load_jobs()
        for j in jobs:
            ensure_job_defaults(j)
            # lazy repair to disk if malformed minimal changes
            try:
                save_job(j)
            except Exception:
                pass

        self.jobs_tree.delete(*self.jobs_tree.get_children())
        for j in sorted(jobs, key=lambda x: (x.get("job_name") or "").lower()):
            self.jobs_tree.insert(
                "", "end", iid=j["id"],
                values=(
                    j.get("job_name", ""),
                    j.get("project_number", ""),
                    j.get("bid_due_date", ""),
                    j.get("status", ""),
                    j.get("estimator", ""),
                    j.get("project_manager", ""),
                )
            )

        self.summary_tree.delete(*self.summary_tree.get_children())
        for j in sorted(jobs, key=lambda x: (x.get("job_name") or "").lower())[:50]:
            self.summary_tree.insert(
                "", "end", iid=j["id"],
                values=(
                    j.get("job_name", ""),
                    j.get("estimator", ""),
                    j.get("project_manager", ""),
                    j.get("status", ""),
                    money_fmt(compute_bid_sheet_total(j)),
                )
            )


def main():
    app = BidManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
