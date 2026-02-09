import json
import math
import uuid
import copy
from pathlib import Path
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

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
    ("Wage Data", "wage_data", "text"),
    ("Estimator", "estimator", "text"),
    ("Tax Exemption", "tax_exemption", "bool"),
    ("Project Manager", "project_manager", "text"),
    ("Contract Type", "contract_type", "text"),
    ("Engineer", "engineer", "text"),
    ("Frame / Sealant Color(s)", "frame_colors", "text"),
    ("General Contractor", "general_contractor", "text"),
    ("Field / Shop Hours Bid", "field_shop_hours", "text"),
    ("Construction Manager", "construction_manager", "text"),
    ("Frame Information", "frame_information", "text"),
    ("Product Vendors", "product_vendors", "text"),
]

def default_schedule_row():
    return {
        "spec_mark": "",
        "qty": "",
        "width": "",
        "height": "",
        "sqft": "",
        "perim": "",
        "caulk_passes": "",  # auto-fill to 3 only after data entered in another column
        "caulk_lf": "",
        "head_sill": "",
        "head": "",
        "jamb": "",
        "sill": "",
        "type": "",
        "matl": "",
        "finish": "",
        "notes": "",
    }

def default_materials_template():
    return [
        {"key": "bracing",     "label": "Bracing and Anchoring",              "basis": "perim_subtotal",     "factor": "1.00",   "rate": "1.50", "qty": "", "unit": "Linear Foot"},
        {"key": "sheet_metal", "label": "Sheet Metal Membrane Air Barriers",  "basis": "perim_subtotal",     "factor": "1.00",   "rate": "1.00", "qty": "", "unit": "Linear Foot"},
        {"key": "flashing",    "label": "Flashing and Sheet Metal",           "basis": "head_sill_subtotal", "factor": "1.00",   "rate": "8.00", "qty": "", "unit": "Linear Foot"},
        {"key": "backer_rods", "label": "Backer Rods",                        "basis": "caulk_lf_subtotal",  "factor": "1.00",   "rate": "0.50", "qty": "", "unit": "Linear Foot"},
        {"key": "sealants",    "label": "Joint Sealants",                     "basis": "caulk_lf_subtotal",  "factor": "0.0833", "rate": "12.00","qty": "", "unit": "Sausage"},
        {"key": "tie_back",    "label": "Tie Back",                           "basis": "manual",             "factor": "",       "rate": "45.00","qty": "", "unit": "Each"},
        {"key": "backpans",    "label": "Backpans Insulation",                "basis": "manual",             "factor": "",       "rate": "48.32","qty": "", "unit": "Linear Foot"},
    ]

def default_materials():
    return copy.deepcopy(default_materials_template())

def default_schedule_section(spec_id: str):
    return {
        "id": str(uuid.uuid4()),
        "spec_id": spec_id,
        "rows": [default_schedule_row()],
        "materials": default_materials(),
        "install_material_total": 0,
    }

def ensure_job_defaults(job: dict) -> None:
    job.setdefault("id", str(uuid.uuid4()))
    for _, key, t in JOB_FIELDS:
        if key not in job:
            job[key] = False if t == "bool" else ""
    job.setdefault("cost_codes", [])
    job.setdefault("quotes", {})
    job.setdefault("frame_schedules", {})
    job.setdefault("frame_schedule_rollups", {})

    for cc in job["cost_codes"]:
        cc.setdefault("id", str(uuid.uuid4()))
        cc.setdefault("code", "00 00 00")
        cc.setdefault("description", "")
        cc.setdefault("alts", [])

    normalize_quotes(job)
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
# Frame schedule model normalization + deterministic totals
# ======================================================

FS_HEADERS = [
    ("SPEC / MARK", "spec_mark", 16, "text"),
    ("QTY", "qty", 6, "num"),
    ("WIDTH", "width", 8, "num"),
    ("HEIGHT", "height", 8, "num"),
    ("SQFT", "sqft", 8, "calc"),
    ("PERIM", "perim", 8, "calc"),
    ("CAULK PASSES", "caulk_passes", 10, "num"),
    ("CAULK LF", "caulk_lf", 10, "calc"),
    ("HEAD/SILL", "head_sill", 10, "calc"),
    ("HEAD", "head", 10, "text"),
    ("JAMB", "jamb", 10, "text"),
    ("SILL", "sill", 10, "text"),
    ("TYPE", "type", 10, "text"),
    ("MAT'L", "matl", 10, "text"),
    ("FINISH", "finish", 10, "text"),
    ("NOTES", "notes", 24, "text"),
]

def qty_populated(row: dict) -> bool:
    return str(row.get("qty", "")).strip() != ""

def calc_schedule_row(row: dict):
    qty = safe_float(row.get("qty", ""))
    wid = safe_float(row.get("width", ""))
    hei = safe_float(row.get("height", ""))

    if qty <= 0:
        row["sqft"] = ""
        row["perim"] = ""
        row["caulk_lf"] = ""
        row["head_sill"] = ""
        return

    row["sqft"] = str(roundup((wid * hei * qty) / 144.0))
    row["perim"] = str(roundup(2.0 * ((wid / 12.0) + (hei / 12.0)) * qty))

    cp = safe_float(row.get("caulk_passes", ""))
    if cp <= 0:
        cp = 0.0
    row["caulk_lf"] = str(roundup(safe_int(row["perim"]) * cp)) if cp > 0 else ""
    row["head_sill"] = str(roundup(qty * wid / 6.0))

def schedule_subtotals(rows):
    return {
        "qty": sum(safe_int(r.get("qty", 0)) for r in rows if qty_populated(r)),
        "sqft": sum(safe_int(r.get("sqft", 0)) for r in rows if qty_populated(r)),
        "perim": sum(safe_int(r.get("perim", 0)) for r in rows if qty_populated(r)),
        "caulk_lf": sum(safe_int(r.get("caulk_lf", 0)) for r in rows if qty_populated(r)),
        "head_sill": sum(safe_int(r.get("head_sill", 0)) for r in rows if qty_populated(r)),
    }

def build_valid_frame_spec_ids(job: dict):
    out = set()
    for cc in job.get("cost_codes", []):
        code = (cc.get("code") or "").strip()
        if not code:
            continue
        for v in variants_for_cc(cc.get("alts", [])):
            out.add(frame_spec_id(code, v))
    return out

def material_qty_for_row(material: dict, subtotals: dict):
    basis = material.get("basis", "manual")
    factor = safe_float(material.get("factor", ""))
    if basis == "perim_subtotal":
        return subtotals["perim"] * factor
    if basis == "head_sill_subtotal":
        return subtotals["head_sill"] * factor
    if basis == "caulk_lf_subtotal":
        return subtotals["caulk_lf"] * factor
    return safe_float(material.get("qty", ""))  # manual

def recompute_section_totals(section: dict) -> int:
    rows = section.get("rows", []) or []
    for r in rows:
        if isinstance(r, dict):
            calc_schedule_row(r)

    subs = schedule_subtotals(rows)
    mats = section.get("materials", []) or []
    total_cost = 0

    basis_subs = {
        "perim": safe_int(subs["perim"]),
        "caulk_lf": safe_int(subs["caulk_lf"]),
        "head_sill": safe_int(subs["head_sill"]),
    }

    for m in mats:
        if not isinstance(m, dict):
            continue
        if m.get("key") == "sealants" and str(m.get("factor", "")).strip() == "":
            m["factor"] = "0.0833"
        qty_val = material_qty_for_row(m, basis_subs)
        rate_val = safe_float(m.get("rate", ""))
        cost_int = roundup(qty_val * rate_val) if qty_val * rate_val != 0 else 0
        total_cost += cost_int

    section["install_material_total"] = total_cost
    return total_cost

def normalize_frame_schedules(job: dict) -> None:
    valid_specs = build_valid_frame_spec_ids(job)
    job.setdefault("frame_schedules", {})
    job.setdefault("frame_schedule_rollups", {})

    for key in list(job["frame_schedules"].keys()):
        if "||" not in key:
            base = (key or "").strip()
            if not base:
                del job["frame_schedules"][key]
                continue
            new_key = frame_spec_id(base, "BASE")
            if new_key in job["frame_schedules"]:
                v = job["frame_schedules"][key]
                if isinstance(v, list):
                    job["frame_schedules"][new_key].extend(v)
                elif isinstance(v, dict):
                    job["frame_schedules"][new_key].append(v)
                del job["frame_schedules"][key]
            else:
                job["frame_schedules"][new_key] = job["frame_schedules"].pop(key)

    for spec_id in list(job["frame_schedules"].keys()):
        if spec_id not in valid_specs:
            del job["frame_schedules"][spec_id]

    for spec_id in valid_specs:
        job["frame_schedules"].setdefault(spec_id, [])

    for spec_id, sections in job["frame_schedules"].items():
        if sections is None:
            job["frame_schedules"][spec_id] = []
            sections = job["frame_schedules"][spec_id]
        if isinstance(sections, dict):
            s = default_schedule_section(spec_id)
            s["rows"] = sections.get("rows", []) or [default_schedule_row()]
            s["materials"] = sections.get("materials", default_materials()) or default_materials()
            sections = [s]
            job["frame_schedules"][spec_id] = sections
        if not isinstance(sections, list):
            job["frame_schedules"][spec_id] = []
            sections = job["frame_schedules"][spec_id]

        for i in range(len(sections)):
            s = sections[i]
            if not isinstance(s, dict):
                s = default_schedule_section(spec_id)
                sections[i] = s
            s.setdefault("id", str(uuid.uuid4()))
            s["spec_id"] = spec_id
            s.setdefault("rows", [default_schedule_row()])
            s.setdefault("materials", default_materials())
            s.setdefault("install_material_total", 0)

            if not isinstance(s["rows"], list) or len(s["rows"]) == 0:
                s["rows"] = [default_schedule_row()]
            for r_i in range(len(s["rows"])):
                base = default_schedule_row()
                r = s["rows"][r_i]
                if isinstance(r, dict):
                    base.update(r)
                s["rows"][r_i] = base

            if not isinstance(s["materials"], list) or len(s["materials"]) == 0:
                s["materials"] = default_materials()

            for m in s["materials"]:
                if m.get("key") == "sealants" and str(m.get("factor", "")).strip() == "":
                    m["factor"] = "0.0833"

            recompute_section_totals(s)

    compute_frame_schedule_rollups(job)

def compute_frame_schedule_rollups(job: dict) -> None:
    job.setdefault("frame_schedule_rollups", {})
    for spec_id, sections in job.get("frame_schedules", {}).items():
        total = 0
        if isinstance(sections, list):
            for s in sections:
                if isinstance(s, dict):
                    total += safe_int(recompute_section_totals(s))
        job["frame_schedule_rollups"][spec_id] = {"install_material_total": total}


# ======================================================
# Job window
# ======================================================

def open_job_window(root, job: dict, refresh_main):
    ensure_job_defaults(job)

    win = tk.Toplevel(root)
    win.title(job.get("job_name") or "Job")
    win.geometry("1600x900")

    # ---------------- UNDO ----------------
    undo_stack = []
    UNDO_LIMIT = 50

    def snapshot_job():
        return json.loads(json.dumps(job))

    def push_undo():
        undo_stack.append(snapshot_job())
        if len(undo_stack) > UNDO_LIMIT:
            del undo_stack[0]

    def apply_snapshot(snap: dict):
        job.clear()
        job.update(snap)
        ensure_job_defaults(job)

    def undo(_evt=None):
        if not undo_stack:
            return
        snap = undo_stack.pop()
        apply_snapshot(snap)
        refresh_job_info_widgets_from_job()
        refresh_cc_tree()
        refresh_quotes_ui(full=True)
        fs_refresh_spec_options()
        fs_manager.reconcile_all()

    win.bind_all("<Control-z>", undo)

    topbar = ttk.Frame(win)
    topbar.pack(fill="x", padx=8, pady=6)
    ttk.Button(topbar, text="Undo (Ctrl+Z)", command=undo).pack(side="left")

    nb = ttk.Notebook(win)
    nb.pack(fill="both", expand=True)

    # --------------------------------------------------
    # Tab: Job Info
    # --------------------------------------------------
    tab_info = ttk.Frame(nb)
    nb.add(tab_info, text="Job Info")

    info_scroll = ScrollFrame(tab_info)
    info_scroll.pack(fill="both", expand=True, padx=8, pady=8)

    info_widgets = {}
    for i, (label, key, ftype) in enumerate(JOB_FIELDS):
        row = i // 2
        col = (i % 2) * 2
        ttk.Label(info_scroll.inner, text=label).grid(row=row, column=col, sticky="w", padx=10, pady=6)

        if ftype == "bool":
            var = tk.BooleanVar(value=bool(job.get(key, False)))
            cb = ttk.Checkbutton(info_scroll.inner, variable=var)
            cb.grid(row=row, column=col + 1, sticky="w", padx=10, pady=6)
            info_widgets[key] = ("bool", var)
            cb.bind("<ButtonRelease-1>", lambda e: win.after(0, push_undo))
        elif ftype == "date":
            w = make_date_widget(info_scroll.inner, job.get(key, ""))
            w.grid(row=row, column=col + 1, sticky="w", padx=10, pady=6)
            info_widgets[key] = ("date", w)
        else:
            e = ttk.Entry(info_scroll.inner, width=52)
            e.insert(0, job.get(key, "") or "")
            e.grid(row=row, column=col + 1, sticky="w", padx=10, pady=6)
            info_widgets[key] = ("text", e)

    # Export Job Info PDF button
    def export_job_info_pdf():
        """Export the current job information to a PDF file."""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("PDF Export", "The reportlab library is not installed.\nPlease install reportlab to enable PDF export.")
            return
        # Pull current values into the job dict
        collect_job_info_into_job()
        ensure_job_defaults(job)
        default_name = f"{job.get('job_name', 'job')}_info.pdf"
        file_path = filedialog.asksaveasfilename(
            title="Save Job Info PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            job_name = job.get("job_name", "").strip() or "Job"
            elements.append(Paragraph(f"Job Information – {job_name}", styles['Title']))
            elements.append(Spacer(1, 12))
            # Build table data
            table_data = [["Field", "Value"]]
            for label, key, ftype in JOB_FIELDS:
                val = job.get(key, "")
                if ftype == "bool":
                    display = "Yes" if val else "No"
                else:
                    display = str(val or "")
                table_data.append([label, display])
            col_widths = [200, 340]
            tbl = Table(table_data, colWidths=col_widths)
            tbl_style = TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 12),
                ('BOTTOMPADDING', (0,0), (-1,0), 8),
                ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
                ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ])
            tbl.setStyle(tbl_style)
            elements.append(tbl)
            doc.build(elements)
            messagebox.showinfo("PDF Export", f"Job info successfully exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("PDF Export", f"Failed to export PDF:\n{e}")

    # Position the export button on its own row
    total_rows = (len(JOB_FIELDS) + 1) // 2
    btn_info_export = ttk.Button(info_scroll.inner, text="Export PDF", command=export_job_info_pdf)
    btn_info_export.grid(row=total_rows, column=0, columnspan=4, pady=(10, 10), sticky="e")

    def refresh_job_info_widgets_from_job():
        for _label, key, _ftype in JOB_FIELDS:
            t, w = info_widgets[key]
            if t == "bool":
                w.set(bool(job.get(key, False)))
            elif t == "date":
                set_date_value(w, job.get(key, "") or "")
            else:
                w.delete(0, tk.END)
                w.insert(0, job.get(key, "") or "")

    # --------------------------------------------------
    # Tab: Job Cost Codes
    # --------------------------------------------------
    tab_cc = ttk.Frame(nb)
    nb.add(tab_cc, text="Job Cost Codes")

    cc_frame = ttk.Frame(tab_cc, padding=10)
    cc_frame.pack(fill="both", expand=True)

    cc_tree = ttk.Treeview(cc_frame, columns=("code", "desc", "alts"), show="headings", height=14)
    cc_tree.heading("code", text="CODE (00 00 00)")
    cc_tree.heading("desc", text="DESCRIPTION")
    cc_tree.heading("alts", text="ALTs (1..25 comma-separated)")
    cc_tree.column("code", width=160, anchor="w")
    cc_tree.column("desc", width=700, anchor="w")
    cc_tree.column("alts", width=260, anchor="w")
    cc_tree.pack(fill="both", expand=True)

    def find_cc_by_id(cc_id: str):
        for cc in job["cost_codes"]:
            if cc.get("id") == cc_id:
                return cc
        return None

    def migrate_code_key(old_code: str, new_code: str):
        old_code = (old_code or "").strip()
        new_code = (new_code or "").strip()
        if not old_code or old_code == new_code:
            return

        if old_code in job["quotes"]:
            if new_code in job["quotes"]:
                for v, rows in job["quotes"][old_code].items():
                    job["quotes"][new_code].setdefault(v, rows)
                del job["quotes"][old_code]
            else:
                job["quotes"][new_code] = job["quotes"].pop(old_code)

        for spec_id in list(job.get("frame_schedules", {}).keys()):
            base, var = parse_frame_spec_id(spec_id)
            if base != old_code:
                continue
            new_spec = frame_spec_id(new_code, var)
            if new_spec in job["frame_schedules"]:
                job["frame_schedules"][new_spec].extend(job["frame_schedules"][spec_id])
                del job["frame_schedules"][spec_id]
            else:
                job["frame_schedules"][new_spec] = job["frame_schedules"].pop(spec_id)

    def refresh_cc_tree():
        cc_tree.delete(*cc_tree.get_children())
        job["cost_codes"].sort(key=lambda x: (x.get("code") or ""))
        for cc in job["cost_codes"]:
            cc_tree.insert(
                "", "end", iid=cc["id"],
                values=(cc.get("code", ""), cc.get("description", ""), ",".join(str(a) for a in cc.get("alts", [])))
            )

    def begin_inline_edit(item_iid: str, col: str):
        if not item_iid or col not in ("#1", "#2", "#3"):
            return
        bbox = cc_tree.bbox(item_iid, col)
        if not bbox:
            return
        x, y, w, h = bbox
        entry = ttk.Entry(cc_tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, cc_tree.set(item_iid, col))
        entry.focus_set()

        push_undo()

        def commit(_=None):
            try:
                cc = find_cc_by_id(item_iid)
                if not cc:
                    return
                val = entry.get().strip()
                if col == "#1":
                    old_code = cc.get("code", "")
                    cc["code"] = val
                    migrate_code_key(old_code, val)
                elif col == "#2":
                    cc["description"] = val
                else:
                    cc["alts"] = parse_alts(val)

                normalize_quotes(job)
                normalize_frame_schedules(job)

                refresh_cc_tree()
                refresh_quotes_ui(full=True)
                fs_refresh_spec_options()
                fs_manager.reconcile_all()
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)

    def cc_on_double_click(event):
        item = cc_tree.identify_row(event.y)
        col = cc_tree.identify_column(event.x)
        if item and col in ("#1", "#2", "#3"):
            begin_inline_edit(item, col)

    cc_tree.bind("<Double-Button-1>", cc_on_double_click)

    cc_btns = ttk.Frame(cc_frame)
    cc_btns.pack(fill="x", pady=8)

    def add_cost_code():
        push_undo()
        cc = {"id": str(uuid.uuid4()), "code": "00 00 00", "description": "", "alts": []}
        job["cost_codes"].append(cc)
        normalize_quotes(job)
        normalize_frame_schedules(job)
        refresh_cc_tree()
        refresh_quotes_ui(full=True)
        fs_refresh_spec_options()
        fs_manager.reconcile_all()
        cc_tree.selection_set(cc["id"])
        cc_tree.see(cc["id"])
        win.after(50, lambda: begin_inline_edit(cc["id"], "#1"))

    def delete_selected_cost_code():
        sel = cc_tree.selection()
        if not sel:
            return
        cc_id = sel[0]
        cc = find_cc_by_id(cc_id)
        if not cc:
            return
        code = (cc.get("code") or "").strip()
        if not messagebox.askyesno("Delete", f"Delete cost code {code}? This removes schedules for BASE + ALTs."):
            return

        push_undo()
        job["cost_codes"] = [c for c in job["cost_codes"] if c.get("id") != cc_id]
        if code in job["quotes"]:
            del job["quotes"][code]
        for spec_id in list(job.get("frame_schedules", {}).keys()):
            base, _v = parse_frame_spec_id(spec_id)
            if base == code:
                del job["frame_schedules"][spec_id]

        normalize_quotes(job)
        normalize_frame_schedules(job)
        refresh_cc_tree()
        refresh_quotes_ui(full=True)
        fs_refresh_spec_options()
        fs_manager.reconcile_all()

    ttk.Button(cc_btns, text="Add Cost Code", command=add_cost_code).pack(side="left")
    ttk.Button(cc_btns, text="Delete Selected", command=delete_selected_cost_code).pack(side="left", padx=8)

    # Export cost codes to PDF
    def export_cost_codes_pdf():
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("PDF Export", "The reportlab library is not installed.\nPlease install reportlab to enable PDF export.")
            return
        collect_job_info_into_job()
        ensure_job_defaults(job)
        default_name = f"{job.get('job_name', 'job')}_cost_codes.pdf"
        file_path = filedialog.asksaveasfilename(
            title="Save Cost Codes PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            job_name = job.get("job_name", "").strip() or "Job"
            elements.append(Paragraph(f"Cost Codes – {job_name}", styles['Title']))
            elements.append(Spacer(1, 12))
            table_data = [["Code", "Description", "ALTs"]]
            for cc in sorted(job.get("cost_codes", []), key=lambda x: (x.get("code") or "")):
                code = (cc.get("code", "")).strip()
                desc = cc.get("description", "")
                alts = ", ".join(str(a) for a in (cc.get("alts", []) or []))
                table_data.append([code, desc, alts])
            col_widths = [100, 280, 160]
            tbl = Table(table_data, colWidths=col_widths)
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 12),
                ('BOTTOMPADDING', (0,0), (-1,0), 8),
                ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
                ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ]))
            elements.append(tbl)
            doc.build(elements)
            messagebox.showinfo("PDF Export", f"Cost codes successfully exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("PDF Export", f"Failed to export PDF:\n{e}")

    ttk.Button(cc_btns, text="Export PDF", command=export_cost_codes_pdf).pack(side="right")

    # --------------------------------------------------
    # Tab: Quotes
    # --------------------------------------------------
    tab_q = ttk.Frame(nb)
    nb.add(tab_q, text="Quotes")

    # Top bar for quotes tab
    quotes_topbar = ttk.Frame(tab_q)
    quotes_topbar.pack(fill="x")

    # Export quotes to PDF
    def export_quotes_pdf():
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("PDF Export", "The reportlab library is not installed.\nPlease install reportlab to enable PDF export.")
            return
        collect_job_info_into_job()
        ensure_job_defaults(job)
        normalize_quotes(job)
        default_name = f"{job.get('job_name', 'job')}_quotes.pdf"
        file_path = filedialog.asksaveasfilename(
            title="Save Quotes PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            job_name = job.get("job_name", "").strip() or "Job"
            elements.append(Paragraph(f"Quotes – {job_name}", styles['Title']))
            elements.append(Spacer(1, 12))
            for cc in sorted(job.get("cost_codes", []), key=lambda x: (x.get("code") or "")):
                code = (cc.get("code", "")).strip()
                if not code:
                    continue
                variants = variants_for_cc(cc.get("alts", []))
                for variant in variants:
                    quotes_list = job["quotes"].get(code, {}).get(variant, [])
                    if not quotes_list:
                        continue
                    label = frame_spec_label(code, variant) if variant != "BASE" else code
                    elements.append(Paragraph(f"{label}", styles['Heading2']))
                    elements.append(Spacer(1, 6))
                    table_data = [["Date", "Vendor(s)", "Price", "Surcharge %", "Cost", "Notes"]]
                    for q in quotes_list:
                        q_date = q.get("date", "")
                        vendor = q.get("vendor", "")
                        price = safe_int(q.get("price", 0))
                        surcharge = safe_float(q.get("surcharge", 0))
                        cost = calc_cost(price, surcharge)
                        notes = q.get("notes", "")
                        table_data.append([
                            q_date,
                            vendor,
                            money_fmt(price),
                            pct_fmt(surcharge),
                            money_fmt(cost),
                            notes,
                        ])
                    col_widths = [70, 130, 70, 70, 70, 160]
                    tbl = Table(table_data, colWidths=col_widths)
                    tbl.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 11),
                        ('BOTTOMPADDING', (0,0), (-1,0), 6),
                        ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
                        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                    ]))
                    elements.append(tbl)
                    elements.append(Spacer(1, 12))
            doc.build(elements)
            messagebox.showinfo("PDF Export", f"Quotes successfully exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("PDF Export", f"Failed to export PDF:\n{e}")

    ttk.Button(quotes_topbar, text="Export PDF", command=export_quotes_pdf).pack(side="right", padx=6, pady=6)

    quotes_scroll = ScrollFrame(tab_q)
    quotes_scroll.pack(fill="both", expand=True, padx=10, pady=10)

    collapsed_specs = set()
    spec_frames = {}

    def summary_total_avg(quotes_list):
        costs = [safe_int(q.get("cost", 0) or 0) for q in quotes_list if safe_int(q.get("cost", 0) or 0) > 0]
        total = sum(costs)
        avg = int(total / len(costs)) if costs else 0
        return total, avg

    def toggle_spec(spec_key):
        if spec_key in collapsed_specs:
            collapsed_specs.remove(spec_key)
        else:
            collapsed_specs.add(spec_key)
        rebuild_spec_block(spec_key)

    def ensure_quote_defaults(q: dict):
        q.setdefault("date", "")
        q.setdefault("vendor", "")
        q.setdefault("price", 0)
        q.setdefault("surcharge", 0.0)
        q.setdefault("cost", 0)
        q.setdefault("notes", "")

    def add_quote(code: str, variant: str):
        push_undo()
        job["quotes"][code][variant].append({})
        rebuild_spec_block((code, variant))

    def delete_quote(code: str, variant: str, idx: int):
        push_undo()
        try:
            job["quotes"][code][variant].pop(idx)
            if len(job["quotes"][code][variant]) == 0:
                job["quotes"][code][variant].append({})
        except Exception:
            pass
        rebuild_spec_block((code, variant))

    def build_spec_block(parent, code: str, desc: str, variant: str):
        spec_key = (code, variant)
        is_collapsed = spec_key in collapsed_specs
        quotes_list = job["quotes"][code][variant]

        header = ttk.Frame(parent)
        header.pack(fill="x", pady=(6, 0))

        ttk.Button(header, text="▶" if is_collapsed else "▼", width=2,
                   command=lambda k=spec_key: toggle_spec(k)).pack(side="left")

        title = f"{variant} — {code} — {desc}" if variant != "BASE" else f"{code} — {desc}"
        ttk.Label(header, text=title, font=("TkDefaultFont", 10, "bold")).pack(side="left", padx=(4, 8))
        ttk.Button(header, text="+ Quote", command=lambda c=code, v=variant: add_quote(c, v)).pack(side="left")

        if is_collapsed:
            total, avg = summary_total_avg(quotes_list)
            ttk.Label(parent, text=f"Total {money_fmt(total)} · Avg {money_fmt(avg)}", foreground="gray")\
                .pack(anchor="w", padx=28, pady=(2, 6))
            return

        grid = ttk.Frame(parent)
        grid.pack(fill="x", padx=22, pady=(2, 6))

        headers = ["DATE", "VENDOR(S)", "PRICE", "SUR %", "COST", "NOTES", ""]
        widths = [12, 18, 12, 10, 12, 30, 2]
        for c, h in enumerate(headers):
            ttk.Label(grid, text=h).grid(row=0, column=c, padx=4, sticky="w")

        for idx, q in enumerate(quotes_list):
            ensure_quote_defaults(q)

            date_w = make_date_widget(grid, q["date"])
            date_w.grid(row=idx + 1, column=0, padx=4, sticky="w")

            vendor_e = ttk.Entry(grid, width=widths[1])
            vendor_e.insert(0, q["vendor"])
            vendor_e.grid(row=idx + 1, column=1, padx=4, sticky="w")

            price_e = ttk.Entry(grid, width=widths[2])
            price_e.insert(0, money_fmt(q["price"]))
            price_e.grid(row=idx + 1, column=2, padx=4, sticky="w")

            sur_e = ttk.Entry(grid, width=widths[3])
            sur_e.insert(0, pct_fmt(q["surcharge"]))
            sur_e.grid(row=idx + 1, column=3, padx=4, sticky="w")

            cost_lbl = ttk.Label(grid, text=money_fmt(q["cost"]))
            cost_lbl.grid(row=idx + 1, column=4, padx=4, sticky="w")

            notes_e = ttk.Entry(grid, width=widths[5])
            notes_e.insert(0, q["notes"])
            notes_e.grid(row=idx + 1, column=5, padx=4, sticky="w")

            ttk.Button(grid, text="−", width=2,
                       command=lambda i=idx, c=code, v=variant: delete_quote(c, v, i))\
                .grid(row=idx + 1, column=6, padx=4, sticky="w")

            def recalc(qref=q, pe=price_e, se=sur_e, cl=cost_lbl, dw=date_w, ve=vendor_e, ne=notes_e):
                price = parse_money(pe.get())
                sur = parse_pct(se.get())
                current_date = get_date_value(dw)
                if price > 0 and not current_date.strip():
                    set_date_value(dw, today_str())
                    current_date = get_date_value(dw)

                qref["date"] = current_date
                qref["vendor"] = ve.get().strip()
                qref["price"] = price
                qref["surcharge"] = sur
                qref["cost"] = calc_cost(price, sur)
                qref["notes"] = ne.get().strip()
                cl.config(text=money_fmt(qref["cost"]))

            def mark_undo_once(_e=None):
                if not hasattr(price_e, "_undo_marked"):
                    price_e._undo_marked = True
                    push_undo()

            for w in (price_e, sur_e, vendor_e, notes_e):
                w.bind("<FocusIn>", mark_undo_once)
                w.bind("<KeyRelease>", lambda e, f=recalc: f())

            if HAS_TKCALENDAR and isinstance(date_w, DateEntry):
                date_w.bind("<<DateEntrySelected>>", lambda e, f=recalc: f())
            else:
                date_w.bind("<KeyRelease>", lambda e, f=recalc: f())

            def fmt_price_on_blur(_e=None, w=price_e, f=recalc):
                v = parse_money(w.get())
                w.delete(0, tk.END)
                w.insert(0, money_fmt(v))
                f()
                if hasattr(w, "_undo_marked"):
                    delattr(w, "_undo_marked")

            def fmt_sur_on_blur(_e=None, w=sur_e, f=recalc):
                v = parse_pct(w.get())
                w.delete(0, tk.END)
                w.insert(0, pct_fmt(v))
                f()

            price_e.bind("<FocusOut>", fmt_price_on_blur)
            sur_e.bind("<FocusOut>", fmt_sur_on_blur)

            recalc()

        total, avg = summary_total_avg(quotes_list)
        ttk.Label(parent, text=f"Total {money_fmt(total)} · Avg {money_fmt(avg)}", foreground="gray")\
            .pack(anchor="w", padx=28, pady=(0, 2))

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

    def refresh_quotes_ui(full=True):
        normalize_quotes(job)
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

    fs_spec_var = tk.StringVar()
    fs_spec_combo = ttk.Combobox(selector_frame, textvariable=fs_spec_var, state="readonly", width=28)
    fs_spec_combo.pack(side="left", padx=6)

    fs_label_to_spec = {}
    fs_spec_to_label = {}

    def fs_refresh_spec_options():
        fs_label_to_spec.clear()
        fs_spec_to_label.clear()
        labels = []
        for cc in sorted(job.get("cost_codes", []), key=lambda x: (x.get("code") or "")):
            code = (cc.get("code") or "").strip()
            if not code:
                continue
            for v in variants_for_cc(cc.get("alts", [])):
                sid = frame_spec_id(code, v)
                lab = frame_spec_label(code, v)
                labels.append(lab)
                fs_label_to_spec[lab] = sid
                fs_spec_to_label[sid] = lab

        fs_spec_combo["values"] = labels
        if labels:
            if fs_spec_var.get() not in labels:
                fs_spec_combo.current(0)
                fs_spec_var.set(labels[0])
        else:
            fs_spec_var.set("")

    class FrameScheduleSectionUI:
        def __init__(self, parent, spec_id: str, section: dict, section_index: int, on_delete_section, on_any_change, push_undo_fn):
            self.parent = parent
            self.spec_id = spec_id
            self.section = section
            self.section_index = section_index
            self.on_delete_section = on_delete_section
            self.on_any_change = on_any_change
            self.push_undo = push_undo_fn

            self._alive = True
            self.install_total_var = tk.StringVar(value=money_fmt(section.get("install_material_total", 0)))

            self.labelframe = ttk.LabelFrame(parent, text=self._title_text())
            self.labelframe.pack(fill="x", padx=8, pady=8)

            topbar = ttk.Frame(self.labelframe)
            topbar.pack(fill="x", padx=6, pady=(6, 2))
            ttk.Button(topbar, text="Remove Schedule", command=self._remove_schedule).pack(side="left")

            self.grid = ttk.Frame(self.labelframe)
            self.grid.pack(fill="x", padx=6, pady=(6, 2))

            self.col_index = {}
            for c, (h, key, _w, _kind) in enumerate(FS_HEADERS):
                ttk.Label(self.grid, text=h, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=c, padx=3, sticky="w")
                self.col_index[key] = c
            ttk.Label(self.grid, text="", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=len(FS_HEADERS), padx=3, sticky="w")

            self.row_widgets = []
            self.row_del_buttons = []

            self.sub_vars = {
                "qty": tk.StringVar(value="0"),
                "sqft": tk.StringVar(value="0"),
                "perim": tk.StringVar(value="0"),
                "caulk_lf": tk.StringVar(value="0"),
                "head_sill": tk.StringVar(value="0"),
            }
            self.subtotal_row_index = None

            self.mat_frame = ttk.LabelFrame(self.labelframe, text="INSTALL MATERIAL COSTS")
            self.mat_grid = None
            self.mat_qty_vars = {}
            self.mat_cost_vars = {}
            self.mat_qty_entries = {}
            self.mat_factor_entries = {}
            self.mat_rate_entries = {}
            self.total_cost_var = tk.StringVar(value=money_fmt(0))

            # render
            self._render_rows_initial_fixed()
            self._render_subtotal_row()
            self._render_materials_block()
            self.recalc_all(allow_row_change=True)

        def destroy(self):
            self._alive = False
            try:
                self.labelframe.destroy()
            except Exception:
                pass

        def _title_text(self):
            return f"Schedule {self.section_index} (Install Total: {self.install_total_var.get()})"

        def set_index(self, idx: int):
            self.section_index = idx
            self.labelframe.configure(text=self._title_text())

        def _remove_schedule(self):
            if not messagebox.askyesno("Remove", "Remove this schedule section?"):
                return
            self.push_undo()
            self.on_delete_section(self.section["id"])

        def _render_rows_initial_fixed(self):
            rows = self.section.get("rows", [])
            if not isinstance(rows, list) or len(rows) == 0:
                rows = [default_schedule_row()]
                self.section["rows"] = rows

            populated = sum(1 for r in rows if qty_populated(r))
            target = max(1, populated + 1)

            while len(rows) < target:
                rows.append(default_schedule_row())
            while len(rows) > target and not qty_populated(rows[-1]):
                rows.pop()

            for idx, row in enumerate(rows):
                self._add_row_widgets(idx, row)

        def _add_row_widgets(self, idx, row_dict):
            ui_row = {}
            grid_row = idx + 1

            for c, (_h, key, width, kind) in enumerate(FS_HEADERS):
                e = ttk.Entry(self.grid, width=width)
                e.insert(0, row_dict.get(key, ""))
                e.grid(row=grid_row, column=c, padx=3, sticky="w")
                ui_row[key] = e

                if kind == "calc":
                    e.config(state="readonly")
                else:
                    e.bind("<KeyRelease>", lambda ev, r=idx, k=key: self._on_edit(r, k))
                    e.bind("<FocusIn>", lambda ev: self._mark_undo_once())
                    e.bind("<FocusOut>", lambda ev, r=idx: self._commit_row_from_widgets(r))

            del_btn = ttk.Button(self.grid, text="−", width=2, command=lambda r=idx: self._delete_row(r))
            del_btn.grid(row=grid_row, column=len(FS_HEADERS), padx=3, sticky="w")

            self.row_widgets.append(ui_row)
            self.row_del_buttons.append(del_btn)

            for col_i, (_h, key, _w, _kind) in enumerate(FS_HEADERS):
                w = ui_row[key]
                w.bind("<Return>", lambda ev, r=idx, c=col_i: self._nav_down(r, c))
                w.bind("<Down>", lambda ev, r=idx, c=col_i: self._nav_down(r, c))

        def _nav_down(self, row_i, col_i):
            next_row = row_i + 1
            if next_row >= len(self.row_widgets):
                return "break"
            if col_i >= len(FS_HEADERS):
                return "break"
            key = FS_HEADERS[col_i][1]
            tgt = self.row_widgets[next_row].get(key)
            if tgt:
                tgt.focus_set()
                tgt.select_range(0, tk.END)
            return "break"

        def _mark_undo_once(self):
            if not hasattr(self, "_undo_marked"):
                self._undo_marked = True
                self.push_undo()

        def _clear_undo_mark(self):
            if hasattr(self, "_undo_marked"):
                delattr(self, "_undo_marked")

        def _commit_row_from_widgets(self, row_idx):
            if not self._alive:
                return
            if row_idx < 0 or row_idx >= len(self.section["rows"]) or row_idx >= len(self.row_widgets):
                return
            row = self.section["rows"][row_idx]
            for (_h, key, _w, kind) in FS_HEADERS:
                if kind != "calc":
                    row[key] = (self.row_widgets[row_idx][key].get() or "").strip()
            self._clear_undo_mark()

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

            ttk.Label(self.grid, text="SUBTOTAL", font=("TkDefaultFont", 9, "bold")).grid(row=r, column=0, padx=3, sticky="w")

            def ro_entry(col_key: str, var: tk.StringVar, width=10):
                e = ttk.Entry(self.grid, width=width, textvariable=var)
                e.grid(row=r, column=self.col_index[col_key], padx=3, sticky="w")
                e.config(state="readonly")
                return e

            ro_entry("qty", self.sub_vars["qty"], width=6)
            ro_entry("sqft", self.sub_vars["sqft"], width=8)
            ro_entry("perim", self.sub_vars["perim"], width=8)
            ro_entry("caulk_lf", self.sub_vars["caulk_lf"], width=10)
            ro_entry("head_sill", self.sub_vars["head_sill"], width=10)

        def _enforce_blank_row(self):
            rows = self.section["rows"]
            populated = sum(1 for r in rows if qty_populated(r))
            target = max(1, populated + 1)

            self._remove_subtotal_row_widgets()

            while len(rows) < target:
                rows.append(default_schedule_row())
                self._add_row_widgets(len(rows) - 1, rows[-1])

            while len(rows) > target:
                if qty_populated(rows[-1]):
                    break
                last_idx = len(rows) - 1
                for w in self.row_widgets[last_idx].values():
                    w.destroy()
                self.row_del_buttons[last_idx].destroy()
                self.row_widgets.pop()
                self.row_del_buttons.pop()
                rows.pop()

            self._render_subtotal_row()

        def _delete_row(self, row_idx):
            if row_idx < 0 or row_idx >= len(self.section["rows"]):
                return
            row_dict = self.section["rows"][row_idx]
            if qty_populated(row_dict) or any(str(row_dict.get(k, "")).strip() for k in row_dict.keys()):
                if not messagebox.askyesno("Confirm", "This row has data. Remove it?"):
                    return

            self.push_undo()
            self._remove_subtotal_row_widgets()

            self.section["rows"].pop(row_idx)

            for widget in self.row_widgets[row_idx].values():
                widget.destroy()
            self.row_del_buttons[row_idx].destroy()
            self.row_widgets.pop(row_idx)
            self.row_del_buttons.pop(row_idx)

            for i in range(len(self.row_widgets)):
                grid_row = i + 1
                for c, (_h, key, _w, kind) in enumerate(FS_HEADERS):
                    w = self.row_widgets[i][key]
                    w.grid_configure(row=grid_row, column=c)
                    if kind != "calc":
                        w.unbind("<KeyRelease>")
                        w.bind("<KeyRelease>", lambda ev, r=i, k=key: self._on_edit(r, k))
                    w.unbind("<Return>")
                    w.unbind("<Down>")
                    w.bind("<Return>", lambda ev, r=i, c=c: self._nav_down(r, c))
                    w.bind("<Down>", lambda ev, r=i, c=c: self._nav_down(r, c))
                self.row_del_buttons[i].grid_configure(row=grid_row, column=len(FS_HEADERS))
                self.row_del_buttons[i].configure(command=lambda r=i: self._delete_row(r))

            if not self.section["rows"]:
                self.section["rows"].append(default_schedule_row())
                self._add_row_widgets(0, self.section["rows"][0])

            self._enforce_blank_row()
            self.update_subtotals_and_materials()
            self.on_any_change()

        def _on_edit(self, row_idx, key):
            if not self._alive:
                return
            if row_idx < 0 or row_idx >= len(self.section["rows"]) or row_idx >= len(self.row_widgets):
                return

            row = self.section["rows"][row_idx]
            for (_h, k, _w, kind) in FS_HEADERS:
                if kind != "calc":
                    row[k] = (self.row_widgets[row_idx][k].get() or "").strip()

            if key != "caulk_passes":
                any_other_data = False
                for (_h, k, _w, kind) in FS_HEADERS:
                    if kind == "calc" or k == "caulk_passes":
                        continue
                    if str(row.get(k, "")).strip() != "":
                        any_other_data = True
                        break
                if any_other_data and str(row.get("caulk_passes", "")).strip() == "":
                    row["caulk_passes"] = "3"
                    wcp = self.row_widgets[row_idx]["caulk_passes"]
                    wcp.delete(0, tk.END)
                    wcp.insert(0, "3")

            calc_schedule_row(row)

            for (_h, k, _w, kind) in FS_HEADERS:
                if kind == "calc":
                    ent = self.row_widgets[row_idx][k]
                    ent.config(state="normal")
                    ent.delete(0, tk.END)
                    ent.insert(0, row.get(k, ""))
                    ent.config(state="readonly")

            if key == "qty":
                self._enforce_blank_row()

            self.update_subtotals_and_materials()
            self.on_any_change()

        def _render_materials_block(self):
            self.mat_frame.pack(fill="x", padx=6, pady=(6, 10))
            self.mat_grid = ttk.Frame(self.mat_frame)
            self.mat_grid.pack(fill="x", padx=6, pady=6)

            for c, h in enumerate(["ITEM", "FT/SAUS", "FACTOR", "RATE", "COST", "UNIT", ""]):
                ttk.Label(self.mat_grid, text=h, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=c, padx=3, sticky="w")

            self._rebuild_material_rows()

            ttk.Button(self.mat_frame, text="+ Add Install Material Row", command=self._add_freeform_material_row)\
                .pack(anchor="w", padx=6, pady=(0, 6))

        def _rebuild_material_rows(self):
            for w in list(self.mat_grid.winfo_children()):
                info = w.grid_info()
                if info and int(info.get("row", 0)) == 0:
                    continue
                w.destroy()

            self.mat_qty_vars.clear()
            self.mat_cost_vars.clear()
            self.mat_qty_entries.clear()
            self.mat_factor_entries.clear()
            self.mat_rate_entries.clear()

            mats = self.section["materials"]

            def delete_material_row(index: int):
                if index < 0 or index >= len(mats):
                    return
                m = mats[index]
                # only allow removing "manual/freeform" rows; template rows stay
                if m.get("basis") != "manual" or m.get("key", "").startswith(("bracing", "sheet_metal", "flashing", "backer_rods", "sealants", "tie_back", "backpans")):
                    # allow delete only if this is a user-added freeform row
                    if not m.get("key", "").startswith("free_"):
                        return
                has_data = any(str(m.get(k, "")).strip() for k in ["label", "qty", "factor", "rate", "unit"])
                if has_data and not messagebox.askyesno("Confirm", "This install material row has data. Remove it?"):
                    return
                self.push_undo()
                mats.pop(index)
                self._rebuild_material_rows()
                self.update_subtotals_and_materials()
                self.on_any_change()

            for r_m, m in enumerate(mats, start=1):
                idx = r_m - 1

                lbl = ttk.Label(self.mat_grid, text=m.get("label", ""))
                lbl.grid(row=r_m, column=0, padx=3, sticky="w")

                qty_var = tk.StringVar(value="")
                self.mat_qty_vars[idx] = qty_var
                qty_e = ttk.Entry(self.mat_grid, width=12, textvariable=qty_var)
                qty_e.grid(row=r_m, column=1, padx=3, sticky="w")
                if m.get("basis") != "manual":
                    qty_e.config(state="readonly")
                else:
                    qty_e.config(state="normal")
                    qty_e.delete(0, tk.END)
                    qty_e.insert(0, m.get("qty", ""))
                    qty_e.bind("<FocusIn>", lambda ev: self._mark_undo_once())
                    qty_e.bind("<KeyRelease>", lambda ev, i=idx, w=qty_e: self._material_manual_qty_edit(i, w))
                self.mat_qty_entries[idx] = qty_e

                factor_e = ttk.Entry(self.mat_grid, width=8)
                factor_e.insert(0, m.get("factor", ""))
                factor_e.grid(row=r_m, column=2, padx=3, sticky="w")
                factor_e.bind("<FocusIn>", lambda ev: self._mark_undo_once())
                factor_e.bind("<KeyRelease>", lambda ev, i=idx, w=factor_e: self._material_factor_edit(i, w))
                self.mat_factor_entries[idx] = factor_e

                rate_e = ttk.Entry(self.mat_grid, width=10)
                rate_e.insert(0, m.get("rate", ""))
                rate_e.grid(row=r_m, column=3, padx=3, sticky="w")
                rate_e.bind("<FocusIn>", lambda ev: self._mark_undo_once())
                rate_e.bind("<KeyRelease>", lambda ev, i=idx, w=rate_e: self._material_rate_edit(i, w))
                self.mat_rate_entries[idx] = rate_e

                cost_var = tk.StringVar(value=money_fmt(0))
                self.mat_cost_vars[idx] = cost_var
                ttk.Label(self.mat_grid, textvariable=cost_var).grid(row=r_m, column=4, padx=3, sticky="w")

                ttk.Label(self.mat_grid, text=m.get("unit", "")).grid(row=r_m, column=5, padx=3, sticky="w")

                btn = ttk.Button(self.mat_grid, text="−", width=2, command=lambda i=idx: delete_material_row(i))
                btn.grid(row=r_m, column=6, padx=3, sticky="w")
                if not str(m.get("key", "")).startswith("free_"):
                    btn.state(["disabled"])

            total_row = len(mats) + 1
            ttk.Label(self.mat_grid, text="SUBTOTAL", font=("TkDefaultFont", 9, "bold")).grid(row=total_row, column=0, padx=3, sticky="w")
            e = ttk.Entry(self.mat_grid, width=12, textvariable=self.total_cost_var)
            e.grid(row=total_row, column=4, padx=3, sticky="w")
            e.config(state="readonly")

        def _material_manual_qty_edit(self, i, widget):
            self.section["materials"][i]["qty"] = (widget.get() or "").strip()
            self.update_materials_only()
            self.on_any_change()

        def _material_factor_edit(self, i, widget):
            self.section["materials"][i]["factor"] = (widget.get() or "").strip()
            self.update_materials_only()
            self.on_any_change()

        def _material_rate_edit(self, i, widget):
            self.section["materials"][i]["rate"] = (widget.get() or "").strip()
            self.update_materials_only()
            self.on_any_change()

        def _add_freeform_material_row(self):
            self.push_undo()
            self.section["materials"].append({
                "key": f"free_{uuid.uuid4().hex[:8]}",
                "label": "",
                "basis": "manual",
                "factor": "",
                "rate": "",
                "qty": "",
                "unit": "",
            })
            self._rebuild_material_rows()
            self.update_subtotals_and_materials()
            self.on_any_change()

        def update_subtotals_and_materials(self):
            subs = schedule_subtotals(self.section["rows"])
            self.sub_vars["qty"].set(str(subs["qty"]))
            self.sub_vars["sqft"].set(str(subs["sqft"]))
            self.sub_vars["perim"].set(str(subs["perim"]))
            self.sub_vars["caulk_lf"].set(str(subs["caulk_lf"]))
            self.sub_vars["head_sill"].set(str(subs["head_sill"]))
            self.update_materials_only()

        def update_materials_only(self):
            mats = self.section["materials"]
            subs = {
                "perim": safe_int(self.sub_vars["perim"].get()),
                "caulk_lf": safe_int(self.sub_vars["caulk_lf"].get()),
                "head_sill": safe_int(self.sub_vars["head_sill"].get()),
            }
            total_cost = 0
            for idx_m, m in enumerate(mats):
                if idx_m in self.mat_factor_entries:
                    m["factor"] = (self.mat_factor_entries[idx_m].get() or "").strip()
                if idx_m in self.mat_rate_entries:
                    m["rate"] = (self.mat_rate_entries[idx_m].get() or "").strip()
                if m.get("basis") == "manual" and idx_m in self.mat_qty_entries:
                    m["qty"] = (self.mat_qty_entries[idx_m].get() or "").strip()

                if m.get("key") == "sealants" and str(m.get("factor", "")).strip() == "":
                    m["factor"] = "0.0833"
                    self.mat_factor_entries[idx_m].delete(0, tk.END)
                    self.mat_factor_entries[idx_m].insert(0, "0.0833")

                qty_val = material_qty_for_row(m, subs)
                rate_val = safe_float(m.get("rate", ""))
                cost_int = roundup(qty_val * rate_val) if qty_val * rate_val != 0 else 0

                if m.get("basis") == "manual":
                    self.mat_qty_vars[idx_m].set(m.get("qty", ""))
                else:
                    self.mat_qty_vars[idx_m].set(str(roundup(qty_val)) if qty_val != 0 else "")

                self.mat_cost_vars[idx_m].set(money_fmt(cost_int))
                total_cost += cost_int

            self.section["install_material_total"] = total_cost
            self.install_total_var.set(money_fmt(total_cost))
            self.labelframe.configure(text=self._title_text())
            self.total_cost_var.set(money_fmt(total_cost))

        def recalc_all(self, allow_row_change: bool):
            for i, row in enumerate(self.section["rows"]):
                if i >= len(self.row_widgets):
                    break
                for (_h, key, _w, kind) in FS_HEADERS:
                    if kind != "calc":
                        row[key] = (self.row_widgets[i][key].get() or "").strip()
                calc_schedule_row(row)
                for (_h, key, _w, kind) in FS_HEADERS:
                    if kind == "calc":
                        ent = self.row_widgets[i][key]
                        ent.config(state="normal")
                        ent.delete(0, tk.END)
                        ent.insert(0, row.get(key, ""))
                        ent.config(state="readonly")

            if allow_row_change:
                self._enforce_blank_row()

            self.update_subtotals_and_materials()

    class FrameScheduleManager:
        def __init__(self):
            self.groups = {}            # spec_id -> LabelFrame
            self.group_sections = {}    # spec_id -> list[FrameScheduleSectionUI]
            self.section_ui = {}        # section_id -> FrameScheduleSectionUI

        def _ensure_group(self, spec_id: str):
            if spec_id in self.groups and self.groups[spec_id].winfo_exists():
                return self.groups[spec_id]
            base, var = parse_frame_spec_id(spec_id)
            label = frame_spec_label(base, var)
            compute_frame_schedule_rollups(job)
            rollup_total = job.get("frame_schedule_rollups", {}).get(spec_id, {}).get("install_material_total", 0)
            grp = ttk.LabelFrame(fs_scroll.inner, text=f"{label} — Schedules (Install Materials Total: {money_fmt(rollup_total)})")
            grp.pack(fill="x", padx=4, pady=8)
            self.groups[spec_id] = grp
            self.group_sections[spec_id] = []
            return grp

        def _update_group_title(self, spec_id: str):
            compute_frame_schedule_rollups(job)
            base, var = parse_frame_spec_id(spec_id)
            label = frame_spec_label(base, var)
            rollup_total = job.get("frame_schedule_rollups", {}).get(spec_id, {}).get("install_material_total", 0)
            grp = self.groups.get(spec_id)
            if grp and grp.winfo_exists():
                grp.configure(text=f"{label} — Schedules (Install Materials Total: {money_fmt(rollup_total)})")

        def add_section_ui(self, spec_id: str, section: dict):
            grp = self._ensure_group(spec_id)

            def on_delete(section_id: str):
                self.delete_section(spec_id, section_id)

            def on_any_change():
                recompute_section_totals(section)
                compute_frame_schedule_rollups(job)
                self._update_group_title(spec_id)
                self._reindex_group(spec_id)

            idx = len(self.group_sections[spec_id]) + 1
            ui = FrameScheduleSectionUI(
                grp, spec_id, section, idx,
                on_delete_section=on_delete,
                on_any_change=on_any_change,
                push_undo_fn=push_undo
            )
            self.group_sections[spec_id].append(ui)
            self.section_ui[section["id"]] = ui
            self._update_group_title(spec_id)

        def delete_section(self, spec_id: str, section_id: str):
            job["frame_schedules"][spec_id] = [s for s in job["frame_schedules"][spec_id] if s.get("id") != section_id]
            normalize_frame_schedules(job)

            ui = self.section_ui.pop(section_id, None)
            if ui:
                ui.destroy()

            if spec_id in self.group_sections:
                self.group_sections[spec_id] = [u for u in self.group_sections[spec_id] if u.section.get("id") != section_id]
                self._reindex_group(spec_id)

            if not job["frame_schedules"].get(spec_id):
                grp = self.groups.pop(spec_id, None)
                if grp and grp.winfo_exists():
                    grp.destroy()
                self.group_sections.pop(spec_id, None)

            self._update_group_title(spec_id)

        def _reindex_group(self, spec_id: str):
            if spec_id not in self.group_sections:
                return
            for i, ui in enumerate(self.group_sections[spec_id], start=1):
                ui.set_index(i)

        def reconcile_all(self):
            normalize_frame_schedules(job)
            compute_frame_schedule_rollups(job)

            for spec_id, sections in job.get("frame_schedules", {}).items():
                for s in sections:
                    sid = s.get("id")
                    if sid and sid not in self.section_ui:
                        self.add_section_ui(spec_id, s)

            existing_ids = set()
            for spec_id, sections in job.get("frame_schedules", {}).items():
                for s in sections:
                    existing_ids.add(s.get("id"))
            for sid in list(self.section_ui.keys()):
                if sid not in existing_ids:
                    ui = self.section_ui.pop(sid, None)
                    if ui:
                        ui.destroy()

            for spec_id in list(self.groups.keys()):
                if not job.get("frame_schedules", {}).get(spec_id):
                    grp = self.groups.pop(spec_id, None)
                    if grp and grp.winfo_exists():
                        grp.destroy()
                    self.group_sections.pop(spec_id, None)

            for spec_id in list(self.group_sections.keys()):
                self._reindex_group(spec_id)
                self._update_group_title(spec_id)

    fs_manager = FrameScheduleManager()

    def fs_add_schedule_section():
        # FIX: do NOT call reconcile-all/normalize in a way that churns widgets while
        # existing entries are mid-edit. Add the UI directly.
        try:
            fs_refresh_spec_options()
            lab = (fs_spec_var.get() or "").strip()
            if not lab:
                messagebox.showerror("Error", "Select a spec first (add a cost code if none exist).")
                return
            spec_id = fs_label_to_spec.get(lab)
            if not spec_id:
                messagebox.showerror("Error", f"Invalid spec selection: {lab}")
                return

            push_undo()
            job["frame_schedules"].setdefault(spec_id, [])

            section = default_schedule_section(spec_id)
            # ensure deterministic totals before UI
            recompute_section_totals(section)

            job["frame_schedules"][spec_id].append(section)
            compute_frame_schedule_rollups(job)

            # Add UI incrementally: no teardown, no overwriting previous sections.
            fs_manager.add_section_ui(spec_id, section)

        except Exception as e:
            messagebox.showerror("Error adding schedule section", str(e))

    ttk.Button(selector_frame, text="Add Schedule Section", command=fs_add_schedule_section).pack(side="left", padx=6)

    # Export frame schedules to PDF
    def export_frame_schedule_pdf():
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("PDF Export", "The reportlab library is not installed.\nPlease install reportlab to enable PDF export.")
            return
        collect_job_info_into_job()
        ensure_job_defaults(job)
        normalize_frame_schedules(job)
        compute_frame_schedule_rollups(job)
        default_name = f"{job.get('job_name', 'job')}_frame_schedule.pdf"
        file_path = filedialog.asksaveasfilename(
            title="Save Frame Schedule PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            job_name = job.get("job_name", "").strip() or "Job"
            elements.append(Paragraph(f"Frame Schedules – {job_name}", styles['Title']))
            elements.append(Spacer(1, 12))
            # Sort specs by their user-visible label for consistency
            sorted_specs = sorted(job.get("frame_schedules", {}).keys(), key=lambda x: fs_spec_to_label.get(x, x))
            for spec_id in sorted_specs:
                sections = job["frame_schedules"].get(spec_id, [])
                if not sections:
                    continue
                base, var = parse_frame_spec_id(spec_id)
                spec_label = frame_spec_label(base, var)
                rollup_total = job.get("frame_schedule_rollups", {}).get(spec_id, {}).get("install_material_total", 0)
                elements.append(Paragraph(f"{spec_label} – Install Materials Total: {money_fmt(rollup_total)}", styles['Heading2']))
                elements.append(Spacer(1, 6))
                for idx, section in enumerate(sections, start=1):
                    recompute_section_totals(section)
                    elements.append(Paragraph(f"Schedule {idx} (Install Total: {money_fmt(section.get('install_material_total', 0))})", styles['Heading3']))
                    elements.append(Spacer(1, 4))
                    # Schedule rows table
                    row_headers = [h for (h, _k, _w, _kind) in FS_HEADERS]
                    row_keys = [k for (_h, k, _w, _kind) in FS_HEADERS]
                    table_data = [row_headers]
                    for r in section.get("rows", []):
                        calc_schedule_row(r)
                        row_data = []
                        for k in row_keys:
                            val = r.get(k, "")
                            row_data.append(str(val or ""))
                        table_data.append(row_data)
                    subs = schedule_subtotals(section.get("rows", []))
                    subtotal_row = []
                    for k in row_keys:
                        if k == "spec_mark":
                            subtotal_row.append("Subtotal")
                        elif k in subs:
                            subtotal_row.append(str(subs[k]))
                        else:
                            subtotal_row.append("")
                    table_data.append(subtotal_row)
                    col_widths = [60, 40, 40, 40, 50, 50, 60, 60, 60, 50, 50, 50, 50, 50, 50, 80]
                    sched_table = Table(table_data, colWidths=col_widths)
                    sched_style = TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 8),
                        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
                        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,-1), (-1,-1), 8),
                        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                    ])
                    sched_table.setStyle(sched_style)
                    elements.append(sched_table)
                    elements.append(Spacer(1, 6))
                    # Install materials table
                    mat_data = [["Item", "Qty", "Factor", "Rate", "Cost", "Unit"]]
                    subs_for_basis = {
                        "perim": safe_int(subs.get("perim", 0)),
                        "caulk_lf": safe_int(subs.get("caulk_lf", 0)),
                        "head_sill": safe_int(subs.get("head_sill", 0)),
                    }
                    total_mat_cost = 0
                    for m in section.get("materials", []):
                        qty_val = material_qty_for_row(m, subs_for_basis)
                        rate_val = safe_float(m.get("rate", ""))
                        cost_val = roundup(qty_val * rate_val) if qty_val * rate_val != 0 else 0
                        total_mat_cost += cost_val
                        display_qty = m.get("qty", "") if m.get("basis") == "manual" else (str(roundup(qty_val)) if qty_val else "")
                        mat_data.append([
                            m.get("label", ""),
                            display_qty,
                            str(m.get("factor", "")),
                            str(m.get("rate", "")),
                            money_fmt(cost_val),
                            m.get("unit", ""),
                        ])
                    mat_data.append(["Total", "", "", "", money_fmt(total_mat_cost), ""])
                    mat_table = Table(mat_data, colWidths=[140, 50, 50, 60, 60, 60])
                    mat_table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 9),
                        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
                        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,-1), (-1,-1), 9),
                        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                    ]))
                    elements.append(mat_table)
                    elements.append(Spacer(1, 12))
            doc.build(elements)
            messagebox.showinfo("PDF Export", f"Frame schedule successfully exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("PDF Export", f"Failed to export PDF:\n{e}")

    ttk.Button(selector_frame, text="Export PDF", command=export_frame_schedule_pdf).pack(side="right", padx=6)

    # --------------------------------------------------
    # Saving
    # --------------------------------------------------
    def collect_job_info_into_job():
        for _label, key, _kind in JOB_FIELDS:
            t, w = info_widgets[key]
            if t == "bool":
                job[key] = bool(w.get())
            elif t == "date":
                job[key] = get_date_value(w)
            else:
                job[key] = (w.get() or "").strip()

    def save_only():
        collect_job_info_into_job()
        normalize_quotes(job)
        normalize_frame_schedules(job)
        compute_frame_schedule_rollups(job)
        if not (job.get("job_name") or "").strip():
            return
        save_job(job)
        refresh_main()

    def save_and_close():
        collect_job_info_into_job()
        if not (job.get("job_name") or "").strip():
            messagebox.showerror("Error", "Job Name is required.")
            return
        normalize_quotes(job)
        normalize_frame_schedules(job)
        compute_frame_schedule_rollups(job)
        save_job(job)
        refresh_main()
        win.destroy()

    ttk.Button(topbar, text="Save & Close", command=save_and_close).pack(side="right")

    def autosave_tick():
        if not win.winfo_exists():
            return
        try:
            save_only()
        except Exception:
            pass
        win.after(10000, autosave_tick)

    win.after(10000, autosave_tick)

    # --------------------------------------------------
    # Initial renders
    # --------------------------------------------------
    refresh_cc_tree()
    refresh_quotes_ui(full=True)
    fs_refresh_spec_options()
    fs_manager.reconcile_all()


# ======================================================
# Main App
# ======================================================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Bid Manager")
        root.geometry("520x520")

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="New Job Name").pack(anchor="w")
        self.new_job = ttk.Entry(frm)
        self.new_job.pack(fill="x")

        ttk.Button(frm, text="Create Job", command=self.create_job).pack(pady=6)

        ttk.Label(frm, text="Jobs (double-click to open):").pack(anchor="w", pady=(10, 0))
        self.listbox = tk.Listbox(frm)
        self.listbox.pack(fill="both", expand=True, pady=10)
        self.listbox.bind("<Double-Button-1>", self.open_selected)

        self.refresh()

    def refresh(self):
        self.jobs = load_jobs()
        self.listbox.delete(0, tk.END)
        for j in self.jobs:
            self.listbox.insert(tk.END, (j.get("job_name") or "").strip())

    def create_job(self):
        name = self.new_job.get().strip()
        if not name:
            return
        job = {"id": str(uuid.uuid4())}
        ensure_job_defaults(job)
        job["job_name"] = name
        save_job(job)
        self.new_job.delete(0, tk.END)
        self.refresh()

    def open_selected(self, _event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        job = self.jobs[sel[0]]
        open_job_window(self.root, job, self.refresh)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
