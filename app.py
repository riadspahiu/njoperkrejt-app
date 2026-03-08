"""
NjoPerKrejt - SmartRegister
Version manuale - pa AI, pa internet
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import sys
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG FILE (ruhet prane app.py) ──────────────────────────────────────────
APP_DIR     = Path(os.path.expanduser("~")) / "NjoPerKrejt"
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APP_DIR / "config.json"

VERSION       = "1.5"
GITHUB_USER   = "riadspahiu"
GITHUB_REPO   = "njoperkrejt-app"
GITHUB_BRANCH = "main"
GITHUB_RAW    = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}"
_EXTERNAL_APP = APP_DIR / "app_update.py"
_VER_FILE     = APP_DIR / "installed_version.txt"


def _boot_external():
    """Nese ka app_update.py te shkarkuar, ekzekutoje ate.
    Nese jemi duke u ekzekutuar SI app_update.py, mos bjer ne loop."""
    # Mos ekzekuto veten — parandalon loop te pafund
    current_file = Path(os.path.abspath(sys.argv[0]))
    if current_file.resolve() == _EXTERNAL_APP.resolve():
        return  # jemi tashme app_update.py, vazhdo normalisht

    if _EXTERNAL_APP.exists():
        try:
            # Verifiko qe file-i eshte i vlefshëm (jo bosh/korrupt)
            content = _EXTERNAL_APP.read_text(encoding="utf-8", errors="ignore")
            if len(content) < 100 or "def " not in content:
                _EXTERNAL_APP.unlink()  # fshi file-in e korruptuar
                return
            import subprocess
            subprocess.Popen([sys.executable, str(_EXTERNAL_APP)] + sys.argv[1:])
            sys.exit(0)
        except Exception:
            # Nese deshtoi, fshi app_update.py dhe vazhdo me app.py origjinal
            try:
                _EXTERNAL_APP.unlink()
            except Exception:
                pass


def _get_current_version():
    """Kthe versionin aktual — pershpejton nese installed_version.txt ekziston."""
    if _VER_FILE.exists():
        try:
            v = _VER_FILE.read_text().strip()
            if v:
                return v
        except Exception:
            pass
    # Shkruaj versionin aktual ne file ne here te pare
    try:
        _VER_FILE.write_text(VERSION)
    except Exception:
        pass
    return VERSION


def check_for_update():
    """Kontrollo GitHub per version te ri. Background thread."""
    import threading
    current = _get_current_version()

    def _check():
        try:
            import urllib.request
            with urllib.request.urlopen(f"{GITHUB_RAW}/version.txt", timeout=5) as r:
                latest = r.read().decode().strip()
            # Verifiko qe eshte version i vlefshëm, jo faqe HTML gabimi
            if not latest or len(latest) > 20:
                return
            if latest == current:
                return
            # Mos shfaq nese ky version eshte refuzuar me pare
            if latest == load_config().get("update_skipped_version", ""):
                return
            _prompt_update(latest, current)
        except Exception:
            pass

    threading.Thread(target=_check, daemon=True).start()


def _prompt_update(latest_version, current_version):
    """Shfaq dritaren e update-it ne thread kryesor."""
    def _show():
        win = tk.Toplevel()
        win.title("Update i disponueshem")
        win.geometry("380x180")
        win.configure(bg="#0d0f14")
        win.resizable(False, False)
        win.grab_set()
        win.lift()
        win.focus_force()

        tk.Label(win, text="Update i ri i disponueshem!",
                 font=("Segoe UI", 11, "bold"), bg="#0d0f14", fg="#66FFCC").pack(pady=(22, 4))
        tk.Label(win, text=f"Versioni aktual: {current_version}   ->   {latest_version}",
                 font=("Segoe UI", 9), bg="#0d0f14", fg="#888").pack()
        tk.Label(win, text="Deshiron ta instalosh tani?",
                 font=("Segoe UI", 9), bg="#0d0f14", fg="#ccc").pack(pady=(8, 16))

        btn_row = tk.Frame(win, bg="#0d0f14")
        btn_row.pack()

        def do_update():
            try:
                import urllib.request
                # Shkarko app-in e ri
                urllib.request.urlretrieve(f"{GITHUB_RAW}/app.py", str(_EXTERNAL_APP))
                # KRITIKAL: shkruaj version PARA se te rinisesh
                # Keshtu app_update.py do te lexoje versionin e ri dhe nuk do te shfaqe popup
                _VER_FILE.write_text(latest_version)
                win.destroy()
                # Rinis me app_update.py
                import subprocess
                subprocess.Popen([sys.executable, str(_EXTERNAL_APP)] + sys.argv[1:])
                os._exit(0)
            except Exception as e:
                import traceback
                tk.messagebox.showerror("Gabim", f"Update deshtoi:\n{e}\n\n{traceback.format_exc()}", parent=win)

        def skip_update():
            """Ruaj versionin si 'i skipped' — mos shfaq popup sërish per kete version."""
            try:
                cfg = load_config()
                cfg["update_skipped_version"] = latest_version
                save_config(cfg)
            except Exception:
                pass
            win.destroy()

        tk.Button(btn_row, text="  Po, instalo  ", font=("Segoe UI", 9, "bold"),
                  bg="#66FFCC", fg="#0d0f14", relief="flat", padx=16, pady=8,
                  cursor="hand2", command=do_update).pack(side="left", padx=(0, 10))
        tk.Button(btn_row, text="Jo tani", font=("Segoe UI", 9),
                  bg="#181b23", fg="#888", relief="flat", padx=16, pady=8,
                  cursor="hand2", command=skip_update).pack(side="left")

    try:
        tk._default_root.after(1000, _show)
    except Exception:
        pass


def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except:
            pass
    return {}

def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

def get_base_dir():
    cfg = load_config()
    p = cfg.get("base_dir", "")
    if p and Path(p).exists():
        return Path(p)
    return APP_DIR

def get_biz_dir():
    """Kthen folder-in e bizneseve — mund te jete direkt folder-i i klienteve."""
    cfg = load_config()
    p = cfg.get("biz_dir", "")
    if p and Path(p).exists():
        return Path(p)
    return get_base_dir() / "Bizneset"

def set_base_dir(path):
    cfg = load_config()
    cfg["base_dir"] = str(path)
    save_config(cfg)

def set_dirs(base_path, biz_path=None):
    """Ruan base_dir dhe opsionalisht biz_dir te ndare."""
    cfg = load_config()
    cfg["base_dir"] = str(base_path)
    if biz_path and str(biz_path) != str(base_path):
        cfg["biz_dir"] = str(biz_path)
    else:
        cfg.pop("biz_dir", None)
    save_config(cfg)

SYS_FOLDER = "NjoPerKrejt - Excel [SpahiuDev]"

# ── DYNAMIC PATHS (rikomputon cdo here) ───────────────────────────────────────
def paths():
    biz  = get_biz_dir()
    base = biz / SYS_FOLDER
    lg   = base / "log.txt"
    for d in [base, biz]:
        d.mkdir(parents=True, exist_ok=True)
    return base, biz, lg

# ── THEME SYSTEM ─────────────────────────────────────────────────────────────
THEMES = {
    "dark": {
        "BG":          "#1c1f2e",
        "SURFACE":     "#242838",
        "CARD":        "#2a2f42",
        "CARD2":       "#323750",
        "BORDER":      "#383d54",
        "ACCENT":      "#4f8ef7",
        "ACCENT2":     "#3b7df5",
        "GREEN":       "#34c97e",
        "YELLOW":      "#f5a623",
        "RED":         "#f25f5c",
        "TEXT":        "#eceef5",
        "MUTED":       "#7b82a0",
        "MUTED2":      "#555c78",
        "WHITE":       "#ffffff",
        "INPUT_BG":    "#1e2235",
        "INPUT_FOCUS": "#4f8ef7",
    },
    "light": {
        "BG":          "#f0f2f8",
        "SURFACE":     "#ffffff",
        "CARD":        "#ffffff",
        "CARD2":       "#eef0f8",
        "BORDER":      "#dde0ef",
        "ACCENT":      "#3b7df5",
        "ACCENT2":     "#2d6de8",
        "GREEN":       "#1daa65",
        "YELLOW":      "#e09400",
        "RED":         "#e04e4b",
        "TEXT":        "#1a1d2e",
        "MUTED":       "#7b82a0",
        "MUTED2":      "#adb5d0",
        "WHITE":       "#ffffff",
        "INPUT_BG":    "#f7f8fc",
        "INPUT_FOCUS": "#3b7df5",
    }
}

_CURRENT_THEME = "dark"

def T(key):
    return THEMES[_CURRENT_THEME][key]

# Globals per backward compat — rikomputon me set_theme()
BG = SURFACE = CARD = BORDER = ACCENT = GREEN = YELLOW = RED = ""
TEXT = MUTED = WHITE = INPUT_BG = INPUT_FG = INPUT_FOCUS = ""
H_BG = "1a3a5c"
ROW_ALT = "f0f4f8"
CARD2 = MUTED2 = ACCENT2 = ""

def set_theme(name):
    global _CURRENT_THEME
    global BG, SURFACE, CARD, CARD2, BORDER, ACCENT, ACCENT2
    global GREEN, YELLOW, RED, TEXT, MUTED, MUTED2, WHITE
    global INPUT_BG, INPUT_FG, INPUT_FOCUS
    _CURRENT_THEME = name
    t = THEMES[name]
    BG = t["BG"]; SURFACE = t["SURFACE"]; CARD = t["CARD"]; CARD2 = t["CARD2"]
    BORDER = t["BORDER"]; ACCENT = t["ACCENT"]; ACCENT2 = t["ACCENT2"]
    GREEN = t["GREEN"]; YELLOW = t["YELLOW"]; RED = t["RED"]
    TEXT = t["TEXT"]; MUTED = t["MUTED"]; MUTED2 = t["MUTED2"]; WHITE = t["WHITE"]
    INPUT_BG = t["INPUT_BG"]; INPUT_FG = t["TEXT"]; INPUT_FOCUS = t["INPUT_FOCUS"]

set_theme("dark")  # default

FONT   = ("Segoe UI", 11)
FONT_B = ("Segoe UI", 11, "bold")
FONT_S = ("Segoe UI", 10)
FONT_H = ("Segoe UI", 13, "bold")
FONT_XS = ("Segoe UI", 9)

# ── LOGGING ───────────────────────────────────────────────────────────────────
def log(msg):
    _, _, lg = paths()
    ts = datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
    with open(lg, "a", encoding="utf-8") as f:
        f.write(f"[{ts}]  {msg}\n")

# ── EXCEL ─────────────────────────────────────────────────────────────────────
HEADERS_BIZNES  = ["Data", "Ora", "Produkti/Sherbimi", "Sasia", "Njesia",
                   "Cmimi/Njesi (EUR)", "Pagesa (EUR)", "Totali (EUR)", "Shenime"]
def style_header(ws, headers):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", size=11)
        cell.fill      = PatternFill("solid", fgColor=H_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = Border(
            bottom=Side(style="medium", color="FFFFFF"),
            right=Side(style="thin",   color="CCCCCC"))
    ws.row_dimensions[1].height = 28

def auto_width(ws):
    for col in ws.columns:
        mx = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 35)

def ensure_biznes(name):
    _, biz, _ = paths()
    path = biz / f"{name.strip().replace(' ', '_')}.xlsx"
    if not path.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dergesat"
        style_header(ws, HEADERS_BIZNES)
        wb.save(path)
        log(f"File i ri u krijua: {path.name}")
    return path

def append_row(path, row_data, headers):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    if ws.max_row == 0 or ws.cell(1, 1).value != headers[0]:
        style_header(ws, headers)
    nr = ws.max_row + 1
    for col, val in enumerate(row_data, 1):
        cell = ws.cell(row=nr, column=col, value=val)
        cell.alignment = Alignment(vertical="center")
        if nr % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=ROW_ALT)
        cell.border = Border(bottom=Side(style="hair", color="CCCCCC"))
    auto_width(ws)
    wb.save(path)

# ── COLUMN DETECTION & SMART MAPPING ─────────────────────────────────────────
FIELD_ALIASES = {
    "date":     ["data", "date", "dt", "datum", "dita"],
    "ora":      ["ora", "koha", "time", "ore", "hour"],
    "biznesi":  ["biznesi", "bizneset", "klienti", "kliente", "firma",
                 "kompania", "emri", "business", "client", "name"],
    "produkti": ["produkti", "produkte", "sherbimi", "sherbime", "artikulli",
                 "artikuj", "product", "service", "item", "pershkrimi",
                 "description", "mall"],
    "sasia":    ["sasia", "quantity", "qty", "sasi"],
    "njesia":   ["njesia", "njesi", "unit", "measure"],
    "cmimi":    ["cmimi", "price", "cmim", "cena", "tarifa",
                 "cmimi/njesi", "price/unit", "cmimi per njesi"],
    "pagesa":   ["pagesa", "payment", "paid", "paguar",
                 "pagesa e marre", "total paguar"],
    "total":    ["total", "totali", "shuma", "gjithsej",
                 "totali i fatures", "grand total"],
    "shenime":  ["shenime", "notes", "note", "koment", "remarks", "info"],
}

def detect_column_mapping(xlsx_path):
    """Lexon headers e file-it dhe kthen mapping { field: col_index_1based }"""
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb.active
        first_row = [str(c.value or "").strip().lower()
                     for c in next(ws.iter_rows(min_row=1, max_row=1))]
        original  = [str(c.value or "").strip()
                     for c in next(ws.iter_rows(min_row=1, max_row=1))]
        wb.close()
    except:
        return {k: None for k in FIELD_ALIASES}, []

    mapping = {}
    for field, aliases in FIELD_ALIASES.items():
        found = None
        for idx, hdr in enumerate(first_row):
            if hdr in aliases:
                found = idx + 1
                break
        mapping[field] = found
    return mapping, original

def scan_folder_columns(folder_path):
    """Skanon te gjitha .xlsx ne folder dhe kthen mappings"""
    result = {}
    for f in Path(folder_path).glob("*.xlsx"):
        if SYS_FOLDER in str(f):
            continue  # Kapërcen file-at e sistemit
        mapping, headers = detect_column_mapping(f)
        result[f.name] = {"mapping": mapping, "headers": headers, "path": f}
    return result

def append_row_mapped(path, entry_dict, mapping):
    """Shton rresht duke respektuar kolonat ekzistuese te file-it.
    Nuk shton kolona te reja — vetem ploteson ato qe gjenden.
    Nese add_timestamp=True ne config, shton kolonën Data/Ora ne fund (vetem here e pare).
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    nr = ws.max_row + 1
    max_col = ws.max_column or 0

    for field, col_idx in mapping.items():
        if col_idx is None:
            continue
        val = entry_dict.get(field, "")
        if val is None:
            val = ""
        cell = ws.cell(row=nr, column=col_idx, value=val)
        cell.alignment = Alignment(vertical="center")
        if nr % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=ROW_ALT)
        cell.border = Border(bottom=Side(style="hair", color="CCCCCC"))

    # Shto Data/Ora vetem nese perdoruesi ka zgjedhur "Po"
    cfg = load_config()
    if cfg.get("add_timestamp", False):
        ts_label = "Data/Ora e shtimit"
        # Kontrolloj nese kolona ekziston tashmë
        ts_col = None
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(row=1, column=c).value or "").strip() == ts_label:
                ts_col = c
                break
        if ts_col is None:
            # Krijo kolonën herën e parë
            ts_col = ws.max_column + 1
            hcell = ws.cell(row=1, column=ts_col, value=ts_label)
            hcell.font = Font(bold=True, color="FFFFFF", size=11)
            hcell.fill = PatternFill("solid", fgColor="1e3a5f")
            hcell.alignment = Alignment(horizontal="center", vertical="center")
        ts_val = entry_dict.get("date", "") + "  " + entry_dict.get("ora", "")
        cell = ws.cell(row=nr, column=ts_col, value=ts_val.strip())
        cell.alignment = Alignment(vertical="center")
        if nr % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=ROW_ALT)

    auto_width(ws)
    wb.save(path)

def save_entry(entry):
    b   = entry["biznesi"].strip()
    now = datetime.now()
    dt  = entry["date"] or now.strftime("%d/%m/%Y")
    ora = now.strftime("%H:%M:%S")

    full_entry = dict(entry)
    full_entry["date"] = dt
    full_entry["ora"]  = ora

    # ── File i biznesit — smart mapping ──
    bpath = ensure_biznes(b)
    mapping, orig_headers = detect_column_mapping(bpath)

    # Nese file eshte i ri (vetem header-i yne) — perdor format te sistemit
    any_mapped = any(v is not None for v in mapping.values())
    if any_mapped:
        append_row_mapped(bpath, full_entry, mapping)
    else:
        row_b = [dt, ora, entry.get("produkti",""), entry.get("sasia",""),
                 entry.get("njesia",""), entry.get("cmimi",""), entry.get("pagesa",""),
                 entry.get("total",""), entry.get("shenime","")]
        append_row(bpath, row_b, HEADERS_BIZNES)

    log(f"Dergese u ruajt: {b}  |  {entry.get('produkti','')}  |  EUR {entry.get('pagesa','')}")
    return bpath

_FIN_KEYWORDS = ["fitim", "shpenzim", "total", "pagese", "pagesa", "vlera",
                 "cmim", "çmim", "sasi", "shuma", "revenue", "profit",
                 "expense", "amount", "price", "cost", "total", "euro", "eur"]

def _is_financial_col(name: str) -> bool:
    n = name.lower().replace("ë","e").replace("ç","c")
    return any(k in n for k in _FIN_KEYWORDS)

def get_stats():
    """Lexon statistikat direkt nga file-at e bizneseve."""
    _, biz_dir, _ = paths()
    total_rows   = 0
    total_biz    = 0
    fin_cols     = {}   # { col_name: total_value }

    for f in biz_dir.glob("*.xlsx"):
        if SYS_FOLDER in str(f):
            continue
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            ws = wb.active
            headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
            fin_idx = {i: h for i, h in enumerate(headers) if h and _is_financial_col(h)}
            rows = 0
            for r in ws.iter_rows(min_row=2, values_only=True):
                if any(r):
                    rows += 1
                    for i, col_name in fin_idx.items():
                        try:
                            val = float(r[i]) if r[i] is not None else 0
                            fin_cols[col_name] = fin_cols.get(col_name, 0) + val
                        except (TypeError, ValueError):
                            pass
            if rows > 0:
                total_rows += rows
                total_biz  += 1
            wb.close()
        except:
            pass
    return {"dergesa": total_rows, "biznese": total_biz, "fin_cols": fin_cols}

# ── COLUMN PREVIEW WINDOW ────────────────────────────────────────────────────
def show_column_preview(folder_path):
    """
    Shfaq shkurtimisht file-at e lexuara — pa pyetje per timestamp.
    """
    p = Path(folder_path)
    # Prioritet: skanon direkt folder-in e zgjedhur, kapërcen SYS_FOLDER
    files = [f for f in p.glob("*.xlsx") if SYS_FOLDER not in str(f)]
    # Nëse nuk ka direkt, provo nënfolderin Bizneset
    if not files:
        biz_dir = p / "Bizneset"
        files = list(biz_dir.glob("*.xlsx")) if biz_dir.exists() else []
        files = [f for f in files if SYS_FOLDER not in str(f)]
    files = files[:20]

    if not files:
        return  # Asnje file ekzistues — nuk nevojitet preview

    win = tk.Tk()
    win.title("NjoPerKrejt - SmartRegister")
    win.geometry("820x580")
    win.configure(bg=BG)
    win.update_idletasks()
    x = (win.winfo_screenwidth()  - 820) // 2
    y = (win.winfo_screenheight() - 580) // 2
    win.geometry(f"820x580+{x}+{y}")

    tk.Label(win, text=f"U lexuan {len(files)} file Excel",
             font=("Segoe UI", 11, "bold"), bg=BG, fg=WHITE).pack(pady=(16,2))
    tk.Label(win, text="Kolonat e formularëve të tu janë lexuar — asgjë nuk do të ndryshohet.",
             font=("Segoe UI", 9), bg=BG, fg=MUTED).pack()

    # Scrollable area
    container = tk.Frame(win, bg=BG)
    container.pack(fill="both", expand=True, padx=16, pady=10)
    canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
    sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=BG)
    wid = canvas.create_window((0,0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    for f in sorted(files):
        row = tk.Frame(inner, bg=SURFACE, padx=14, pady=8,
                       highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", pady=(0, 4))
        tk.Label(row, text="📄", font=("Segoe UI", 10),
                 bg=SURFACE, fg=MUTED).pack(side="left", padx=(0, 8))
        tk.Label(row, text=f.stem.replace("_", " "),
                 font=("Segoe UI", 9, "bold"), bg=SURFACE, fg=WHITE).pack(side="left")
        tk.Label(row, text="✔ lexuar", font=("Segoe UI", 8),
                 bg=SURFACE, fg=GREEN).pack(side="right")

    # Buton i thjeshtë Vazhdo
    tk.Button(win, text="  VAZHDO  ▶", font=FONT_B,
              bg=ACCENT, fg="#0d0f14", relief="flat", pady=12,
              cursor="hand2", command=win.destroy).pack(fill="x", padx=16, pady=(8,12))

    win.mainloop()


# ── SETUP WIZARD (here e pare) ────────────────────────────────────────────────
def _detect_biz_count(p: Path) -> int:
    """Numeron sa file .xlsx ka brenda (direkt ose ne Bizneset/)."""
    p = Path(p)
    biz_sub = p / "Bizneset"
    files = list(biz_sub.glob("*.xlsx")) if biz_sub.exists() else []
    if not files:
        files = [f for f in p.glob("*.xlsx") if SYS_FOLDER not in str(f)]
    return len(files)


def run_setup_if_needed():
    cfg = load_config()
    if cfg.get("base_dir"):
        return  # tashmë konfiguruar

    # ── Dritarja e Setup-it ──────────────────────────────────────────────────
    root = tk.Tk()
    root.title("NjoPerKrejt - SmartRegister  |  Konfigurimi i Parë")
    root.geometry("560x420")
    root.configure(bg=BG)
    root.resizable(False, False)
    root.update_idletasks()
    x = (root.winfo_screenwidth()  - 560) // 2
    y = (root.winfo_screenheight() - 420) // 2
    root.geometry(f"560x420+{x}+{y}")

    # ── Titulli ──────────────────────────────────────────────────────────────
    tk.Label(root, text="  NjoPerKrejt", font=("Segoe UI", 15, "bold"),
             bg=BG, fg=WHITE).pack(pady=(24, 2))
    tk.Label(root, text="Mirë se erdhe!  Konfiguro folder-in e punës njëherë e mirë.",
             font=("Segoe UI", 9), bg=BG, fg=MUTED).pack()

    # ── Karta me udhëzime ────────────────────────────────────────────────────
    info = tk.Frame(root, bg="#0d2137", padx=16, pady=12,
                    highlightbackground="#1d4ed8", highlightthickness=1)
    info.pack(fill="x", padx=24, pady=(14, 6))

    steps = (
        "1.  Klikoni  «Zgjidh Folder»  dhe gjeni folder-in ku i keni bizneset.",
        "2.  P.sh.  D:\\Klientet   ose   C:\\Punë\\Bizneset",
        "3.  Klikoni  «FILLO PUNËN»  —  gjithçka tjetër bëhet vetë.",
    )
    for s in steps:
        tk.Label(info, text=s, font=("Segoe UI", 9),
                 bg="#0d2137", fg="#93c5fd", anchor="w", justify="left").pack(
                 fill="x", pady=1)

    # ── Inputi i folder-it ───────────────────────────────────────────────────
    frame = tk.Frame(root, bg=CARD, padx=20, pady=14,
                     highlightbackground=BORDER, highlightthickness=1)
    frame.pack(fill="x", padx=24, pady=6)

    tk.Label(frame, text="Folder-i i zgjedhur:", font=FONT_S, bg=CARD, fg=MUTED).pack(anchor="w")

    row = tk.Frame(frame, bg=CARD)
    row.pack(fill="x", pady=(4, 0))

    path_var   = tk.StringVar(value="")
    status_var = tk.StringVar(value="")

    ent = tk.Entry(row, textvariable=path_var, font=FONT_S,
                   bg=SURFACE, fg=TEXT, insertbackground=WHITE,
                   relief="flat", highlightbackground=BORDER, highlightthickness=1)
    ent.pack(side="left", fill="x", expand=True, ipady=7)

    def _update_status(*_):
        p = path_var.get().strip()
        if not p:
            status_var.set("")
            btn_ok.config(state="disabled", bg="#374151")
            return
        pp = Path(p)
        if not pp.exists():
            status_var.set("⚠  Folder-i nuk ekziston — do të krijohet automatikisht.")
            lbl_status.config(fg=YELLOW)
            btn_ok.config(state="normal", bg=GREEN)
            return
        n = _detect_biz_count(pp)
        if n > 0:
            status_var.set(f"✔  U gjet {n} biznes(e) në këtë folder — gati për punë!")
            lbl_status.config(fg=GREEN)
        else:
            status_var.set("ℹ  Folder bosh — bizneset do të krijohen kur të shtosh të parat.")
            lbl_status.config(fg=MUTED)
        btn_ok.config(state="normal", bg=GREEN)

    path_var.trace_add("write", _update_status)

    def browse():
        p = filedialog.askdirectory(title="Zgjidh folder-in ku i ke bizneset")
        if p:
            pp = Path(p)
            # Nëse klienti zgjodhi direkt folder-in "Bizneset", ngjitu një nivel
            if pp.name.lower() == "bizneset":
                pp = pp.parent
            path_var.set(str(pp))

    tk.Button(row, text="📁  Zgjidh Folder", font=FONT_S, bg=ACCENT, fg=WHITE,
              relief="flat", cursor="hand2", padx=10,
              command=browse).pack(side="right", padx=(8, 0), ipady=7)

    lbl_status = tk.Label(frame, textvariable=status_var,
                          font=("Segoe UI", 8), bg=CARD, fg=MUTED, anchor="w")
    lbl_status.pack(fill="x", pady=(6, 0))

    # ── Butoni Vazhdo ────────────────────────────────────────────────────────
    btn_ok = tk.Button(root, text="FILLO PUNËN  ▶", font=("Segoe UI", 11, "bold"),
                       bg="#374151", fg=WHITE, relief="flat", pady=13,
                       cursor="hand2", state="disabled")
    btn_ok.pack(fill="x", padx=24, pady=(10, 0))

    def confirm():
        p = Path(path_var.get().strip())
        try:
            p.mkdir(parents=True, exist_ok=True)
            # Detekto nëse klienti zgjodhi direkt folder-in me xlsx të bizneseve
            direct_xlsx = [f for f in p.glob("*.xlsx") if SYS_FOLDER not in str(f)]
            has_biz_sub = (p / "Bizneset").exists()
            if direct_xlsx or (not has_biz_sub and p.name.lower() not in ("", "njoperkrejt", "njo per krejt")):
                # Ky folder ËSHTË folder i bizneseve — përdore direkt, mos krijo nënfolder
                base_p = p.parent
                base_p.mkdir(parents=True, exist_ok=True)
                set_dirs(base_p, biz_path=p)
            else:
                # Folder i ri bosh — struktura standarde me nënfolder Bizneset
                set_dirs(p)
                (p / "Bizneset").mkdir(exist_ok=True)
            root.destroy()
            show_column_preview(p)
        except Exception as ex:
            messagebox.showerror("Gabim", f"Folder nuk u krijua:\n{ex}")

    btn_ok.config(command=confirm)

    tk.Label(root, text="Ky konfigurim bëhet vetëm njëherë.  Mund ta ndryshosh më vonë te  Cilësimet.",
             font=("Segoe UI", 8), bg=BG, fg="#4b5563").pack(pady=(8, 0))

    root.mainloop()

# ── APP ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # Ngarko temën nga config
        cfg = load_config()
        saved_theme = cfg.get("theme", "dark")
        set_theme(saved_theme)
        self.title("NjoPerKrejt - SmartRegister")
        self.geometry("860x660")
        self.minsize(760, 580)
        self.configure(bg=BG)
        self._build()

    def _build(self):
        set_theme(_CURRENT_THEME)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=SURFACE, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        # Emri i programit
        tk.Label(hdr, text="  NjoPerKrejt",
                 font=("Segoe UI", 12, "bold"), bg=SURFACE, fg=TEXT).pack(side="left", padx=(16,0))

        # Logo — ngarkohet nga base64 e embed-uar
        try:
            import base64 as _b64, io as _io
            from PIL import Image as _Img, ImageTk as _ITk
            _LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAN8AAAAWCAYAAABAHklQAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9ArFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTmlyQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKeDHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAAcMklEQVR4nO2ce3RU1dn/P3ufy8wkk4QECDEGCERF8QLWYhVEUbCIthYFBF6BNj+5qKgVrbbaKt4t0mpbXbQLVMSflIBQRbuEKmC5CAqtIigoIIECiQYTQi5k5lz2/v0xOYcJBAh9318v7+qz1lmTzNm38+z9nO/3efazR0gptWEY3HbbbYwYMYKuXbuilKK8vJyXX36Z559/HiEEWmtaEyklSilKSkq4++678TwP0zR5/fXXefvtt8P7rdUBaNeuHb1796Znz54UFBSQl5eHEAKA+vp6qqqqKC8vZ8OGDezduxcA0zTxPK/V8QSSPmYpJeeccw59+vShsLCQ9u3bY1nWUXWUUkgp+fjjj5k1axYAw4cPZ8CAAeG9topSCsMwmDt3Lu+99x4AnTt35r777sP3fQzD4I9//CNvvfVW2K5Siu9///t861vfQmtNIpHg4Ycfpq6u7rhz0NqzA/zoRz+iuLgYIQRff/01jzzyCJ7nHbMtKSVaa/r378/IkSPDZ5g+fTrl5eUAXHzxxUyYMIH6+npM08R1XR599FFqamoQQhw11+liGAa+71NUVMS9996LlBLf98nKymLWrFmsW7cuXBsCMEyJ56X+7jroAroO7Edm90K0KeFYqtAawzA4tDfBjl9vQx4SFN7clfxe2TiOQksNiDbp8b8jmlQ3wvVp2F3J3rfWUr7yLwBYhsRXCtM0TX7/+98zbNiwFpWLioro378/F154IRMnTmzViODwRJ966qnccsst4fcVFRWtGl8wAQUFBfzwhz9k9OjRdO3a9YQPU11dzYoVK5g+fTobNmw45nigpXGPGzeOyZMnc9555xGNRk/YD8CSJUuYOXMmAAMHDmTSpEltqteabNq0KTS+/Pz8Fjo6cODAUcY3ZMgQRo4cCYDWmmnTplFXV3dSfQZGMHbsWM4991wAqqqqeOKJJ4770pJS4nkevXr14tZbbw2/f+WVVygvL8eyLLZu3Uq3bt0YMGBAeL+kpIRrr70WIcRxDVspRUZGBmVlZfTr1y+8t2LFCrZu3Roa/+GxKHJPL+KKZ35C5yF98WUEj8B00vsQ4f8ajcQm8dkBvniyCX1Q0vVb/SkcXEIDjRhh+WNq75htn4zI5loCEAjOvvsH7F+6jj/f/XOqP9+DsCTmzTffzLBhw/A8Dylli4Xg+z4TJkxg3bp1zJ49OzSc1sTzPHzfx3EcbNsmmUweVSZArH79+lFWVkZRUVFYV2vdKrIERtS+fXtGjBjB8OHDueuuu/jVr37V6niklAghyM7OZs6cOQwdOhRILWTHccL7rUmASI2NjeF3hw4dwvf9ENGBNqFQUD5dD67rtmirqanpqHpBfwB1dXXHRZJjSTC2hoaGFm21FTmTyWSLcbquG7ZbW1vLkCFDWLVqFX369CGRSPCd73yHp59+milTpmDbNo7jtGhPCBEa9ty5c+nXrx9NTU3EYjHWrFnDkCFDcBwn1KuQKWTIP+80rv/TLKyCQurUQfAdBILj8Q+tNFIm8BoSgAbhcagpSZ2qQ7hNuIb8RwBf6rm1RguBD0S1oOiaQYz45tksvGo8NRt3IAN6ESho3bp1LFmyJDREpRQ//vGPsSwrLNdqR0JgGEZ4HVkuoBhnnXUWb731FkVFRTiOg9Ya0zSxLAvDMJBStvi0LCukiMlkEq01zzzzDDfffHNoLOljCCZw/vz5DB06FMdxcF0XIQS2bWOaZotxtnalvwSCcRz5bIF+0uscq+yxdNTay+bINv478ve2day5DF6QiUSCESNGUFlZSTQaJZlMcuedd3LrrbfiOE4LSh+05Xkezz77LEOHDiWRSBCLxSgvL2fEiBHhS1GjEQLQGjMryuC5TyALCkgkqzGEwBYGpjBAmiCtZnQRaG2AFqAFQqewxm2ml4b2kFJhSY2QBhgWQkqQGo2B1jKsixZoLRFaIEnNsSFMhBT4CPzmtk1hIqVANM95CFqGgZRmym4Aw7eQSmIa4BmC+mQ1dqcChsx5DCM7hllQUBAugvLyci677DJc1+WVV17hxhtvxHEcevTowdVXX83ixYuPi34nEsMweOGFF8jOzsZ13RAh586dy9KlS9m9ezeNjY2hAdm2TceOHenbty+lpaV06dIlRMmnn36aFStWsH379vAlERj4HXfcwVVXXRVOqmmarFy5krKyMrZs2UJtbe0xn0EIQV1dXUiBpk2bxvPPPx+iRoBmpaWl3Hvvvbiui2VZTJw4kTVr1mDbdti2lJJ9+/YdlyL/u4lK+Srs3r2b66+/nnfffRfLsvA8j9/85jds27aNZcuWhSwnQM577rmH2267DcdxiEQi1NTUcO211/Lll18eXlMChGGgPJ9zx36HTuf0ptapRkQsRJqrplEpfRoRTCQ2CpWGhxITL6sRJcCTFgKRIo/CQAsH7SsMYWEaEoFBQBJ1M0kEhYOL8l20NLG1S4aRgUCjSdKkfUwMQIWE9DCn0PgoYkYE3wALg6R2EYBpWzR6B8g7rxfn/5/rUj5fINXV1SH9nDZtGiNHjgwNc8qUKSxevLjN1CVdAuUOGjSIiy++GNd1MQyDqqoqrrvuOtauXXvc+m+//TYzZsxg/vz5DBgwgGQySSwW4/bbb+f2228PqWTgU0yZMgXd7HgbhsH999/Pk08+eVJjDt72lZWVVFZWHnV/z549wGGKt337dj777LNW2zqZQM2/gwRG9f7771NaWsq8efNCxCsrK6Nv375s27aNSCRCMplk9OjRPPXUU+G8+77PDTfcwCeffIJlWSGtRadoowCKhw2iSbsIkdKdQKJQ+MJFYhI34iT3N1C9owZV44Av0EKA0hiWTWJHPQYCqfxmo0p1oJUgamShPU3tx1W4XybxXR8tQGiBwEBmQbwkm5wuedRThydsalbuxv1KYHcxyL2oU8r408idLwRR5eMJjSEiVK/cS/IrjXkK5PQ/BZrXiSEMEtql+L+u4rDlcTjaJaVk8+bNLFmyhO9+97s4jsOll17KJZdcwpo1a/5u9Bs7dixa67CPO++8k7Vr14Zo0ZphB7SlqqqKG2+8kc2bN9OuXTu01gwfPpz7778/jLx5nseAAQMoLi4Ofc/XX3+dJ598MqRPQf/Hk/QyAZUNJOjHtu0WdSKRSEjz0nXz97ys/h3E87zQ2EpKSnjsscdIJpO0b9+eP/zhD/Tv358DBw4wYMAAZs+eHerEMAxKS0tZvnx5S8MDkAKtFHZmhJzuRXjCA0MgdApZlNCYyiKmMtj8zAZ2zy5HVUgSbiOExEIgUEhDE4tn0WAkUoYFKKXJME3qNtXw0ZT3aNh0iETCA09BM64hBJZtQY6i+9gz6Pmzs5DRCJ+8uZfyX24hWpLF1SuHoU9VCCVCAzS1RmkfLeJ4W+tYPXIVia9czn7wPE7p351G1ZB6CUuBEC4ZZ3Ru3XcNFtsvf/nLFt/dddddJz1JQgh838eyLPr06RP6Xjt27GDhwoUYhhEGIpRSR11BEMeyLCoqKli0aFFILwsKCjjzzDPDSQXo169fC+N57rnnQmQMgkKt9ZN+pRuM1vqEZU6m3P8mCRDw8ccfZ/bs2UQiEZqamjj77LOZOXMmvXv3ZtGiRUQikZCeP/DAA7z00ktHGx4c5m62hbRMRAsyJ5C+wpRRPvrZOjbf/ynOHhfPrierMJOszhnEO2eQVRQj1jWOeUoWoLH9lA+nAWkIDtVpPhj/Z2pXN+J7mowOFlldMskqyiCza4ScUzOxLI150OLTn29mz0sVSDxOLz2T7KKOJL/0KH9jF1EiR7kSnhZEhMWOss+RBwTZZ2Zz1vgzOUQjsoWpacxopCXyBeL7PlJKVq1axdq1a+nbty+e53HNNdfQs2fPMCzcFj8mfSuiuLg49M3Wrl0bTkhbUFRrjRCClStXMmHCBHzfxzRNevXqxYYNG8JF3qtXL4QQRCIRqqur2bhx4/96I/hnidY6nIdJkyZRXFzM5Zdfjud5DB8+nKuvvpqMjAw8zyMajTJz5kwee+yx1g2PtKC+1qHZBb6e9F1sM8aBv1aw44XPyYpbtL8oj+4PfovMLhqlbRAarcCyoGljkj+PX4bEbG5EYRo21Wt20ri1ESM3yul3n0XRyO4gFIhUsEYIOLTnIJt/+Be8zU3sWb6LU2/uTrue7WjXrz3JBQ3sWbSDbjd1Q1ji8H6CEpiGib+/jsrX/4avBJ2v7ILdOZMGL4GZFu8Knu2YzkhAQZ955hmAkGrdcccdx9wWOJ7k5uYSiRx+W3zxxRfhBLZFAgMKNnuD/jt27Niinby8vLBOZWUlNTU1JzXO/8jJSYD4rusyatQoPv300zC6mZGRgeu6mKbJ0qVLueWWW8J7JysKhYdB5ZJKzBoT1dHinOf70uHibOxTo9hFJpFTTSKdTcwCE6MzmL5ACB+dIpQIoKEyiWoyyCzK5My7exLtKol0iRDtbBPpIhGdBaf27U7+4CLcpELXeSgFSvgUjimGuEXDX2upXrUfW9hopRAIPO1jiCh7Fu+habuD2V5SMqYYlyS2bt1WjmlBAfq98cYbbN26lUgkgu/7jB49msLCwvD+iSRAvqysrJQSm43v4MGDJ6f9Zqmvr2/Rbrt27YDDxpeTkxOWbWhoCBHzP8j3/0+UUliWRVVVFffcc0/oJwfuxtatWxk6dGibfe5WRVhIXOq3N+GrQ+Sen0f81Ewcp5GkMvCVwlMaz9NIbeF6HggfoSWHt+UFlqtQPpDhk/yqCbca3K8dnK9d3K99xNeSqi2V7F9bhS1srI4mlhS4vqbwigIyemXh1Sl2zdvRHEMVKBSWlMikT/n8LxBJyLk0n5w+HXB9UObh1IB0aZV2AmG00HEcnn32WWbMmIHneWRnZzNp0iSmTp2KYRj/9EXdmu/1H/nHipQS13Xp0KEDTzzxRIs5UErRqVMn+vXrx4oVK05+7zJYs8JDY+LUJzAwsPNMtDbRUiOFwqB5s0BotPBQovmbZjqpkQg8PEwsS+HuVLzbfwlKmAiRAoTmEBt+TQKvzsXL8uk++gx8NL4P0QyTzkO78Nn6WqreraRxRz2xkkwcN4FhR/nqzxXU/rUGGVecPuostFD4WqSobytZMseFLt/3EUIwd+5c9u3bh23bKKUYP348OTk54f22SIBYAVoGSNhWCfrJzs4GDhtZgKDB/fRUrHg8HtLnto7zX03+1ccdoJxpmpSVldG7d+8w+cEwDJRS5OXlsXDhQnr06HFUYkRbRSORCCyhcZF4nkA1G1fKzCAVsRT4KGzfAASe9JDCx2jGKbRACcDw0dlgxDUyk9QVF8g4ZJ6ZQYfRnblswdV0uqYzjnIwDIGDS9H1XYgVZuDvTbB74U6EMJFaAh47/+9uOCjIOacd+YM7ktBJIkKAaD2mcUzkA8Lsk7q6OmbNmsVDDz1EMpmksLCQsWPH8txzz52QwwdGUltbSzKZDLMfSkpKTkr5wSQXFxcDh+nr/v37gcNGXV1dHfZbWFhIbm5u+N1/5H9egvl/8cUXGThwYJg29v777/OnP/2JqVOn4jgOubm5LbYg2px4EARcFCipiBRkoCW4OxpSeZrCQHkeqdQYgWqOiIpDjaikQJoGxAx8NGZq+x3luNjFNle8eQ1J0wUl0QIkGlNplCkgQ6DwaVKHsJtR1Pc0md2yyP12expfbGDf4n2cfktPyDJo+KSJr9/9Em1D0ciuiEwD7ZEyci1bzWg7odMWpJTNmjWLgwcPYlkWWmtuu+02IpHICZ3nwPj27t3Lrl27QiO55JJLME2zzagU+AtXXHEFcHhrYdOmTS3Kbdy4MczjzMvL4xvf+Ea4V/jvIIE+09Pu/l70S6d/wR7kieTIfc3jSZDZ8uCDD1JaWkoikSAajVJVVcW4ceN46KGHePHFF7Ftm6amJnr27ElZWVm49XNkP2EisgB5BEuTSiDQdLigA0YkwsFPD/JF2TbiMpeIZRIxo0TMCLFIlAyRwY7XtiETAhk3yCyKp1ASjcqwkVEbZ5/m0IGDRDItrCyBHReYcYnOlugMAwFYSCIyE0f4eIJmhJMU/Vd37CyLhk/rqXqngg6yHeWvbMffnyCza4wuw7uh8TGkgRL6qGxUDfjiBMgHh4/GVFRU8MorrzB58uQw5Wzo0KHMnz8/nLTWJFhEruuyYcMGzjjjDFzX5bTTTmPEiBHMmzevRUpWer2gzcD3LCoq4rrrrgvHVFlZyZYtWwDC+mvWrAlzLyGVmfPOO++EaVHH8wnTgzP/rHSw2tpatNa4rkssFqNbt25UVFSc8EUXjD14WWqtwwRxrTXt27cnMzOTZDJ5FOoEdYMEgiDhPZAgOTxdd8Gcjhs3jocffphkMolpmjiOw/Dhw9m+fTuWZTFp0iS6d+/OgAEDSCQSXHnllcyYMYNJkyYd82iYTji4joOVhhe+CVr75H2vmMxf/5XEDo9N922k9u2viPXIAa2aU8ig+uMvqV5Zi9Q+OT1zyS3pQIPfgDBM8s7LxYpq3IOHeH/MOjpcWYQZ0bTc2wAQeNqlePTpRLvaKCVQpsZVSTpdfCrZvdtRs2Y/X8z9gmj7OJVv7kYDHb9bSKQwl0N+HYY0QEvSMgAAMJAkDyVPbHzpk/Pcc88xYcKE8A165513Mn/+fIQQ4b5Na0YYTNqcOXMYM2ZMuEh+/etfU1FRwcqVK4/bv+/7FBYWUlZWRm5uLslkkkgkwsKFC2loaMA0zdD/XLlyJdu3b+e0007D8zyGDBnCo48+ytSpU08qxP3PipBu2rSphQ6nTp3K9773vVZPQBxLAl+roqICrTXJZJJ4PM5NN93Ez3/+82Puq/q+T15eHqNGjQoDbg0NDVRVVbUoFxjN5ZdfzvPPPx9Gvk3TZNy4caxevTqcE6UUI0eOZN26dXTv3p1kMsnEiRPZsWMH06dPb3EKQqARpsRtcqnftpvsbqeRUE1IaSCFj69MMtpLvvGby1j3g3WYe5oof2kb4ohlLIVE2j5uvk3Ph3vjWi7aF3jao935ORROOIMdz3yG/mA/B96rgXAzIrhSeTJJFAUXn0K8uCOur/EMH8MVSBtOGXMaX71bRf2ar1n75+VYromfY9DjB6ehaDz22QulUDKC82l524wvQJrPPvuMN954g+HDh+M4DhdddBGjRo2irKwszDQJ9uPSF24wOe+++y6rV6+mf//+OI5Dx44dWbZsGYsWLWLZsmX87W9/o76+Pqwbi8XIz8/noosuYvTo0XTq1Cnc1G1sbOTZZ58NDTlYLIlEgunTpzNz5sxwE/hnP/sZgwcPZtGiRXzyySccPHgwPOmQLsH+ZU1NDdu2bfuHGmCARMuXL6exsZFYLIbv+1x55ZWsW7eOOXPmsGXLlmMeDQrGvnv3br766isgxQLGjh0bRqUff/xxunfvzptvvkl1dXWLgFksFqNHjx5MnjyZkpKSMA9zy5YtVFZWhnoOfLwePXqwYMGCEAFt2+anP/0p8+bNa4FopmlSVVXFsGHDWLVqFZmZmTiOw1NPPcWuXbt49dVXw/IakAgUsOP3Syi+ajAJHZzTS50kUJ5D3mWFXL58IF/MKad2czXUeajmDE4NqKgit1tHim8poV3PHBx1CFNGAEVCOfR+8gLa98xnz5924dYk0KpZn2nJ2wKJ4zdh5Fq4SLR0MbRGGyZJ1UTJ6K40flTFnsW7MNwMvIImLrjnfGJntyfhN7W6DScAVyuyhMn62W8gdu/erbt06QLAhx9+yAUXXNDqogtyFi+99FJWrlwZ5k4mEgn+8pe/0Lt3b+LxeJhyNHr0aMrKykLFBlTn9NNP54MPPiA3NzdMG2urj5FMJrFtGyEE48eP54UXXjgqlzL4f+HChQwbNixE5NZOrh9LFi9ezNChQ1vNYQ2eZ8qUKTz99NOhHr797W/zzjvvHPOMoVKKXr16sXHjxlBHU6dO5ZFHHmlxAsDzPB5//HHuv//+Fqcy2io/+clPmDZtGlJKcnNz2bJlC/n5+S3OI55IfN/H931s26a0tJSXXnopNGCtNXl5eaxZs4Yzzzwz9PNmzZrFxIkTW6WSwXdDhw7ltddew3Gc8HTIgAEDWL9+fai31FIQYBuMXD6TDv0uoSn5Ndg2Eo1AoH0fw5RYxNC4KP9w6jQatAkGBj4Ojg/ycDgULVTqmI80ENho30WjglhoS9HgSY2SKnVcCUV4TFaARYSmPQ0kG5JkFGQQyY2R0AkMJEpojFQ1fAGGAt9xyIx25KtVK3l18K3IIH8yoAjHmxApJatXr2bp0qUhXbBtm0suuaSF4dXV1bF8+fKwHhxGz+3bt3PVVVexc+fO0JCC/M30w6bBp+M44WQGfs/kyZNbNTw4vD85ZswY5s6d2+I8YNBHevvBswdj8Dyv1dSno+alGVWDq60Iebw6QRj+oYce4tVXXw3PHwYBpPSxH3klk8nwHqReNtXV1UycODH0u4+ng0DXwRhs22bevHm8/PLLIY0N1kfAdJqamohGo7z99tth9kprlDbIjnr99de55557sG0b13WJRCIsWLCAgoKCcH1pnVqwJD2WjvkpB7duJSOSjxQa4St85aEEuK5Po1PHIS9JUjg0NV8J6ZD0HRqdQyRcH1D4vsJXPr5KbbB7yqfJcTjk1pMQSZLCJSGcoy/p4GkX5fso5eErja88fKXwPZ8mtxGjs0nmWVnoXDjkNoLSeMpD+T6e8nGVQvgKJQXRaCcat33GW6UPohIeMh6Ph3sy8Xi8TQvopptu4qOPPsK27RbwGtDB8ePHs3///hY/C5C+uNavX8+FF17IAw88wNatWwGwbTs8UBsceDVNM1yA+/btY86cOXzzm99kxowZYXL1kRIskkQiwZgxY7j++utZvnw5tbW1YR/p7acfGg36ysjIOKEOIpEIhmEQi8XCQ78nkqC/oG4kEmlxPwj0eJ7HDTfcwMSJE/n444/DxZs+9iOvSCQS6itd14sXL2bQoEGsWrUqZA6t6SCoaxgGO3fu5L777mPMmDHhmIQQxGIxFi5cyKBBg4AUVd24cSOjRo1q1d1IlwDtfvGLX/Db3/42PAXStWtXVqxYQWFhIdAcM1AapKB2VyWvX/YDtr34CnajR9SMEzWyyTSyyLSyidvtiJvZZMps4s1Xpswm08gmbucQt7LJMLJS5YN6RhaZZjZxO5u4lZMqf7zLSL/S2jFT9WPEiZBBBlnErZzm/lLlY2Y2WUYWMSOOVe+y7YW5/P6yH1C3sxJMELfccovOzs5GCEFlZSVz5sw57gIKKGlOTg633347gwcPJj8/P6Sfv/vd7074Gyvp1CQajXLeeedxxhln0KFDByKRSDiBSimqq6vZs2cPH374YZin2ZYjTUE4OyhXXFzM+eefT0FBAdnZ2cf0mwzD4PPPP+e1115rlX4Hz9WnTx8GDhwYJocvWLCAnTt3tvrcQTv5+fmUlpaGfuvq1at57733jqqTPnbbtjn33HM5++yz6dChQ7jVc6QEKV4rVqzggw8+CNtM1/X5559P79696dChQ6tUv6GhgW3btrF+/foWP9oUfHbu3JkJEyZw8ODB8PT2ggULKC8vb/OPWgX1br31VmKxGK7rkpWVxVtvvcX69etb6CL9747ndKfrZX1o160QbRtHbUX8y4rr8/XuCvYu28D+rTsBEIZE+or/B2aXS8K8xVjXAAAAAElFTkSuQmCC"
            _png = _b64.b64decode(_LOGO_B64)
            _img = _Img.open(_io.BytesIO(_png)).convert("RGBA")
            self._logo_img = _ITk.PhotoImage(_img)
            tk.Label(hdr, image=self._logo_img, bg=SURFACE).pack(side="right", padx=(0, 14))
        except Exception:
            pass

        # Butoni theme (djathtas)
        self._theme_btn = tk.Button(
            hdr,
            text="☀  Light" if _CURRENT_THEME == "dark" else "☾  Dark",
            font=FONT_XS, bg=CARD2, fg=MUTED, relief="flat",
            cursor="hand2", padx=10, pady=3, bd=0,
            activebackground=CARD2, activeforeground=ACCENT,
            command=self._toggle_theme)
        self._theme_btn.pack(side="right", padx=(0, 14))

        # Butoni settings
        tk.Button(hdr, text="⚙", font=("Segoe UI", 12),
                  bg=SURFACE, fg=MUTED, relief="flat", cursor="hand2", bd=0,
                  activebackground=SURFACE, activeforeground=TEXT,
                  command=self.open_settings).pack(side="right", padx=(0, 4))

        # Emri i biznesit (qendër)
        _biz_name = load_config().get("biznes_name", "")
        center = tk.Frame(hdr, bg=SURFACE)
        center.place(relx=0.5, rely=0.5, anchor="center")
        self.lbl_biznes = tk.Label(center, text=_biz_name,
                                   font=("Segoe UI", 11, "bold"), bg=SURFACE, fg=ACCENT)
        self.lbl_biznes.pack()

        # ── Path bar ──────────────────────────────────────────────────────────
        pb = tk.Frame(self, bg=BG, pady=4)
        pb.pack(fill="x")
        tk.Label(pb, text="📁", font=("Segoe UI", 9),
                 bg=BG, fg=MUTED2).pack(side="left", padx=(16, 3))
        self.lbl_path = tk.Label(pb, text=str(get_biz_dir()),
                                  font=("Consolas", 8), bg=BG, fg=MUTED)
        self.lbl_path.pack(side="left")

        # ── Separator ─────────────────────────────────────────────────────────
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── Tabs ──────────────────────────────────────────────────────────────
        style = ttk.Style()
        style.theme_use("default")
        style.configure("T.TNotebook",     background=BG, borderwidth=0, tabmargins=[0,0,0,0])
        style.configure("T.TNotebook.Tab", background=BG, foreground=MUTED,
                        padding=[20, 9],   font=FONT_S,  borderwidth=0,
                        relief="flat")
        style.map("T.TNotebook.Tab",
                  background=[("selected", BG)],
                  foreground=[("selected", ACCENT)])

        nb = ttk.Notebook(self, style="T.TNotebook")
        nb.pack(fill="both", expand=True)

        self.t_form = tk.Frame(nb, bg=BG)
        self.t_biz  = tk.Frame(nb, bg=BG)
        self.t_log  = tk.Frame(nb, bg=BG)

        nb.add(self.t_form, text="  Dërgese  ")
        nb.add(self.t_biz,  text="  Bizneset  ")
        nb.add(self.t_log,  text="  Historia  ")
        nb.bind("<<NotebookTabChanged>>", self._tab_change)

        self._build_form()
        self._build_biz()
        self._build_log()

    def _toggle_theme(self):
        global _CURRENT_THEME
        new = "light" if _CURRENT_THEME == "dark" else "dark"
        cfg = load_config()
        cfg["theme"] = new
        save_config(cfg)
        # Rinicio app me temen e re
        self.destroy()
        set_theme(new)
        app = App()
        app.mainloop()

    def _stat(self, parent, label, val, color):
        pass  # stats bar u hoq

    def _get_biz_list(self):
        _, biz, _ = paths()
        return sorted([f.stem.replace("_", " ") for f in biz.glob("*.xlsx") if SYS_FOLDER not in str(f)])

    def _get_biz_columns(self, biz_name):
        """Lexon kolonat origjinale te file-it xlsx te biznesit."""
        _, biz, _ = paths()
        path = biz / f"{biz_name.strip().replace(' ', '_')}.xlsx"
        if not path.exists():
            return []
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            headers = []
            for cell in next(ws.iter_rows(min_row=1, max_row=1)):
                v = str(cell.value or "").strip()
                if v:
                    headers.append(v)
            wb.close()
            return headers
        except:
            return []

    def _build_form(self):
        self._last_biznesi  = ""
        self._current_cols  = []
        self.ents           = {}
        self._form_frame    = None

        wrapper = tk.Frame(self.t_form, bg=BG)
        wrapper.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(wrapper, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(wrapper, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._win_id, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        outer = tk.Frame(self._inner, bg=BG)
        outer.pack(fill="both", expand=True, padx=20, pady=14)
        self._outer = outer

        # ── Karta e kërkimit të biznesit ─────────────────────────────────────
        search_card = tk.Frame(outer, bg=CARD,
                               highlightbackground=BORDER, highlightthickness=1)
        search_card.pack(fill="x", pady=(0, 8))

        # Padding inner
        sc_inner = tk.Frame(search_card, bg=CARD, padx=16, pady=10)
        sc_inner.pack(fill="x")

        tk.Label(sc_inner, text="BIZNESI", font=FONT_XS,
                 bg=CARD, fg=MUTED2).pack(anchor="w", pady=(0, 4))

        biz_row = tk.Frame(sc_inner, bg=CARD)
        biz_row.pack(fill="x")

        self._biz_var = tk.StringVar()

        # Entry e madhe për emrin e biznesit
        self._biz_entry = tk.Entry(
            biz_row, textvariable=self._biz_var,
            font=("Segoe UI", 14, "bold"),
            bg=INPUT_BG, fg=TEXT,
            insertbackground=ACCENT,
            relief="flat",
            highlightbackground=BORDER, highlightthickness=2)
        self._biz_entry.pack(side="left", fill="x", expand=True, ipady=7)

        tk.Button(biz_row, text="＋", font=("Segoe UI", 13),
                  bg=ACCENT, fg=WHITE, relief="flat", cursor="hand2",
                  padx=10, pady=4, bd=0,
                  activebackground=ACCENT2, activeforeground=WHITE,
                  command=self._add_new_biz).pack(side="right", padx=(8, 0))

        # Info label
        self._lbl_cols = tk.Label(sc_inner,
                                   text="Shkruaj ose zgjidh biznesin...",
                                   font=FONT_XS, bg=CARD, fg=MUTED2)
        self._lbl_cols.pack(anchor="w", pady=(5, 0))

        # ── Suggest Toplevel ──────────────────────────────────────────────────
        self._suggest_win  = None
        self._suggest_list = None

        # ── Fusha dinamike ────────────────────────────────────────────────────
        self._fields_area = tk.Frame(outer, bg=BG)
        self._fields_area.pack(fill="x")

        # Status
        self.lbl_status = tk.Label(outer, text="", font=FONT_XS, bg=BG, fg=GREEN)
        self.lbl_status.pack(pady=(6, 0), anchor="w")

        # Bindings
        self._biz_var.trace_add("write", self._on_biz_type)
        self._biz_entry.bind("<Down>",     self._suggest_down)
        self._biz_entry.bind("<Return>",   self._on_entry_enter)
        self._biz_entry.bind("<Escape>",   lambda e: self._hide_suggestions())
        self._biz_entry.bind("<FocusIn>",  lambda e: self._biz_entry.config(highlightbackground=ACCENT))
        self._biz_entry.bind("<FocusOut>", lambda e: (self._biz_entry.config(highlightbackground=BORDER),
                                                       self.after(200, self._hide_suggestions)))

    def _on_biz_type(self, *args):
        """Kur shkruhet ne entry — shfaq sugjerimet."""
        typed = self._biz_var.get().strip().lower()
        all_biz = self._get_biz_list()

        if not typed:
            self._hide_suggestions()
            return

        matches = [b for b in all_biz if typed in b.lower()]

        if not matches:
            self._hide_suggestions()
            return

        # Match i sakte i vetem — gjenero direkt
        if len(matches) == 1 and matches[0].lower() == typed:
            self._hide_suggestions()
            self._load_biz(matches[0])
            return

        self._show_suggestions(matches[:8])

    def _show_suggestions(self, matches):
        """Hap Toplevel nen entry me listen e sugjerimeve."""
        # Krijo Toplevel nese nuk ekziston
        if self._suggest_win is None or not self._suggest_win.winfo_exists():
            self._suggest_win = tk.Toplevel(self)
            self._suggest_win.overrideredirect(True)   # pa border/title
            self._suggest_win.attributes("-topmost", True)

            outer_f = tk.Frame(self._suggest_win, bg=BORDER)
            outer_f.pack(fill="both", expand=True, padx=1, pady=1)

            self._suggest_list = tk.Listbox(
                outer_f,
                font=("Segoe UI", 12),
                bg=CARD, fg=TEXT,
                selectbackground=ACCENT,
                selectforeground=WHITE if _CURRENT_THEME=="dark" else BG,
                activestyle="none",
                relief="flat", borderwidth=0,
                highlightthickness=0,
                cursor="hand2"
            )
            self._suggest_list.pack(fill="both", expand=True)
            self._suggest_list.bind("<ButtonRelease-1>", self._on_suggest_click)
            self._suggest_list.bind("<Return>",          self._on_suggest_confirm)
            self._suggest_list.bind("<Escape>",          lambda e: self._hide_suggestions())
            self._suggest_list.bind("<Up>",              self._on_suggest_up)
            self._suggest_list.bind("<Down>",            self._on_suggest_nav_down)
            # Klikimi jasht — kontrollo pas pak ms
            self._suggest_win.bind("<FocusOut>",
                lambda e: self.after(100, self._check_focus_out))
            self._suggest_list.bind("<FocusOut>",
                lambda e: self.after(100, self._check_focus_out))

        # Mbush listen
        self._suggest_list.delete(0, "end")
        for m in matches:
            self._suggest_list.insert("end", f"  {m}")

        # Pozicionim sakte nen entry
        self._biz_entry.update_idletasks()
        ex = self._biz_entry.winfo_rootx()
        ey = self._biz_entry.winfo_rooty()
        ew = self._biz_entry.winfo_width()
        eh = self._biz_entry.winfo_height()
        row_h = 32
        h = min(len(matches), 7) * row_h + 4
        self._suggest_win.geometry(f"{ew}x{h}+{ex}+{ey + eh}")
        self._suggest_win.deiconify()

    def _check_focus_out(self):
        """Mbyll listen vetem nese fokusi ka shkuar diku tjeter."""
        try:
            fw = self.focus_get()
            # Mos mbyll nese fokusi eshte ende te entry ose te listbox
            if self._suggest_list and fw == self._suggest_list:
                return
            if fw == self._biz_entry:
                return
            self._hide_suggestions()
        except:
            pass

    def _hide_suggestions(self):
        if self._suggest_win and self._suggest_win.winfo_exists():
            self._suggest_win.withdraw()

    def _suggest_down(self, event):
        """Shigjeta ↓ — shko ne listbox, MOS e fshi tekstin."""
        if self._suggest_win and self._suggest_win.winfo_exists()                 and self._suggest_list and self._suggest_list.size() > 0:
            # Siguro qe lista eshte e dukshme
            self._suggest_win.deiconify()
            self._suggest_win.lift()
            self._suggest_list.focus_set()
            self._suggest_list.selection_clear(0, "end")
            self._suggest_list.selection_set(0)
            self._suggest_list.activate(0)
            self._suggest_list.see(0)
        return "break"

    def _on_suggest_nav_down(self, event):
        """Navigim ↓ brenda listboxes."""
        cur = self._suggest_list.curselection()
        nxt = (cur[0] + 1) if cur else 0
        if nxt < self._suggest_list.size():
            self._suggest_list.selection_clear(0, "end")
            self._suggest_list.selection_set(nxt)
            self._suggest_list.activate(nxt)
        return "break"

    def _on_suggest_up(self, event):
        """Navigim ↑ — nese jemi ne item 0 kthehu te entry."""
        cur = self._suggest_list.curselection()
        if not cur or cur[0] == 0:
            self._biz_entry.focus_set()
        else:
            prv = cur[0] - 1
            self._suggest_list.selection_clear(0, "end")
            self._suggest_list.selection_set(prv)
            self._suggest_list.activate(prv)
        return "break"

    def _get_selected_suggestion(self):
        sel = self._suggest_list.curselection()
        if sel:
            return self._suggest_list.get(sel[0]).strip()
        return None

    def _on_suggest_click(self, event=None):
        name = self._get_selected_suggestion()
        if name:
            self._biz_var.set(name)
            self._hide_suggestions()
            self._load_biz(name)
            self._biz_entry.focus_set()

    def _on_suggest_confirm(self, event=None):
        self._on_suggest_click()
        return "break"

    def _on_entry_enter(self, event=None):
        """Enter ne entry — nese ka match unik ose item i zgjedhur ne liste."""
        # Nese lista eshte hapur dhe ka selektim
        if self._suggest_win and self._suggest_win.winfo_exists():
            name = self._get_selected_suggestion()
            if name:
                self._biz_var.set(name)
                self._hide_suggestions()
                self._load_biz(name)
                return "break"

        typed = self._biz_var.get().strip()
        all_biz = self._get_biz_list()
        matches = [b for b in all_biz if typed.lower() in b.lower()]
        if len(matches) == 1:
            self._biz_var.set(matches[0])
            self._hide_suggestions()
            self._load_biz(matches[0])
        return "break"

    def _load_biz(self, name):
        cols = self._get_biz_columns(name)
        self._regenerate_fields(name, cols)

    def _on_biz_select(self, event=None):
        name = self._biz_var.get().strip()
        if name:
            self._load_biz(name)

    def _refresh_biz_combo(self):
        pass

    def _regenerate_fields(self, biz_name, columns):
        """Gjenero fushat — 1 kolonë, X menjëherë, buton reload."""
        for w in self._fields_area.winfo_children():
            w.destroy()
        self.ents = {}
        self._current_cols = columns
        self._current_biz_name = biz_name

        if not columns:
            tk.Label(self._fields_area,
                     text="Nuk u lexuan kolona nga ky file. Kontrolloje.",
                     font=FONT_S, bg=BG, fg=YELLOW).pack(anchor="w", pady=8)
            self._lbl_cols.config(text="Asnje kolone u lexua.", fg=YELLOW)
            return

        self._lbl_cols.config(
            text=f"{len(columns)} kolona  ·  {biz_name.replace(' ','_')}.xlsx",
            fg=ACCENT)

        # Fusha te fshehura per kete biznes (session)
        hidden_key = f"_hidden_{biz_name}"
        if not hasattr(self, hidden_key):
            setattr(self, hidden_key, set())

        self._draw_fields(biz_name, columns)

    def _draw_fields(self, biz_name, columns):
        """Vizato fushat — Soft UI, fusha kompakte."""
        for w in self._fields_area.winfo_children():
            w.destroy()
        self.ents = {}

        hidden_key = f"_hidden_{biz_name}"
        hidden_set = getattr(self, hidden_key, set())
        today      = datetime.now().strftime("%d/%m/%Y")

        # Karta kryesore
        form = tk.Frame(self._fields_area, bg=CARD,
                        highlightbackground=BORDER, highlightthickness=1)
        form.pack(fill="x", pady=(0, 4))

        # Toolbar — reload button
        toolbar = tk.Frame(form, bg=CARD, padx=14, pady=6)
        toolbar.pack(fill="x")
        hidden_count = len([i for i in range(len(columns)) if i in hidden_set])
        reload_text  = (f"↺  Rivendos ({hidden_count} fshehura)"
                        if hidden_count > 0 else "↺  Rivendos")
        self._reload_btn = tk.Button(
            toolbar, text=reload_text,
            font=FONT_XS, bg=CARD, fg=MUTED2, relief="flat",
            cursor="hand2", bd=0,
            activebackground=CARD, activeforeground=ACCENT,
            command=lambda: self._reload_fields(biz_name, columns))
        self._reload_btn.pack(side="right")

        # ── Fushat ───────────────────────────────────────────────────────────
        fields_wrap = tk.Frame(form, bg=CARD, padx=14, pady=2)
        fields_wrap.pack(fill="x")

        for idx, col_title in enumerate(columns):
            if idx in hidden_set:
                continue

            row = tk.Frame(fields_wrap, bg=CARD, pady=3)
            row.pack(fill="x")

            # Label + X
            lbl_row = tk.Frame(row, bg=CARD)
            lbl_row.pack(fill="x")
            tk.Label(lbl_row, text=col_title, font=FONT_XS,
                     bg=CARD, fg=MUTED, anchor="w").pack(side="left")

            def make_hide(i=idx, r=row, bname=biz_name, cols=columns):
                def hide():
                    getattr(self, f"_hidden_{bname}").add(i)
                    if i in self.ents: del self.ents[i]
                    r.destroy()
                    self._rebind_entries()
                    hc = len(getattr(self, f"_hidden_{bname}", set()))
                    if hasattr(self,"_reload_btn") and self._reload_btn.winfo_exists():
                        self._reload_btn.config(
                            text=f"↺  Rivendos ({hc} fshehura)" if hc > 0 else "↺  Rivendos",
                            fg=ACCENT if hc > 0 else MUTED2)
                return hide

            tk.Button(lbl_row, text="✕", font=FONT_XS, bg=CARD,
                      fg=MUTED2, relief="flat", cursor="hand2", bd=0,
                      activebackground=CARD, activeforeground=RED,
                      command=make_hide()).pack(side="right")

            # Entry kompakte
            e = tk.Entry(row,
                         font=FONT,
                         bg=INPUT_BG, fg=TEXT,
                         insertbackground=ACCENT,
                         relief="flat",
                         highlightbackground=BORDER,
                         highlightthickness=1)
            e.pack(fill="x", ipady=5)
            e.bind("<FocusIn>",  lambda ev, w=e: w.config(highlightbackground=ACCENT))
            e.bind("<FocusOut>", lambda ev, w=e: w.config(highlightbackground=BORDER))

            col_lower = col_title.lower()
            if any(x in col_lower for x in ["data","date","dt","dita"]):
                e.insert(0, today)

            self.ents[idx] = {"entry": e, "col_title": col_title}

        self._rebind_entries()

        # Separator
        tk.Frame(form, bg=BORDER, height=1).pack(fill="x", padx=14, pady=(6,0))

        # Butonat
        btn_row = tk.Frame(form, bg=CARD, padx=14, pady=10)
        btn_row.pack(fill="x")

        self._btn_save = tk.Button(
            btn_row, text="  RUAJ  ",
            font=FONT_B, bg=ACCENT,
            fg=BG if _CURRENT_THEME=="light" else "#0d0f14",
            relief="flat", pady=9, cursor="hand2",
            activebackground=ACCENT2,
            activeforeground=WHITE,
            command=self.save)
        self._btn_save.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._btn_save.bind("<Return>", lambda e: self.save())

        tk.Button(btn_row, text="Pastro",
                  font=FONT_S, bg=CARD2, fg=MUTED,
                  relief="flat", pady=9, cursor="hand2",
                  activebackground=CARD2, activeforeground=TEXT,
                  command=self.clear_form).pack(side="right", ipadx=14)

        self._canvas.yview_moveto(0)

    def _reload_fields(self, biz_name, columns):
        """Rivendos të gjitha fushat — fshin hidden_set."""
        hidden_key = f"_hidden_{biz_name}"
        setattr(self, hidden_key, set())
        self._draw_fields(biz_name, columns)
        # Fokus te fusha e parë
        if self.ents:
            list(self.ents.values())[0]["entry"].focus_set()

    def _rebind_entries(self):
        """Rilidh Enter navigation pas fshehjes se fushave."""
        entries_list = [info["entry"] for info in self.ents.values()
                        if info["entry"].winfo_exists()]
        self._enter_count = 0

        def make_enter_handler(i):
            def handler(event):
                self._enter_count = 0
                if i + 1 < len(entries_list):
                    entries_list[i + 1].focus_set()
                else:
                    self._on_last_field_enter()
                return "break"
            return handler

        for i, ew in enumerate(entries_list[:-1]):
            ew.bind("<Return>", make_enter_handler(i))
        if entries_list:
            entries_list[-1].bind("<Return>", lambda e: (self._on_last_field_enter(), "break")[1])

    def _add_new_biz(self):
        win = tk.Toplevel(self)
        win.title("Biznes i Ri")
        win.geometry("400x160")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="Emri i Biznesit te Ri:", font=FONT_S,
                 bg=BG, fg=MUTED).pack(anchor="w", padx=24, pady=(20, 4))
        e = tk.Entry(win, font=FONT, bg=SURFACE, fg=TEXT,
                     insertbackground=WHITE, relief="flat",
                     highlightbackground=BORDER, highlightthickness=1)
        e.pack(fill="x", padx=24, ipady=7)
        e.focus_set()

        def confirm():
            name = e.get().strip()
            if not name:
                return
            ensure_biznes(name)
            self._biz_var.set(name)
            log(f"Biznes i ri u shtua: {name}")
            win.destroy()
            # Gjenero formularin per biznesin e ri
            cols = self._get_biz_columns(name)
            self._regenerate_fields(name, cols)
            self._biz_entry.focus_set()

        tk.Button(win, text="Shto Biznesin", font=FONT_B,
                  bg=GREEN, fg=WHITE, relief="flat", pady=10,
                  cursor="hand2", command=confirm).pack(fill="x", padx=24, pady=12)
        win.bind("<Return>", lambda ev: confirm())

    def _on_last_field_enter(self):
        """Enter ne fushen e fundit — here e pare highlight, here e dyte ruaj."""
        self._enter_count += 1
        if self._enter_count == 1:
            if hasattr(self, "_btn_save") and self._btn_save.winfo_exists():
                self._btn_save.config(bg=ACCENT2)
                self._btn_save.focus_set()
        else:
            self._enter_count = 0
            if hasattr(self, "_btn_save") and self._btn_save.winfo_exists():
                self._btn_save.config(bg=ACCENT)
            self.save()

    def _refresh_biz_dropdown(self):
        # Rifresko sugjerimet nese nevojitet
        if self._last_biznesi:
            self._biz_var.set(self._last_biznesi)

    def save(self):
        biznesi = self._biz_var.get().strip()
        if not biznesi:
            messagebox.showerror("Gabim", "Zgjidh ose shto nje biznes!")
            return
        if not self.ents:
            messagebox.showerror("Gabim", "Formulari eshte bosh. Zgjidh biznesin nga dropdown!")
            return

        # Nderto entry dict nga kolonat origjinale
        _, biz, _ = paths()
        biz_path = biz / f"{biznesi.strip().replace(' ', '_')}.xlsx"

        # Gjej mapping per keto kolona
        col_mapping, orig_headers = detect_column_mapping(biz_path)

        # Mblidh vlerat nga fushat — ruaj me col_index
        col_values = {}
        for idx, info in self.ents.items():
            val = info["entry"].get().strip()
            col_values[idx + 1] = val  # 1-based

        now = datetime.now()
        dt  = ""
        # Gjej fushen date nga col_values
        for idx, info in self.ents.items():
            if any(x in info["col_title"].lower() for x in ["data","date","dt","dita"]):
                dt = info["entry"].get().strip()
                break
        if not dt:
            dt = now.strftime("%d/%m/%Y")
        ora = now.strftime("%H:%M:%S")

        # Nderto produktin per log
        produkti = ""
        for idx, info in self.ents.items():
            if any(x in info["col_title"].lower() for x in ["produkt","sherbim","item","mall","artikull"]):
                produkti = info["entry"].get().strip()
                break

        # Shto ne file te biznesit — SAKTESISHT sipas kolonave te tij
        try:
            wb = openpyxl.load_workbook(biz_path)
            ws = wb.active
            nr = ws.max_row + 1
            H_BG_loc = H_BG
            for idx, info in self.ents.items():
                col_num = idx + 1
                val = info["entry"].get().strip()
                # Konverto numra
                try:
                    if val and not any(c.isalpha() for c in val):
                        val = float(val.replace(",", "."))
                except:
                    pass
                cell = ws.cell(row=nr, column=col_num, value=val)
                cell.alignment = Alignment(vertical="center")
                if nr % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=ROW_ALT)
                cell.border = Border(bottom=Side(style="hair", color="CCCCCC"))
            # Shto Data/Ora nëse përdoruesi ka zgjedhur "Po"
            if load_config().get("add_timestamp", False):
                ts_label = "Data/Ora e shtimit"
                ts_col = None
                for c in range(1, ws.max_column + 1):
                    if str(ws.cell(row=1, column=c).value or "").strip() == ts_label:
                        ts_col = c
                        break
                if ts_col is None:
                    ts_col = ws.max_column + 1
                    hcell = ws.cell(row=1, column=ts_col, value=ts_label)
                    hcell.font = Font(bold=True, color="FFFFFF", size=11)
                    hcell.fill = PatternFill("solid", fgColor="1e3a5f")
                    hcell.alignment = Alignment(horizontal="center", vertical="center")
                ts_val = f"{dt}  {ora}".strip()
                cell = ws.cell(row=nr, column=ts_col, value=ts_val)
                cell.alignment = Alignment(vertical="center")
                if nr % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=ROW_ALT)
            auto_width(ws)
            wb.save(biz_path)
        except Exception as ex:
            messagebox.showerror("Gabim ne ruajtje", str(ex))
            return

        log(f"Dergese u ruajt: {biznesi}  |  {produkti}  |  {biz_path.name}")
        self._last_biznesi = biznesi
        self.refresh_stats()
        self._refresh_biz_dropdown()
        self.clear_form()
        self.lbl_status.config(
            text=f"U ruajt: {biznesi}  |  {produkti}  |  {biz_path.name}")

    def clear_form(self):
        """Pastro vlerat e fushatve, mbaj biznesin e fundit"""
        today = datetime.now().strftime("%d/%m/%Y")
        for idx, info in self.ents.items():
            info["entry"].delete(0, "end")
            col_lower = info["col_title"].lower()
            if any(x in col_lower for x in ["data", "date", "dt", "dita"]):
                info["entry"].insert(0, today)
        if self._last_biznesi:
            self._biz_var.set(self._last_biznesi)
        self._enter_count = 0
        if hasattr(self, "_btn_save") and self._btn_save.winfo_exists():
            self._btn_save.config(bg=ACCENT)
        if self.ents:
            list(self.ents.values())[0]["entry"].focus_set()

    # ── BIZNESET ──────────────────────────────────────────────────────────────
    def _build_biz(self):
        top = tk.Frame(self.t_biz, bg=BG, pady=10)
        top.pack(fill="x", padx=16)
        tk.Button(top, text="⟳  Rifresko", font=FONT_XS, bg=CARD2, fg=MUTED,
                  relief="flat", cursor="hand2", padx=8, pady=4, bd=0,
                  activebackground=CARD2, activeforeground=TEXT,
                  command=self.refresh_biz).pack(side="right")
        tk.Button(top, text="📁  Hap Dosjen", font=FONT_XS, bg=CARD2, fg=ACCENT,
                  relief="flat", cursor="hand2", padx=8, pady=4, bd=0,
                  activebackground=CARD2, activeforeground=ACCENT,
                  command=self._open_biz_folder).pack(side="right", padx=(0,6))

        style = ttk.Style()
        style.configure("B.Treeview", background=CARD, foreground=TEXT,
                        fieldbackground=CARD, font=FONT, rowheight=28, borderwidth=0)
        style.configure("B.Treeview.Heading", background=SURFACE,
                        foreground=MUTED, font=FONT_XS, relief="flat")
        style.map("B.Treeview", background=[("selected", ACCENT)],
                  foreground=[("selected", WHITE)])

        fr = tk.Frame(self.t_biz, bg=BG)
        fr.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        cols = ("biznesi", "dergesa", "pagesa", "file")
        self.tree = ttk.Treeview(fr, columns=cols, show="headings", style="B.Treeview")
        self.tree.heading("biznesi", text="  Biznesi")
        self.tree.heading("dergesa", text="Dergesa")
        self.tree.heading("pagesa",  text="Totali (EUR)")
        self.tree.heading("file",    text="File Excel")
        self.tree.column("biznesi", width=240, anchor="w")
        self.tree.column("dergesa", width=80,  anchor="center")
        self.tree.column("pagesa",  width=130, anchor="e")
        self.tree.column("file",    width=260, anchor="w")

        sb = ttk.Scrollbar(fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._open_biz_file)
        self.refresh_biz()

    def refresh_biz(self):
        _, biz, _ = paths()
        for r in self.tree.get_children():
            self.tree.delete(r)
        for f in sorted(f for f in biz.glob("*.xlsx") if SYS_FOLDER not in str(f)):
            name = f.stem.replace("_", " ")
            try:
                wb   = openpyxl.load_workbook(f, data_only=True, read_only=True)
                ws   = wb.active
                rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
                cnt  = len(rows)
                # pagesa eshte col index 6 (0-based) ne file biznesit
                tot  = sum(float(r[6] or 0) for r in rows if len(r) > 6 and r[6])
                wb.close()
            except:
                cnt, tot = 0, 0.0
            self.tree.insert("", "end",
                values=(f"  {name}", cnt, f"EUR {tot:,.2f}", f.name))

    def _open_biz_file(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        _, biz, _ = paths()
        fname = self.tree.item(sel[0])["values"][3]
        path  = biz / fname
        if path.exists():
            if sys.platform == "win32":
                os.startfile(str(path))
            else:
                os.system(f'open "{path}"')

    def _open_biz_folder(self):
        _, biz, _ = paths()
        if sys.platform == "win32":
            os.startfile(str(biz))
        else:
            os.system(f'open "{biz}"')

    # ── HISTORIA ──────────────────────────────────────────────────────────────
    def _build_log(self):
        top = tk.Frame(self.t_log, bg=BG)
        top.pack(fill="x", padx=24, pady=(18, 8))
        tk.Label(top, text="HISTORIA E VEPRIMEVE", font=("Segoe UI", 8),
                 bg=BG, fg=MUTED).pack(side="left")
        tk.Button(top, text="Rifresko", font=FONT_S, bg=SURFACE, fg=MUTED,
                  relief="flat", cursor="hand2",
                  command=self.refresh_log).pack(side="right")

        fr = tk.Frame(self.t_log, bg=BG)
        fr.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        self.log_txt = tk.Text(fr, font=("Consolas", 9), bg=CARD, fg=TEXT,
                                relief="flat", state="disabled", wrap="word",
                                highlightbackground=BORDER, highlightthickness=1)
        sb = ttk.Scrollbar(fr, orient="vertical", command=self.log_txt.yview)
        self.log_txt.configure(yscrollcommand=sb.set)
        self.log_txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.refresh_log()

    def refresh_log(self):
        _, _, lg = paths()
        self.log_txt.config(state="normal")
        self.log_txt.delete("1.0", "end")
        if lg.exists():
            lines = lg.read_text(encoding="utf-8").strip().split("\n")
            self.log_txt.insert("1.0", "\n".join(reversed(lines)))
        else:
            self.log_txt.insert("1.0", "Ende nuk ka veprime.")
        self.log_txt.config(state="disabled")

    # ── CILESIMET ─────────────────────────────────────────────────────────────
    def open_settings(self):
        win = tk.Toplevel(self)
        win.title("Nderro Folderin")
        win.geometry("520x240")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="NDERRO FOLDERIN", font=FONT_H,
                 bg=BG, fg=WHITE).pack(pady=(20, 12))

        frame = tk.Frame(win, bg=CARD, padx=20, pady=16,
                         highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill="x", padx=24)

        tk.Label(frame, text="Folder ku ruhen te dhenat:",
                 font=FONT_S, bg=CARD, fg=MUTED).pack(anchor="w")

        row = tk.Frame(frame, bg=CARD)
        row.pack(fill="x", pady=(6, 0))

        path_var = tk.StringVar(value=str(get_base_dir()))
        ent = tk.Entry(row, textvariable=path_var, font=FONT_S,
                       bg=SURFACE, fg=TEXT, insertbackground=WHITE,
                       relief="flat", highlightbackground=BORDER, highlightthickness=1)
        ent.pack(side="left", fill="x", expand=True, ipady=6)

        def browse():
            p = filedialog.askdirectory(title="Zgjidh Folder")
            if p:
                path_var.set(p)

        tk.Button(row, text="Shfleto", font=FONT_S, bg=ACCENT, fg=WHITE,
                  relief="flat", cursor="hand2", padx=10,
                  command=browse).pack(side="right", padx=(6, 0), ipady=6)

        def apply():
            p = Path(path_var.get().strip())
            try:
                p.mkdir(parents=True, exist_ok=True)
                # Detekto nëse klienti zgjodhi direkt folder-in me xlsx të bizneseve
                direct_xlsx = [f for f in p.glob("*.xlsx") if SYS_FOLDER not in str(f)]
                has_biz_sub = (p / "Bizneset").exists()
                if direct_xlsx or (not has_biz_sub and p.parent != p):
                    # Ky folder ËSHTË folder i bizneseve — përdore direkt
                    base_p = p.parent
                    base_p.mkdir(parents=True, exist_ok=True)
                    set_dirs(base_p, biz_path=p)
                else:
                    set_dirs(p)
                    (p / "Bizneset").mkdir(exist_ok=True)
                self.lbl_path.config(text=str(get_biz_dir()))
                self.refresh_stats()
                self.refresh_biz()
                log(f"Folder u ndryshua ne: {p}")
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Gabim", str(ex))

        btn_row2 = tk.Frame(win, bg=BG)
        btn_row2.pack(fill="x", padx=24, pady=(12,4))
        tk.Button(btn_row2, text="RUAJ", font=FONT_B,
                  bg=GREEN, fg=WHITE, relief="flat", pady=10,
                  cursor="hand2", command=apply).pack(side="left", fill="x", expand=True, padx=(0,6))


        # ── Butoni Reset Setup ───────────────────────────────────────────
        sep = tk.Frame(win, bg=BORDER, height=1)
        sep.pack(fill="x", padx=24, pady=(10, 0))

        def reset_setup():
            # Dritarja e fjalëkalimit
            pw_win = tk.Toplevel(win)
            pw_win.title("Autorizim")
            pw_win.geometry("340x180")
            pw_win.configure(bg=BG)
            pw_win.resizable(False, False)
            pw_win.grab_set()
            pw_win.update_idletasks()
            x = (pw_win.winfo_screenwidth()  - 340) // 2
            y = (pw_win.winfo_screenheight() - 180) // 2
            pw_win.geometry(f"340x180+{x}+{y}")

            tk.Label(pw_win, text="🔒  Kërkohet fjalëkalimi i adminit",
                     font=FONT_B, bg=BG, fg=WHITE).pack(pady=(22, 6))

            pw_var = tk.StringVar()
            err_var = tk.StringVar()

            pw_entry = tk.Entry(pw_win, textvariable=pw_var, show="●",
                                font=("Segoe UI", 11), bg=SURFACE, fg=WHITE,
                                insertbackground=WHITE, relief="flat",
                                highlightbackground=BORDER, highlightthickness=1,
                                justify="center")
            pw_entry.pack(fill="x", padx=30, ipady=8)
            pw_entry.focus_set()

            tk.Label(pw_win, textvariable=err_var,
                     font=("Segoe UI", 8), bg=BG, fg=RED).pack(pady=(4, 0))

            def _do_reset(event=None):
                if pw_var.get() != "admin":
                    err_var.set("✖  Fjalëkalim i gabuar")
                    pw_entry.delete(0, "end")
                    return
                pw_win.destroy()
                try:
                    cfg = load_config()
                    cfg.pop("base_dir", None)
                    cfg.pop("biz_dir", None)
                    cfg.pop("biznes_name", None)
                    cfg.pop("add_timestamp", None)
                    save_config(cfg)
                    win.destroy()
                    self.destroy()
                    run_setup_if_needed()
                    App().mainloop()
                except Exception as ex:
                    messagebox.showerror("Gabim", str(ex))

            pw_entry.bind("<Return>", _do_reset)
            tk.Button(pw_win, text="VAZHDO", font=FONT_B,
                      bg=RED, fg=WHITE, relief="flat", pady=8,
                      cursor="hand2", command=_do_reset).pack(fill="x", padx=30, pady=(8, 0))

        tk.Button(win, text="⚙  Rikonfiguro — Fillo Setup Sërisht",
                  font=FONT_S, bg="#1a1d26", fg="#ef4444",
                  relief="flat", pady=8, cursor="hand2",
                  command=reset_setup).pack(fill="x", padx=24, pady=(6, 14))

        win.geometry("520x320")

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def refresh_stats(self):
        pass  # stats bar u hoq

    def _tab_change(self, event):
        tab = event.widget.tab("current", "text").strip()
        if "Bizneset" in tab:
            self.refresh_biz()
        elif "Historia" in tab:
            self.refresh_log()
        elif "Dergese" in tab:
            self._refresh_biz_dropdown()
        self.refresh_stats()


def run_biznes_name_if_needed():
    """Pyet per emrin e biznesit nese nuk eshte vendosur akoma."""
    cfg = load_config()
    if cfg.get("biznes_name", "").strip():
        return  # tasme konfiguruar

    root = tk.Tk()
    root.title("NjoPerKrejt - SmartRegister")
    root.geometry("420x220")
    root.configure(bg=BG)
    root.resizable(False, False)
    root.update_idletasks()
    x = (root.winfo_screenwidth()  - 420) // 2
    y = (root.winfo_screenheight() - 220) // 2
    root.geometry(f"420x220+{x}+{y}")

    tk.Label(root, text="Emri i Biznesit Tuaj",
             font=("Segoe UI", 13, "bold"), bg=BG, fg=WHITE).pack(pady=(28, 4))
    tk.Label(root, text="Ky emer do te shfaqet ne krye te programit.",
             font=("Segoe UI", 9), bg=BG, fg=MUTED).pack()

    frame = tk.Frame(root, bg=CARD, padx=20, pady=14,
                     highlightbackground=BORDER, highlightthickness=1)
    frame.pack(fill="x", padx=24, pady=14)

    name_var = tk.StringVar()
    ent = tk.Entry(frame, textvariable=name_var, font=("Segoe UI", 12),
                   bg=SURFACE, fg=WHITE, insertbackground=WHITE,
                   relief="flat", highlightbackground=BORDER, highlightthickness=1,
                   justify="center")
    ent.pack(fill="x", ipady=9)
    ent.focus_set()

    def confirm(event=None):
        name = name_var.get().strip()
        if not name:
            ent.config(highlightbackground=RED)
            return
        cfg = load_config()
        cfg["biznes_name"] = name
        save_config(cfg)
        root.destroy()

    ent.bind("<Return>", confirm)
    tk.Button(root, text="VAZHDO  ▶", font=("Segoe UI", 10, "bold"),
              bg=GREEN, fg=WHITE, relief="flat", pady=10,
              cursor="hand2", command=confirm).pack(fill="x", padx=24)

    root.mainloop()


if __name__ == "__main__":
    _boot_external()
    run_setup_if_needed()
    run_biznes_name_if_needed()
    check_for_update()
    App().mainloop()
