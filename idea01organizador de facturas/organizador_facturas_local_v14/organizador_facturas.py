import json, os, re, shutil, threading, unicodedata
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document

CONFIG_FILE = "config.json"
DEFAULT_CONFIG = {
    "watch_folder": "",
    "excel_file": "",
    "community_columns": "B",
    "id_columns": "A",
    "route_column": "",
    "alias_id_column": "R",
    "alias_name_column": "S",
    "use_general_route": True,
    "general_route": "",
    "general_route_pattern": r"{base}\\PH{ID4}\\{ID}-Año {year}\\{ID}-FACTURAS {year}",
    "concepts_excel": "",
    "concept_column": "A",
    "concept_words_column": "B",
    "sheet": "",
    "filename_template": "{ID}-Fact.{concepto}.{mes}.{año}.pdf",
    "ocr": True,
    "move": False,
    "overwrite": False,
    "visitas_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Detalles de la Visita",
    "partes_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\PARTES DE TRABAJO",
    "otro_formato_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Otro formato",
    "comunidad_no_encontrada_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Comunidad no encontrada",
    "id_no_encontrado_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Id no encontrado",
    "concepto_no_encontrado_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Concepto no encontrado",
    "ruta_no_encontrada_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Ruta no encontrada",
    "movimientos_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Movimientos",
    "aux_year": str(datetime.now().year),
}

try:
    import pytesseract
    import pypdfium2 as pdfium
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False


def normalize_text(value):
    if value is None:
        return ""
    s = str(value)
    s = re.sub(r'\bN\s*[º°Oo]?\s*(?=\d)', ' ', s, flags=re.I)
    s = unicodedata.normalize("NFKD", s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.upper()
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def normalize_community_address(text):
    if text is None:
        return ""
    s = str(text).upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("º", " ").replace("°", " ").replace("ª", " ")
    s = re.sub(r"\bN\s*[.º°]?\s*O\b", " ", s)
    s = re.sub(r"\bN[.º°]?\b", " ", s)
    s = s.replace("/", " ").replace("\\", " ").replace(",", " ").replace("-", " ")
    s = re.sub(r"\bC\s*/?\b", " ", s)
    s = re.sub(r"\bCALLE\b", " ", s)
    s = re.sub(r"\bCL\b", " ", s)
    s = re.sub(r"\bC\.\b", " ", s)
    s = re.sub(r"\bCOMUNIDAD\s+DE\s+PROPIETARIOS\b", " ", s)
    s = re.sub(r"\bCOMUNIDAD\s+PROPIETARIOS\b", " ", s)
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def split_columns(value):
    if value is None:
        return []
    return [x.strip().upper() for x in re.split(r"[,;\s]+", str(value)) if x.strip()]


def col_index(col):
    n = 0
    for c in str(col).strip().upper():
        if 'A' <= c <= 'Z':
            n = n * 26 + ord(c) - 64
    return n - 1


def read_cell(row, col):
    i = col_index(col)
    return row[i].value if 0 <= i < len(row) else None


def format_community_id(value):
    if value is None:
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+(?:\.0+)?", s):
        try:
            return f"{int(float(s)):03d}"
        except Exception:
            pass
    return s


def read_communities(path, sheet, community_cols, id_cols, route_col, alias_id_col, alias_name_col):
    wb = load_workbook(path, data_only=True, read_only=True)
    ccols, icols = split_columns(community_cols), split_columns(id_cols)
    if len(ccols) != len(icols):
        raise ValueError("Las columnas de comunidad e ID deben tener el mismo número de elementos y estar emparejadas por posición (A,C,E / B,D,F).")

    # IMPORTANTE: algunos Excels tienen una primera hoja vacía (por ejemplo
    # 'Hoja1') y los datos reales están en otra hoja. Si el usuario no indica
    # una hoja, elegimos automáticamente la que tenga más datos en las
    # columnas comunidad/ID configuradas.
    if sheet and sheet in wb.sheetnames:
        ws = wb[sheet]
    else:
        best_ws = None
        best_count = -1
        for candidate in wb.worksheets:
            count = 0
            for row in candidate.iter_rows():
                for ccol, icol in zip(ccols, icols):
                    c = read_cell(row, ccol)
                    cid = read_cell(row, icol)
                    if c is not None and str(c).strip() and cid is not None and str(cid).strip():
                        count += 1
            if count > best_count:
                best_count = count
                best_ws = candidate
        ws = best_ws if best_ws is not None else wb[wb.sheetnames[0]]

    # Si la columna de ruta coincide con una columna usada para ID o comunidad,
    # no puede tratarse como ruta. Esto evita que C (que puede ser otro ID)
    # se interprete accidentalmente como una ruta.
    effective_route_col = route_col.strip().upper() if route_col else ""
    used_cols = set(ccols + icols)
    if effective_route_col in used_cols:
        effective_route_col = ""

    rows = []
    by_id = {}
    for row in ws.iter_rows():
        for ccol, icol in zip(ccols, icols):
            c = read_cell(row, ccol)
            cid = read_cell(row, icol)
            if c is None or not str(c).strip():
                continue
            item = {
                "community": str(c).strip(),
                "id": cid,
                "id_fmt": format_community_id(cid),
                "route": str(read_cell(row, effective_route_col) or "").strip() if effective_route_col else "",
                "source": "principal",
            }
            rows.append(item)
            if item["id_fmt"]:
                by_id.setdefault(item["id_fmt"], item)

    # R = ID y S = nombre alternativo. Se asocia el alias al registro principal por ID.
    alias_count = 0
    if alias_id_col and alias_name_col:
        for row in ws.iter_rows():
            aid = read_cell(row, alias_id_col)
            aname = read_cell(row, alias_name_col)
            if aid is None or aname is None or not str(aname).strip():
                continue
            aid_fmt = format_community_id(aid)
            if not aid_fmt:
                continue
            base = by_id.get(aid_fmt)
            if base:
                item = dict(base)
                item["community"] = str(aname).strip()
                item["source"] = "alias"
            else:
                item = {"community": str(aname).strip(), "id": aid, "id_fmt": aid_fmt, "route": "", "source": "alias"}
            rows.append(item)
            alias_count += 1
    return rows, alias_count


def split_terms(value):
    if value is None:
        return []
    return [x.strip() for x in re.split(r'[,;\n]+', str(value)) if x.strip()]


def read_concepts(path, sheet, concept_col, first_word_col):
    if not path:
        return []
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    ci, wi = col_index(concept_col), col_index(first_word_col)
    concepts = []
    for row in ws.iter_rows():
        if ci >= len(row):
            continue
        canonical = row[ci].value
        if canonical is None or not str(canonical).strip():
            continue
        terms = []
        for j in range(wi, len(row)):
            for raw in split_terms(row[j].value):
                n = normalize_text(raw)
                if n and n not in terms:
                    terms.append(n)
        if terms:
            concepts.append({"concept": str(canonical).strip(), "terms": terms})
    return concepts


def extract_text(pdf_path, use_ocr):
    parts = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            try:
                parts.append(page.extract_text() or "")
            except Exception:
                pass
    except Exception:
        pass
    normal = "\n".join(parts)
    if use_ocr and OCR_AVAILABLE and len(normalize_text(normal)) < 100:
        try:
            doc = pdfium.PdfDocument(pdf_path)
            ocr_parts = []
            for i in range(len(doc)):
                page = doc[i]
                bitmap = page.render(scale=3)
                img = bitmap.to_pil_image()
                try:
                    txt = pytesseract.image_to_string(img, lang="spa+eng")
                except Exception:
                    txt = pytesseract.image_to_string(img)
                ocr_parts.append(txt)
            return "\n".join(ocr_parts), "OCR"
        except Exception as e:
            return normal, "texto (OCR falló: %s)" % e
    return normal, "texto PDF"


def contains_term(normalized_text, term):
    return re.search(r'(?<![A-Z0-9])' + re.escape(term) + r'(?![A-Z0-9])', normalized_text) is not None


def detect_concept(text, concepts, logger=None):
    nt = normalize_text(text)
    candidates = []
    for item in concepts:
        for term in item["terms"]:
            candidates.append((len(term), item["concept"], term))
    candidates.sort(key=lambda x: x[0], reverse=True)
    for _, concept, term in candidates:
        if contains_term(nt, term):
            return concept, term
    for item in concepts:
        for term in item["terms"]:
            for token in term.split():
                if len(token) >= 4 and contains_term(nt, token):
                    return item["concept"], token
    return None, None


def community_match_score(community, invoice_text):
    target = normalize_community_address(community)
    text = normalize_community_address(invoice_text)
    if not target or not text:
        return 0
    if re.search(r"(?<!\w)" + re.escape(target) + r"(?!\w)", text):
        return 100 + len(target)
    tokens = [t for t in target.split() if len(t) >= 2]
    if tokens and all(re.search(r"(?<!\w)" + re.escape(t) + r"(?!\w)", text) for t in tokens):
        return 80 + len(tokens)
    return 0


def detect_community(text, communities):
    scored = []
    for row in communities:
        score = community_match_score(row.get("community", ""), text)
        if score > 0:
            scored.append((score, row))
    if not scored:
        return None, "not_found", []
    max_score = max(x[0] for x in scored)
    best = [row for score, row in scored if score == max_score]
    # Varias coincidencias solo son ambiguas si corresponden a IDs distintos.
    ids = {r.get("id_fmt") or str(r.get("id")) for r in best}
    if len(ids) > 1:
        return None, "multiple", best
    return best[0], "ok", best


def extract_date(text, fallback_path=None):
    patterns = [
        r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b',
        r'\b(\d{1,2})\s+(\d{1,2})\s+(\d{2,4})\b',
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            month, year = int(m.group(2)), int(m.group(3))
            if 1 <= month <= 12:
                return f"{month:02d}/{year % 100:02d}", f"{month:02d}", f"{year % 100:02d}", str(year if year >= 100 else 2000 + year)
    for m in re.finditer(r'\b(\d{1,2})[/-](\d{2,4})\b', text):
        month, year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return f"{month:02d}/{year % 100:02d}", f"{month:02d}", f"{year % 100:02d}", str(year if year >= 100 else 2000 + year)
    # Si la factura no contiene ninguna fecha, usamos la fecha en la que
    # el archivo fue guardado/creado en Windows.
    if fallback_path:
        try:
            ts = os.path.getctime(fallback_path)
        except Exception:
            try:
                ts = os.path.getmtime(fallback_path)
            except Exception:
                ts = time.time()
        dt = datetime.fromtimestamp(ts)
        return dt.strftime("%m/%y"), dt.strftime("%m"), dt.strftime("%y"), dt.strftime("%Y")
    return "", "", "", ""


def clean_filename(s):
    s = "" if s is None else str(s)
    return re.sub(r'[<>:"/\\|?*]', '_', s).strip().rstrip('.')


def make_filename(template, community, concept, date_info):
    fecha, mes, ano, _ = date_info
    cid = community.get("id_fmt") or format_community_id(community.get("id"))
    vals = {
        "ID": clean_filename(cid), "id": clean_filename(cid),
        "comunidad": clean_filename(community["community"]),
        "concepto": clean_filename(concept or "SIN_CONCEPTO"),
        "mes": mes or "SIN_MES", "año": ano or "SIN_AÑO", "anio": ano or "SIN_ANIO",
        "fecha": fecha or "SIN_FECHA", "numero": "", "original": ""
    }
    try:
        name = template.format(**vals)
    except Exception:
        name = f"{vals['ID']}-Fact.{vals['concepto']}.{vals['mes']}.{vals['año']}.pdf"
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return clean_filename(name)


def unique_path(path):
    if not os.path.exists(path):
        return path
    folder, base = os.path.dirname(path), os.path.basename(path)
    stem, ext = os.path.splitext(base)
    n = 1
    while True:
        candidate = os.path.join(folder, f"{stem}.{n}{ext}")
        if not os.path.exists(candidate):
            return candidate
        n += 1


def build_general_route(base_route, community_id, year, pattern=None):
    base = os.path.expandvars(os.path.expanduser(str(base_route or "").strip()))
    cid = format_community_id(community_id)
    try:
        id4 = f"{int(cid):04d}" if cid.isdigit() else cid.zfill(4)
    except Exception:
        id4 = cid
    pattern = pattern or r"{base}\\PH{ID4}\\{ID}-Año {year}\\{ID}-FACTURAS {year}"
    vals = {"base": base.rstrip("\\/"), "ID": cid, "id": cid, "ID4": id4, "id4": id4, "year": str(year), "año": str(year), "anio": str(year)}
    try:
        return pattern.format(**vals)
    except Exception:
        return os.path.join(base, f"PH{id4}", f"{cid}-Año {year}", f"{cid}-FACTURAS {year}")


def resolve_route(row, year, cfg):
    individual = str(row.get("route", "") or "").strip()
    if individual:
        return os.path.expandvars(os.path.expanduser(individual)), "Excel"
    base = str(cfg.get("general_route", "") or "").strip()
    if base:
        return build_general_route(base, row.get("id_fmt") or row.get("id"), year, cfg.get("general_route_pattern")), "General"
    return "", "Sin ruta"


def file_year(cfg, date_info):
    return date_info[3] or str(cfg.get("aux_year") or datetime.now().year)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de documentos v13")
        self.root.geometry("1080x900")
        self.cfg = self.load_cfg()
        self.v = {}
        self.running = False
        self.last_undo_file = os.path.join(os.path.dirname(os.path.abspath(CONFIG_FILE)), "ultimo_proceso_undo.json")
        self.build(); self.load_ui()

    def load_cfg(self):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                d = DEFAULT_CONFIG.copy(); d.update(json.load(f)); return d
        except Exception:
            return DEFAULT_CONFIG.copy()

    def build(self):
        main = ttk.Frame(self.root, padding=10); main.pack(fill="both", expand=True)
        ttk.Label(main, text="Organizador de documentos v13", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        self.entry(main, "Carpeta de observación", "watch_folder", self.folder)
        self.entry(main, "Excel comunidades", "excel_file", self.file_community)
        self.entry(main, "Columnas comunidad", "community_columns")
        ttk.Label(main, text="Ej.: B,D,F  (se emparejan con A,C,E)").pack(anchor="w", padx=4)
        self.entry(main, "Columnas ID", "id_columns")
        ttk.Label(main, text="Ej.: A,C,E  (A→B, C→D, E→F)").pack(anchor="w", padx=4)
        self.entry(main, "Columna ruta (opcional)", "route_column")
        self.entry(main, "ID alternativos (R)", "alias_id_column")
        self.entry(main, "Nombres alternativos (S)", "alias_name_column")
        ttk.Label(main, text="R y S permiten añadir nombres alternativos: R=ID, S=nombre alternativo.").pack(anchor="w", padx=4)
        self.entry(main, "Ruta general (base)", "general_route", self.folder)
        ttk.Label(main, text=r"Ej.: ...\PH0002\002-Año 2026\002-FACTURAS 2026").pack(anchor="w", padx=4)
        self.entry(main, "Excel conceptos", "concepts_excel", self.file_concepts)
        self.entry(main, "Columna concepto", "concept_column")
        self.entry(main, "Primera columna palabras", "concept_words_column")
        self.entry(main, "Hoja (vacío=primera)", "sheet")
        self.entry(main, "Plantilla nombre", "filename_template")
        self.entry(main, "Año carpetas auxiliares", "aux_year")
        ttk.Separator(main).pack(fill="x", pady=8)
        ttk.Label(main, text="Carpetas auxiliares", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        for label, key in [
            ("Detalles de la Visita", "visitas_folder"), ("PARTES DE TRABAJO", "partes_folder"),
            ("Otro formato", "otro_formato_folder"), ("Comunidad no encontrada", "comunidad_no_encontrada_folder"),
            ("Id no encontrado", "id_no_encontrado_folder"), ("Concepto no encontrado", "concepto_no_encontrado_folder"),
            ("Ruta no encontrada", "ruta_no_encontrada_folder"),
            ("Movimientos", "movimientos_folder")]:
            self.entry(main, label, key, self.folder_for(key))
        self.ocr = tk.BooleanVar(value=True); self.move = tk.BooleanVar(value=False); self.overwrite = tk.BooleanVar(value=False)
        ttk.Checkbutton(main, text="Usar OCR", variable=self.ocr).pack(anchor="w")
        ttk.Checkbutton(main, text="Mover en vez de copiar facturas", variable=self.move).pack(anchor="w")
        ttk.Checkbutton(main, text="Sobrescribir", variable=self.overwrite).pack(anchor="w")
        b = ttk.Frame(main); b.pack(fill="x", pady=8)
        self.save_btn = ttk.Button(b, text="Guardar configuración", command=self.save); self.save_btn.pack(side="left", padx=4)
        self.run_btn = ttk.Button(b, text="Organizar", command=self.start); self.run_btn.pack(side="left", padx=4)
        self.undo_btn = ttk.Button(b, text="↩ Revertir cambios", command=self.undo); self.undo_btn.pack(side="left", padx=4)
        self.pb = ttk.Progressbar(main, mode="determinate"); self.pb.pack(fill="x", pady=6)
        ttk.Label(main, text="Registro detallado").pack(anchor="w")
        self.log = tk.Text(main, height=16, wrap="word"); self.log.pack(fill="both", expand=True)

    def entry(self, parent, label, key, browse=None):
        f = ttk.Frame(parent); f.pack(fill="x", pady=2)
        ttk.Label(f, text=label, width=29).pack(side="left")
        var = tk.StringVar(); self.v[key] = var
        ttk.Entry(f, textvariable=var).pack(side="left", fill="x", expand=True)
        if browse: ttk.Button(f, text="Examinar", command=browse).pack(side="left", padx=4)

    def folder_for(self, key):
        def choose():
            p = filedialog.askdirectory()
            if p: self.v[key].set(p)
        return choose

    def load_ui(self):
        for k, var in self.v.items(): var.set(self.cfg.get(k, ""))
        self.ocr.set(self.cfg.get("ocr", True)); self.move.set(self.cfg.get("move", False)); self.overwrite.set(self.cfg.get("overwrite", False))

    def save(self):
        for k, var in self.v.items(): self.cfg[k] = var.get().strip()
        self.cfg["ocr"] = self.ocr.get(); self.cfg["move"] = self.move.get(); self.cfg["overwrite"] = self.overwrite.get()
        with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(self.cfg, f, ensure_ascii=False, indent=2)
        self.log_msg("✓ Configuración guardada.")

    def folder(self):
        p = filedialog.askdirectory()
        if p: self.v["watch_folder"].set(p)

    def file_community(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p: self.v["excel_file"].set(p)

    def file_concepts(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p: self.v["concepts_excel"].set(p)

    def log_msg(self, s):
        self.root.after(0, lambda: (self.log.insert("end", s + "\n"), self.log.see("end")))

    def set_progress(self, value, maximum):
        self.root.after(0, lambda: (self.pb.configure(maximum=maximum, value=value)))

    def set_running(self, value):
        self.running = value
        def ui():
            self.run_btn.configure(state="disabled" if value else "normal")
            self.save_btn.configure(state="disabled" if value else "normal")
        self.root.after(0, ui)

    def start(self):
        if self.running: return
        self.save(); self.set_running(True)
        threading.Thread(target=self.run, daemon=True).start()

    def register_op(self, undo_ops, original, action, destination, new_name):
        undo_ops.append({"original": original, "action": action, "destination": destination, "new_name": new_name})

    def process_file_to_folder(self, src, folder, new_name, action, undo_ops):
        os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, new_name)
        if not self.cfg.get("overwrite", False): target = unique_path(target)
        if action == "copy": shutil.copy2(src, target)
        else: shutil.move(src, target)
        self.register_op(undo_ops, src, action, target, os.path.basename(target))
        return target

    def process_aux_pdf(self, src, kind, community, undo_ops):
        year = str(self.cfg.get("aux_year") or datetime.now().year)
        cid = community.get("id_fmt")
        base = self.cfg.get("general_route", "")
        if not cid or not base:
            return None, "no_route"
        id4 = str(int(cid)).zfill(4) if cid.isdigit() else cid.zfill(4)
        year_folder = os.path.join(base, f"PH{id4}", f"{cid}-Año {year}")
        if kind == "visita":
            folder = year_folder
            new_name = os.path.basename(src)
        else:
            folder = os.path.join(year_folder, f"{cid}-PARTES DE TRABAJO")
            new_name = f"{cid}-{os.path.basename(src)}"
        target = self.process_file_to_folder(src, folder, new_name, "copy", undo_ops)
        return target, "ok"

    def move_error(self, src, folder, undo_ops):
        target = self.process_file_to_folder(src, folder, os.path.basename(src), "move", undo_ops)
        return target

    def run(self):
        undo_ops = []
        movement_rows = []
        try:
            cfg = dict(self.cfg)
            folder = cfg["watch_folder"]; ce = cfg["excel_file"]
            if not folder or not os.path.isdir(folder): raise ValueError("La carpeta de observación no existe.")
            if not ce or not os.path.isfile(ce): raise ValueError("No se ha encontrado el Excel de comunidades.")
            communities, aliases = read_communities(folder and ce, cfg.get("sheet", ""), cfg.get("community_columns", "B"), cfg.get("id_columns", "A"), cfg.get("route_column", "C"), cfg.get("alias_id_column", "R"), cfg.get("alias_name_column", "S"))
            concepts = read_concepts(cfg.get("concepts_excel", ""), cfg.get("sheet", ""), cfg.get("concept_column", "A"), cfg.get("concept_words_column", "B"))
            self.log_msg(f"Comunidades/alias cargados: {len(communities)} ({aliases} alias)")
            pdfs = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
            visita_folder = cfg.get("visitas_folder", "")
            partes_folder = cfg.get("partes_folder", "")
            # Los auxiliares se procesan además de la carpeta principal.
            aux_files = []
            for src_folder, kind in [(visita_folder, "visita"), (partes_folder, "partes")]:
                if src_folder and os.path.isdir(src_folder):
                    for f in os.listdir(src_folder):
                        p = os.path.join(src_folder, f)
                        if os.path.isfile(p): aux_files.append((p, kind))
            total = len(pdfs) + len(aux_files)
            self.set_progress(0, max(1, total))
            done = 0

            for pdf in pdfs:
                done += 1; self.set_progress(done, max(1,total))
                original_name = os.path.basename(pdf)
                self.log_msg("\n" + "="*70 + "\nProcesando factura: " + original_name)
                try:
                    text, source = extract_text(pdf, cfg.get("ocr", True))
                    community, status, matches = detect_community(text, communities)
                    if status == "multiple":
                        self.log_msg("⚠ Varias comunidades coinciden. Se mueve a Comunidad no encontrada.")
                        target = self.move_error(pdf, cfg["comunidad_no_encontrada_folder"], undo_ops)
                        movement_rows.append((original_name, target, os.path.basename(target))); continue
                    if status != "ok":
                        self.log_msg("⚠ Comunidad no encontrada.")
                        target = self.move_error(pdf, cfg["comunidad_no_encontrada_folder"], undo_ops)
                        movement_rows.append((original_name, target, os.path.basename(target))); continue
                    if not community.get("id_fmt"):
                        target = self.move_error(pdf, cfg["id_no_encontrado_folder"], undo_ops)
                        movement_rows.append((original_name, target, os.path.basename(target))); continue
                    concept, matched = detect_concept(text, concepts)
                    if not concept:
                        self.log_msg("⚠ Concepto no encontrado. Se mueve a Concepto no encontrado.")
                        target = self.move_error(pdf, cfg["concepto_no_encontrado_folder"], undo_ops)
                        movement_rows.append((original_name, target, os.path.basename(target)))
                        continue
                    date_info = extract_date(text, pdf)
                    target_name = make_filename(cfg["filename_template"], community, concept, date_info)
                    route, route_source = resolve_route(community, file_year(cfg, date_info), cfg)
                    if not route:
                        target = self.move_error(pdf, cfg["ruta_no_encontrada_folder"], undo_ops)
                        movement_rows.append((original_name, target, os.path.basename(target))); continue
                    os.makedirs(route, exist_ok=True)
                    target = os.path.join(route, target_name)
                    if not cfg.get("overwrite", False): target = unique_path(target)
                    action = "move" if cfg.get("move", False) else "copy"
                    if action == "move": shutil.move(pdf, target)
                    else: shutil.copy2(pdf, target)
                    undo_ops.append({"original": pdf, "action": action, "destination": target, "new_name": os.path.basename(target)})
                    movement_rows.append((original_name, target, os.path.basename(target)))
                    self.log_msg(f"✓ Factura {('movida' if action=='move' else 'copiada')}: {target}")
                except Exception as e:
                    self.log_msg(f"⚠ Error en {original_name}: {e}")

            for src, kind in aux_files:
                done += 1; self.set_progress(done, max(1,total))
                name = os.path.basename(src)
                self.log_msg("\n" + "="*70 + f"\nProcesando {kind}: {name}")
                if not src.lower().endswith(".pdf"):
                    target = self.move_error(src, cfg["otro_formato_folder"], undo_ops)
                    movement_rows.append((name, target, os.path.basename(target))); continue
                try:
                    text, source = extract_text(src, cfg.get("ocr", True))
                    community, status, matches = detect_community(text + "\n" + name, communities)
                    if status != "ok":
                        target = self.move_error(src, cfg["comunidad_no_encontrada_folder"], undo_ops)
                        movement_rows.append((name, target, os.path.basename(target))); continue
                    if not community.get("id_fmt"):
                        target = self.move_error(src, cfg["id_no_encontrado_folder"], undo_ops)
                        movement_rows.append((name, target, os.path.basename(target))); continue
                    target, st = self.process_aux_pdf(src, kind, community, undo_ops)
                    if st != "ok":
                        target = self.move_error(src, cfg["ruta_no_encontrada_folder"], undo_ops)
                    movement_rows.append((name, target, os.path.basename(target)))
                    self.log_msg(f"✓ {kind}: {target}")
                except Exception as e:
                    self.log_msg(f"⚠ Formato no legible: {e}")
                    target = self.move_error(src, cfg["otro_formato_folder"], undo_ops)
                    movement_rows.append((name, target, os.path.basename(target)))

            # Generar Word de movimientos al final.
            mov_folder = cfg.get("movimientos_folder", "")
            if mov_folder:
                os.makedirs(mov_folder, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                mov_path = os.path.join(mov_folder, f"movimientos{stamp}.docx")
                doc = Document()
                doc.add_heading(f"Movimientos {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", level=1)
                table = doc.add_table(rows=1, cols=3)
                hdr = table.rows[0].cells; hdr[0].text = "Nombre antes"; hdr[1].text = "Ruta destino"; hdr[2].text = "Nombre ahora"
                for before, dest, now in movement_rows:
                    cells = table.add_row().cells; cells[0].text = before; cells[1].text = dest or ""; cells[2].text = now or ""
                doc.save(mov_path)
                undo_ops.append({"original": None, "action": "delete", "destination": mov_path, "new_name": os.path.basename(mov_path)})
                self.log_msg(f"✓ Documento de movimientos: {mov_path}")

            with open(self.last_undo_file, "w", encoding="utf-8") as f: json.dump({"operations": undo_ops}, f, ensure_ascii=False, indent=2)
            self.log_msg("\n✓ Proceso terminado. Ya puedes usar 'Revertir cambios' para deshacer este último proceso.")
        except Exception as e:
            self.log_msg("✗ ERROR: " + repr(e))
        finally:
            self.set_running(False)

    def undo(self):
        if self.running: return
        if not os.path.isfile(self.last_undo_file):
            messagebox.showinfo("Revertir", "No hay un proceso registrado para revertir."); return
        try:
            with open(self.last_undo_file, encoding="utf-8") as f: data = json.load(f)
            ops = data.get("operations", [])
            if not ops: messagebox.showinfo("Revertir", "No hay cambios para revertir."); return
            if not messagebox.askyesno("Revertir cambios", "Se intentará deshacer el último proceso. ¿Continuar?"): return
            errors = []
            for op in reversed(ops):
                try:
                    action = op.get("action"); dest = op.get("destination"); orig = op.get("original")
                    if action == "delete":
                        if dest and os.path.exists(dest): os.remove(dest)
                    elif action == "copy":
                        if dest and os.path.exists(dest): os.remove(dest)
                    elif action == "move":
                        if dest and os.path.exists(dest):
                            os.makedirs(os.path.dirname(orig), exist_ok=True)
                            back = orig if not os.path.exists(orig) else unique_path(orig)
                            shutil.move(dest, back)
                except Exception as e: errors.append(str(e))
            if errors:
                self.log_msg("⚠ Reversión terminada con errores: " + " | ".join(errors))
            else:
                self.log_msg("✓ Último proceso revertido correctamente.")
            os.remove(self.last_undo_file)
        except Exception as e:
            messagebox.showerror("Revertir", str(e))

if __name__ == "__main__":
    root = tk.Tk(); App(root); root.mainloop()
