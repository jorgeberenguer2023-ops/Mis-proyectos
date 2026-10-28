import json, os, re, shutil, threading, unicodedata, time, copy
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
from tkinter import simpledialog

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
    "aurea_folder": "",
    "aurea_keyword": "AUREA",
    "propias_folder": "",
    "propias_terms": "",
    "varias_comunidades_folder": r"C:\\Users\\Jorge Berenguer\\Desktop\\Nueva carpeta (3)\\aprende\\Varias comunidades",
    "aux_year": str(datetime.now().year),
    "errores_revision_folder": "",
    "version_history_folder": "versiones_configuracion",
    "hidden_elements": [],
    "deleted_elements": [],
    "custom_fields": [],
    "app_version": "V18",
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
    reader_error = None
    page_count = 0
    try:
        reader = PdfReader(pdf_path)
        page_count = len(reader.pages)
        for page in reader.pages:
            try:
                parts.append(page.extract_text() or "")
            except Exception:
                parts.append("")
    except Exception as e:
        reader_error = e

    if reader_error is not None:
        # El archivo tiene extensión PDF pero no puede abrirse como PDF.
        # El proceso principal lo enviará a "Otro formato".
        raise ValueError(f"PDF no válido o no legible: {reader_error}")

    normal = "\n".join(parts)
    if use_ocr and OCR_AVAILABLE and (len(normalize_text(normal)) < 100 or not normalize_text(normal)):
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
            ocr_text = "\n".join(ocr_parts)
            return ocr_text, "OCR" if normalize_text(ocr_text) else "PDF sin texto"
        except Exception as e:
            # El PDF es estructuralmente válido, pero no hemos podido OCRizarlo.
            # Se conserva como PDF válido para que el flujo de comunidad decida.
            return normal, "texto (OCR falló: %s)" % e
    return normal, "texto PDF"


def contains_term(normalized_text, term):
    return re.search(r'(?<![A-Z0-9])' + re.escape(term) + r'(?![A-Z0-9])', normalized_text) is not None


def count_term_occurrences(normalized_text, term):
    if not term:
        return 0
    return len(re.findall(r"(?<![A-Z0-9])" + re.escape(term) + r"(?![A-Z0-9])", normalized_text))


def parse_propias_rules(value):
    """Parses rules like: casa, padre, AGF **2. **N means at least N occurrences."""
    rules = []
    for raw in split_terms(value):
        raw = raw.strip()
        if not raw:
            continue
        m = re.match(r"^(.*?)\s*\*\*\s*(\d+)\s*$", raw)
        if m:
            term = normalize_text(m.group(1))
            count = max(1, int(m.group(2)))
        else:
            term = normalize_text(raw)
            count = 1
        if term:
            rules.append((term, count))
    return rules


def matches_propias(text, rules):
    nt = normalize_text(text)
    for term, minimum in rules:
        if count_term_occurrences(nt, term) >= minimum:
            return True, term, minimum, count_term_occurrences(nt, term)
    return False, None, None, 0


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

    # Si aparecen dos o más comunidades con IDs diferentes en el mismo
    # documento, se considera un documento de VARIAS COMUNIDADES aunque una
    # coincidencia tenga una puntuación mayor que las demás.
    by_id = {}
    for score, row in scored:
        rid = row.get("id_fmt") or str(row.get("id") or "").strip()
        if rid:
            if rid not in by_id or score > by_id[rid][0]:
                by_id[rid] = (score, row)

    if len(by_id) > 1:
        matches = [item[1] for item in sorted(by_id.values(), key=lambda x: x[0], reverse=True)]
        return None, "multiple", matches

    # Una sola comunidad/ID: elegimos la coincidencia con mayor puntuación.
    max_score = max(x[0] for x in scored)
    best = [row for score, row in scored if score == max_score]
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




def find_movement_record(folder, original_name):
    if not folder or not os.path.isdir(folder): return None
    docs=sorted([os.path.join(folder,f) for f in os.listdir(folder) if f.lower().endswith('.docx')], key=lambda p: os.path.getmtime(p), reverse=True)
    for path in docs:
        try:
            doc=Document(path)
            for table in doc.tables:
                for row in table.rows[1:]:
                    cells=[c.text.strip() for c in row.cells]
                    if cells and cells[0].lower()==original_name.lower():
                        return {"before":cells[0],"route":cells[1] if len(cells)>1 else "","after":cells[2] if len(cells)>2 else "","word":path}
        except Exception: continue
    return None


def _redaction_rectangles(pdf_path):
    rects_by_page={}
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for pno,page in enumerate(pdf.pages):
                words=page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
                lines={}
                for w in words:
                    top=float(w.get('top',0)); key=round(top/3)*3
                    lines.setdefault(key,[]).append(w)
                rects=[]
                for _,ws in lines.items():
                    text=' '.join(w.get('text','') for w in ws)
                    nt=normalize_text(text)
                    if re.search(r'\b(?:CIF|NIF|CIFNIF)\b',nt):
                        x0=max(0,min(float(w.get('x0',0)) for w in ws)-2)
                        x1=min(float(page.width),max(float(w.get('x1',0)) for w in ws)+2)
                        top=min(float(w.get('top',0)) for w in ws)-2
                        bottom=max(float(w.get('bottom',0)) for w in ws)+3
                        rects.append((x0,top,x1,bottom))
                rects_by_page[pno]=rects
    except Exception:
        return {}
    return rects_by_page


def _redact_sensitive_text(text):
    # Oculta el CIF/NIF y su valor también en el texto que se imprime
    # en la página de revisión. No se conserva el identificador completo.
    s = text or ''
    patterns = [
        r'(?i)(\bCIF\b\s*[:.\-]?\s*)([A-Z][A-Z0-9\-]{7,10})',
        r'(?i)(\bNIF\b\s*[:.\-]?\s*)([0-9XYZ][0-9XYZ][0-9XYZ0-9]{6,8}[A-Z])',
    ]
    for pat in patterns:
        s = re.sub(pat, lambda m: m.group(1) + '████████', s)
    # Por seguridad, también oculta identificadores españoles que queden
    # inmediatamente después de CIF/NIF aunque el OCR haya separado el texto.
    return s

def _draw_wrapped_lines(c, text, x, y, width, font='Helvetica', size=8, leading=11):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    c.setFont(font,size)
    maxw=width
    for raw in (text or '').splitlines() or ['']:
        line=raw
        if not line:
            c.drawString(x,y,''); y-=leading
            continue
        while stringWidth(line,font,size)>maxw and len(line)>10:
            cut=max(10,int(len(line)*maxw/stringWidth(line,font,size))-1)
            c.drawString(x,y,line[:cut]); y-=leading; line=line[cut:]
            if y<45:
                c.showPage(); y=A4[1]-45; c.setFont(font,size)
        c.drawString(x,y,line[:250]); y-=leading
        if y<45:
            c.showPage(); y=A4[1]-45; c.setFont(font,size)
    return y

def _append_review_pages(writer, read_text, metadata):
    from io import BytesIO
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    buf=BytesIO(); c=canvas.Canvas(buf,pagesize=A4); width,height=A4
    x=40; y=height-45
    c.setFont('Helvetica-Bold',14); c.drawString(x,y,'Revisión del documento'); y-=25
    c.setFont('Helvetica',9)
    fields=[
        ('Archivo original', metadata.get('original','')),
        ('Comunidad encontrada', metadata.get('community','')),
        ('ID', metadata.get('id','')),
        ('Ruta anterior', metadata.get('route','')),
        ('Nombre guardado', metadata.get('after','')),
        ('Texto señalado como error', metadata.get('error','')),
        ('Comentario', metadata.get('comment','')),
    ]
    for label,value in fields:
        value=str(value or 'No disponible')
        c.setFont('Helvetica-Bold',9); c.drawString(x,y,label+':'); y-=12
        c.setFont('Helvetica',9)
        for part in value.splitlines() or ['']:
            # wrap using a small helper
            from reportlab.pdfbase.pdfmetrics import stringWidth
            line=part
            while stringWidth(line,'Helvetica',9)>width-80 and len(line)>10:
                cut=max(10,int(len(line)*(width-80)/stringWidth(line,'Helvetica',9))-1)
                c.drawString(x+8,y,line[:cut]); y-=12; line=line[cut:]
                if y<60: c.showPage(); y=height-45; c.setFont('Helvetica',9)
            c.drawString(x+8,y,line[:250]); y-=12
        y-=5
        if y<60: c.showPage(); y=height-45
    c.setFont('Helvetica-Bold',12); c.drawString(x,y,'Texto leído para la revisión'); y-=18
    y=_draw_wrapped_lines(c,_redact_sensitive_text(read_text),x,y,width-80,'Helvetica',8,11)
    c.save(); buf.seek(0)
    for page in PdfReader(buf).pages: writer.add_page(page)

def create_redacted_review_copy(pdf_path, output_path, read_text, metadata=None):
    """Create a reviewed copy, blacking CIF/NIF lines and appending the text read."""
    from io import BytesIO
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import black
    rects=_redaction_rectangles(pdf_path)
    try:
        reader=PdfReader(pdf_path); writer=__import__('pypdf').PdfWriter()
        if rects and any(rects.values()):
            for i,page in enumerate(reader.pages):
                page_rects=rects.get(i,[])
                if page_rects:
                    width=float(page.mediabox.width); height=float(page.mediabox.height)
                    buf=BytesIO(); c=canvas.Canvas(buf,pagesize=(width,height)); c.setFillColor(black)
                    for x0,top,x1,bottom in page_rects:
                        y=height-bottom; c.rect(x0,y,max(1,x1-x0),max(1,bottom-top),fill=1,stroke=0)
                    c.save(); buf.seek(0); page.merge_page(PdfReader(buf).pages[0])
                writer.add_page(page)
            _append_review_pages(writer,read_text,metadata or {})
            with open(output_path,'wb') as f: writer.write(f)
            return True
    except Exception:
        pass
    # Scanned-PDF fallback: render pages, OCR just now, and black out CIF/NIF lines on the rendered image.
    if OCR_AVAILABLE:
        try:
            doc=pdfium.PdfDocument(pdf_path); images=[]
            for i in range(len(doc)):
                img=doc[i].render(scale=2.2).to_pil_image().convert('RGB')
                data=pytesseract.image_to_data(img,lang='spa+eng',output_type=pytesseract.Output.DICT)
                lines={}
                for j,txt in enumerate(data.get('text',[])):
                    txt=(txt or '').strip()
                    if not txt: continue
                    key=(data['block_num'][j],data['par_num'][j],data['line_num'][j])
                    lines.setdefault(key,[]).append(j)
                for idxs in lines.values():
                    line=' '.join((data['text'][j] or '') for j in idxs)
                    if re.search(r'\b(?:CIF|NIF)\b',normalize_text(line)):
                        xs=[data['left'][j] for j in idxs]; ys=[data['top'][j] for j in idxs]; x2=[data['left'][j]+data['width'][j] for j in idxs]; y2=[data['top'][j]+data['height'][j] for j in idxs]
                        from PIL import ImageDraw
                        ImageDraw.Draw(img).rectangle((max(0,min(xs)-6),max(0,min(ys)-4),min(img.width,max(x2)+6),min(img.height,max(y2)+6)),fill='black')
                images.append(img)
            if images:
                tmp=output_path+'.tmp.pdf'; images[0].save(tmp,'PDF',resolution=150.0,save_all=True,append_images=images[1:])
                reader=PdfReader(tmp); writer=__import__('pypdf').PdfWriter()
                for page in reader.pages: writer.add_page(page)
                _append_review_pages(writer,read_text,metadata or {})
                with open(output_path,'wb') as f: writer.write(f)
                try: os.remove(tmp)
                except OSError: pass
                return True
        except Exception:
            pass
    # Last fallback preserves the source and appends the text page; this is only used if redaction is technically impossible.
    try:
        reader=PdfReader(pdf_path); writer=__import__('pypdf').PdfWriter()
        for page in reader.pages: writer.add_page(page)
        _append_review_pages(writer,read_text,metadata or {})
        with open(output_path,'wb') as f: writer.write(f)
        return False
    except Exception:
        return False

def extract_text_for_review_once(pdf_path, use_ocr):
    return extract_text(pdf_path, use_ocr)[0]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de documentos V18")
        self.root.geometry("1080x900")
        self.cfg = self.load_cfg()
        self.v = {}
        self.ui_widgets = {}
        self.ui_vars = {}
        self.running = False
        self.last_undo_file = os.path.join(os.path.dirname(os.path.abspath(CONFIG_FILE)), "ultimo_proceso_undo.json")
        self.version_dir = os.path.join(os.path.dirname(os.path.abspath(CONFIG_FILE)), self.cfg.get("version_history_folder", "versiones_configuracion"))
        self.error_review_records = []
        self.current_log_context = ""
        self.build()
        self.load_ui()

    def load_cfg(self):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                d = DEFAULT_CONFIG.copy(); d.update(json.load(f)); return d
        except Exception:
            return DEFAULT_CONFIG.copy()

    def is_hidden(self, key):
        return key in set(self.cfg.get("hidden_elements", [])) or key in set(self.cfg.get("deleted_elements", []))

    def build(self):
        for w in self.root.winfo_children():
            w.destroy()
        self.v = {}; self.ui_widgets = {}; self.ui_vars = {}
        outer = ttk.Frame(self.root, padding=8); outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        main = ttk.Frame(canvas, padding=8)
        window_id = canvas.create_window((0, 0), window=main, anchor="nw")
        main.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: (canvas.yview_scroll(int(-e.delta / 120), "units"), "break")[1], add="+")
        ttk.Label(main, text="Organizador de documentos V18", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        top = ttk.Frame(main); top.pack(fill="x", pady=4)
        if not self.is_hidden("config_button"):
            ttk.Button(top, text="⚙ Configuración", command=self.open_settings).pack(side="left", padx=3)
        if not self.is_hidden("versions_button"):
            ttk.Button(top, text="Control de versiones", command=self.open_versions).pack(side="left", padx=3)
        if not self.is_hidden("report_button"):
            ttk.Button(top, text="Informe configuración", command=self.export_config_report).pack(side="left", padx=3)

        self.entry(main, "Carpeta de observación", "watch_folder", self.folder)
        self.entry(main, "Excel comunidades", "excel_file", self.file_community)
        self.entry(main, "Columnas comunidad", "community_columns")
        self.hint(main, "community_hint", "Ej.: B,D,F  (se emparejan con A,C,E)")
        self.entry(main, "Columnas ID", "id_columns")
        self.hint(main, "id_hint", "Ej.: A,C,E  (A→B, C→D, E→F)")
        self.entry(main, "Columna ruta (opcional)", "route_column")
        self.entry(main, "ID alternativos (R)", "alias_id_column")
        self.entry(main, "Nombres alternativos (S)", "alias_name_column")
        self.hint(main, "alias_hint", "R y S permiten añadir nombres alternativos: R=ID, S=nombre alternativo.")
        self.entry(main, "Ruta general (base)", "general_route", self.folder)
        self.hint(main, "route_hint", r"Ej.: ...\PH0002\002-Año 2026\002-FACTURAS 2026")
        self.entry(main, "Excel conceptos", "concepts_excel", self.file_concepts)
        self.entry(main, "Columna concepto", "concept_column")
        self.entry(main, "Primera columna palabras", "concept_words_column")
        self.entry(main, "Hoja (vacío=automática)", "sheet")
        self.entry(main, "Plantilla nombre", "filename_template")
        self.entry(main, "Año carpetas auxiliares", "aux_year")
        self.separator(main, "sep1")
        self.section(main, "sec_aux", "Carpetas auxiliares")
        for label, key in [
            ("Detalles de la Visita", "visitas_folder"), ("PARTES DE TRABAJO", "partes_folder"),
            ("Otro formato", "otro_formato_folder"), ("Comunidad no encontrada", "comunidad_no_encontrada_folder"),
            ("Id no encontrado", "id_no_encontrado_folder"), ("Concepto no encontrado", "concepto_no_encontrado_folder"),
            ("Ruta no encontrada", "ruta_no_encontrada_folder"), ("Movimientos", "movimientos_folder"),
            ("Errores revisados", "errores_revision_folder")]:
            self.entry(main, label, key, self.folder_for(key))

        self.separator(main, "sep2")
        self.section(main, "sec_special", "Rutas especiales")
        self.entry(main, "AUREA - ruta", "aurea_folder", self.folder_for("aurea_folder"))
        self.entry(main, "AUREA - palabra", "aurea_keyword")
        self.hint(main, "aurea_hint", "Si aparece esta palabra, se mueve aquí sin comprobar comunidad ni ID.")
        self.entry(main, "PROPIAS - ruta", "propias_folder", self.folder_for("propias_folder"))
        self.entry(main, "PROPIAS - palabras", "propias_terms")
        self.hint(main, "propias_hint", "Ej.: casa, padre, AGF **2  → AGF debe aparecer al menos 2 veces.")
        self.entry(main, "VARIAS COMUNIDADES - ruta", "varias_comunidades_folder", self.folder_for("varias_comunidades_folder"))
        self.hint(main, "varias_hint", "Los documentos con varias comunidades se mueven aquí con su nombre original.")

        if not self.is_hidden("ocr_checkbox"):
            self.ocr = tk.BooleanVar(value=True); self.ocr_cb = ttk.Checkbutton(main, text="Usar OCR", variable=self.ocr); self.ocr_cb.pack(anchor="w"); self.ui_widgets["ocr_checkbox"] = self.ocr_cb
        else: self.ocr = tk.BooleanVar(value=True)
        if not self.is_hidden("simulate_button"):
            self.simulate_btn = ttk.Button(main, text="Simular proceso (sin mover ni crear carpetas)", command=self.start_simulation); self.simulate_btn.pack(anchor="w", pady=2); self.ui_widgets["simulate_button"] = self.simulate_btn
        if not self.is_hidden("move_checkbox"):
            self.move = tk.BooleanVar(value=False); w=ttk.Checkbutton(main, text="Mover en vez de copiar facturas", variable=self.move); w.pack(anchor="w"); self.ui_widgets["move_checkbox"]=w
        else: self.move=tk.BooleanVar(value=False)
        if not self.is_hidden("overwrite_checkbox"):
            self.overwrite = tk.BooleanVar(value=False); w=ttk.Checkbutton(main, text="Sobrescribir", variable=self.overwrite); w.pack(anchor="w"); self.ui_widgets["overwrite_checkbox"]=w
        else: self.overwrite=tk.BooleanVar(value=False)

        for item in self.cfg.get("custom_fields", []):
            if item.get("key") and not self.is_hidden(item["key"]):
                self.entry(main, item.get("label", item["key"]), item["key"])

        b = ttk.Frame(main); b.pack(fill="x", pady=8)
        if not self.is_hidden("save_button"):
            self.save_btn = ttk.Button(b, text="Guardar configuración", command=self.save); self.save_btn.pack(side="left", padx=4); self.ui_widgets["save_button"]=self.save_btn
        if not self.is_hidden("run_button"):
            self.run_btn = ttk.Button(b, text="Organizar", command=self.start); self.run_btn.pack(side="left", padx=4); self.ui_widgets["run_button"]=self.run_btn
        if not self.is_hidden("undo_button"):
            self.undo_btn = ttk.Button(b, text="↩ Revertir cambios", command=self.undo); self.undo_btn.pack(side="left", padx=4); self.ui_widgets["undo_button"]=self.undo_btn
        self.pb = ttk.Progressbar(main, mode="determinate"); self.pb.pack(fill="x", pady=6); self.ui_widgets["progress"] = self.pb
        if not self.is_hidden("log_label"): ttk.Label(main, text="Registro detallado").pack(anchor="w")
        if not self.is_hidden("log_text"):
            searchf=ttk.Frame(main); searchf.pack(fill="x", pady=2)
            ttk.Label(searchf,text="Buscar en registro:").pack(side="left")
            self.search_var=tk.StringVar(); ttk.Entry(searchf,textvariable=self.search_var).pack(side="left",fill="x",expand=True,padx=4)
            ttk.Button(searchf,text="Buscar",command=self.search_log).pack(side="left")
            ttk.Button(searchf,text="Siguiente",command=lambda:self.search_log(True)).pack(side="left",padx=2)
            self.log = tk.Text(main, height=16, wrap="word")
            self.log.pack(fill="both", expand=True)
            self.log.tag_configure("error_selection", background="yellow", foreground="red")
            self.log.bind("<KeyPress>", self.on_log_key)
            self.ui_widgets["log_text"] = self.log

    def entry(self, parent, label, key, browse=None):
        if self.is_hidden(key): return
        f=ttk.Frame(parent); f.pack(fill="x",pady=2); self.ui_widgets[key]=f
        ttk.Label(f,text=label,width=29).pack(side="left")
        var=tk.StringVar(); self.v[key]=var; self.ui_vars[key]=var
        ttk.Entry(f,textvariable=var).pack(side="left",fill="x",expand=True)
        if browse: ttk.Button(f,text="Examinar",command=browse).pack(side="left",padx=4)

    def hint(self, parent, key, text):
        if self.is_hidden(key): return
        w=ttk.Label(parent,text=text); w.pack(anchor="w",padx=4); self.ui_widgets[key]=w
    def separator(self,parent,key):
        if self.is_hidden(key): return
        w=ttk.Separator(parent); w.pack(fill="x",pady=8); self.ui_widgets[key]=w
    def section(self,parent,key,text):
        if self.is_hidden(key): return
        w=ttk.Label(parent,text=text,font=("Segoe UI",10,"bold")); w.pack(anchor="w"); self.ui_widgets[key]=w
    def folder_for(self,key):
        def choose():
            p=filedialog.askdirectory();
            if p and key in self.v: self.v[key].set(p)
        return choose
    def load_ui(self):
        for k,var in self.v.items(): var.set(self.cfg.get(k,""))
        self.ocr.set(self.cfg.get("ocr",True)); self.move.set(self.cfg.get("move",False)); self.overwrite.set(self.cfg.get("overwrite",False))
    def save(self):
        for k,var in self.v.items(): self.cfg[k]=var.get().strip()
        self.cfg["ocr"]=self.ocr.get(); self.cfg["move"]=self.move.get(); self.cfg["overwrite"]=self.overwrite.get(); self.cfg["app_version"]="V18"
        os.makedirs(self.version_dir,exist_ok=True)
        stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        with open(CONFIG_FILE,"w",encoding="utf-8") as f: json.dump(self.cfg,f,ensure_ascii=False,indent=2)
        with open(os.path.join(self.version_dir,f"config_{stamp}.json"),"w",encoding="utf-8") as f: json.dump(self.cfg,f,ensure_ascii=False,indent=2)
        self.log_msg("✓ Configuración guardada y versión archivada: "+stamp)
    def folder(self):
        p=filedialog.askdirectory();
        if p: self.v["watch_folder"].set(p)
    def file_community(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")]);
        if p: self.v["excel_file"].set(p)
    def file_concepts(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")]);
        if p: self.v["concepts_excel"].set(p)
    def log_msg(self,s):
        def ui():
            if hasattr(self,"log"):
                self.log.insert("end",s+"\n"); self.log.see("end")
        self.root.after(0,ui)
    def set_progress(self,value,maximum): self.root.after(0,lambda:(self.pb.configure(maximum=maximum,value=value)))
    def set_running(self,value):
        self.running=value
        def ui():
            if hasattr(self,"run_btn"): self.run_btn.configure(state="disabled" if value else "normal")
            if hasattr(self,"save_btn"): self.save_btn.configure(state="disabled" if value else "normal")
            if hasattr(self,"simulate_btn"): self.simulate_btn.configure(state="disabled" if value else "normal")
        self.root.after(0,ui)
    def start(self):
        if self.running:return
        self.save(); self.set_running(True); threading.Thread(target=lambda:self.run(False),daemon=True).start()
    def start_simulation(self):
        if self.running:return
        self.save(); self.set_running(True); threading.Thread(target=lambda:self.run(True),daemon=True).start()
    def register_op(self,undo_ops,original,action,destination,new_name): undo_ops.append({"original":original,"action":action,"destination":destination,"new_name":new_name})
    def ensure_folder(self,folder,undo_ops,simulate=False):
        if not folder:return
        if simulate:return
        folder=os.path.abspath(folder); missing=[]; current=folder
        while current and not os.path.exists(current):
            missing.append(current); parent=os.path.dirname(current)
            if parent==current:break
            current=parent
        os.makedirs(folder,exist_ok=True)
        for created in reversed(missing): undo_ops.append({"original":None,"action":"rmdir","destination":created,"new_name":os.path.basename(created)})
    def process_file_to_folder(self,src,folder,new_name,action,undo_ops,simulate=False):
        self.ensure_folder(folder,undo_ops,simulate); target=os.path.join(folder,new_name)
        if not self.cfg.get("overwrite",False):target=unique_path(target)
        if not simulate:
            if action=="copy":shutil.copy2(src,target)
            else:shutil.move(src,target)
            self.register_op(undo_ops,src,action,target,os.path.basename(target))
        return target
    def process_aux_pdf(self,src,kind,community,undo_ops,simulate=False):
        year=str(self.cfg.get("aux_year") or datetime.now().year); cid=community.get("id_fmt"); base=self.cfg.get("general_route","")
        if not cid or not base:return None,"no_route"
        id4=str(int(cid)).zfill(4) if cid.isdigit() else cid.zfill(4); year_folder=os.path.join(base,f"PH{id4}",f"{cid}-Año {year}")
        if kind=="visita":folder=year_folder; new_name=os.path.basename(src)
        else:folder=os.path.join(year_folder,f"{cid}-PARTES DE TRABAJO"); new_name=f"{cid}-{os.path.basename(src)}"
        return self.process_file_to_folder(src,folder,new_name,"copy",undo_ops,simulate),"ok"
    def move_error(self,src,folder,undo_ops,simulate=False): return self.process_file_to_folder(src,folder,os.path.basename(src),"move",undo_ops,simulate)
    def run(self, simulate=False):
        undo_ops=[]; movement_rows=[]
        try:
            cfg=dict(self.cfg); folder=cfg.get("watch_folder",""); ce=cfg.get("excel_file","")
            if not folder or not os.path.isdir(folder): raise ValueError("La carpeta de observación no existe.")
            if not ce or not os.path.isfile(ce): raise ValueError("No se ha encontrado el Excel de comunidades.")
            communities,aliases=read_communities(ce,cfg.get("sheet",""),cfg.get("community_columns","B"),cfg.get("id_columns","A"),cfg.get("route_column",""),cfg.get("alias_id_column","R"),cfg.get("alias_name_column","S"))
            concepts=read_concepts(cfg.get("concepts_excel",""),cfg.get("sheet",""),cfg.get("concept_column","A"),cfg.get("concept_words_column","B"))
            self.log_msg(f"Comunidades/alias cargados: {len(communities)} ({aliases} alias)")
            pdfs=[os.path.join(folder,f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder,f)) and f.lower().endswith('.pdf')]
            visita_folder=cfg.get("visitas_folder",""); partes_folder=cfg.get("partes_folder",""); aux_files=[]
            for src_folder,kind in [(visita_folder,"visita"),(partes_folder,"partes")]:
                if src_folder and os.path.isdir(src_folder):
                    for f in os.listdir(src_folder):
                        p=os.path.join(src_folder,f)
                        if os.path.isfile(p): aux_files.append((p,kind))
            total=len(pdfs)+len(aux_files); self.set_progress(0,max(1,total)); done=0
            propria_rules=parse_propias_rules(cfg.get("propias_terms",""))
            for pdf in pdfs:
                done+=1; self.set_progress(done,max(1,total)); original_name=os.path.basename(pdf)
                self.log_msg("\n"+"="*70+"\nProcesando factura: "+original_name)
                try:
                    text,source=extract_text(pdf,cfg.get("ocr",True))
                    nt=normalize_text(text)
                    aurea_keyword=str(cfg.get("aurea_keyword","AUREA") or "AUREA").strip()
                    if aurea_keyword and contains_term(nt,normalize_text(aurea_keyword)) and cfg.get("aurea_folder",""):
                        target=os.path.join(cfg["aurea_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["aurea_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); self.log_msg(f"✓ AUREA detectado. {'Simulado hacia' if simulate else 'Movido a'}: {target}"); continue
                    matched_prop,prop_term,prop_min,prop_count=matches_propias(text,propria_rules)
                    if matched_prop and cfg.get("propias_folder",""):
                        target=os.path.join(cfg["propias_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["propias_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); self.log_msg(f"✓ PROPIA detectada ({prop_term}: {prop_count} >= {prop_min}). {'Simulado hacia' if simulate else 'Movido a'}: {target}"); continue
                    community,status,matches=detect_community(text,communities)
                    if status=="multiple":
                        target=os.path.join(cfg["varias_comunidades_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["varias_comunidades_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); self.log_msg("⚠ Varias comunidades: "+target); continue
                    if status!="ok":
                        target=os.path.join(cfg["comunidad_no_encontrada_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["comunidad_no_encontrada_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); self.log_msg("⚠ Comunidad no encontrada: "+target); continue
                    if not community.get("id_fmt"):
                        target=os.path.join(cfg["id_no_encontrado_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["id_no_encontrado_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); continue
                    concept,matched=detect_concept(text,concepts)
                    if not concept:
                        target=os.path.join(cfg["concepto_no_encontrado_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["concepto_no_encontrado_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); self.log_msg("⚠ Concepto no encontrado: "+target); continue
                    date_info=extract_date(text,pdf); target_name=make_filename(cfg["filename_template"],community,concept,date_info)
                    route,route_source=resolve_route(community,file_year(cfg,date_info),cfg)
                    if not route:
                        target=os.path.join(cfg["ruta_no_encontrada_folder"],original_name)
                        if not simulate: target=self.move_error(pdf,cfg["ruta_no_encontrada_folder"],undo_ops,False)
                        movement_rows.append((original_name,target,os.path.basename(target))); continue
                    target=os.path.join(route,target_name)
                    if not cfg.get("overwrite",False) and not simulate: target=unique_path(target)
                    action="move" if cfg.get("move",False) else "copy"
                    if not simulate:
                        self.ensure_folder(route,undo_ops,False)
                        if action=="move": shutil.move(pdf,target)
                        else: shutil.copy2(pdf,target)
                        self.register_op(undo_ops,pdf,action,target,os.path.basename(target))
                    movement_rows.append((original_name,target,os.path.basename(target)))
                    self.log_msg(f"✓ {'SIMULADO: ' if simulate else ''}Factura: {target}")
                except Exception as e:
                    self.log_msg(f"⚠ PDF no legible / formato no compatible en {original_name}: {e}")
                    target=os.path.join(cfg.get("otro_formato_folder",""),original_name)
                    if not simulate:
                        target=self.move_error(pdf,cfg.get("otro_formato_folder",""),undo_ops,False)
                    movement_rows.append((original_name,target,os.path.basename(target)))
            for src,kind in aux_files:
                done+=1; self.set_progress(done,max(1,total)); name=os.path.basename(src); self.log_msg("\n"+"="*70+f"\nProcesando {kind}: {name}")
                if not src.lower().endswith('.pdf'):
                    target=os.path.join(cfg.get("otro_formato_folder",""),name)
                    if not simulate: target=self.move_error(src,cfg.get("otro_formato_folder",""),undo_ops,False)
                    movement_rows.append((name,target,os.path.basename(target))); continue
                try:
                    text,source=extract_text(src,cfg.get("ocr",True)); combined=text+"\n"+name
                    if cfg.get("aurea_folder","") and cfg.get("aurea_keyword","") and contains_term(normalize_text(combined),normalize_text(cfg["aurea_keyword"])):
                        target=os.path.join(cfg["aurea_folder"],name)
                        if not simulate: target=self.move_error(src,cfg["aurea_folder"],undo_ops,False)
                        movement_rows.append((name,target,os.path.basename(target))); continue
                    matched_prop,_,_,_=matches_propias(combined,propria_rules)
                    if matched_prop and cfg.get("propias_folder",""):
                        target=os.path.join(cfg["propias_folder"],name)
                        if not simulate: target=self.move_error(src,cfg["propias_folder"],undo_ops,False)
                        movement_rows.append((name,target,os.path.basename(target))); continue
                    community,status,matches=detect_community(combined,communities)
                    if status=="multiple":
                        target=os.path.join(cfg["varias_comunidades_folder"],name)
                        if not simulate: target=self.move_error(src,cfg["varias_comunidades_folder"],undo_ops,False)
                    elif status!="ok":
                        target=os.path.join(cfg["comunidad_no_encontrada_folder"],name)
                        if not simulate: target=self.move_error(src,cfg["comunidad_no_encontrada_folder"],undo_ops,False)
                    elif not community.get("id_fmt"):
                        target=os.path.join(cfg["id_no_encontrado_folder"],name)
                        if not simulate: target=self.move_error(src,cfg["id_no_encontrado_folder"],undo_ops,False)
                    else:
                        target,_st=self.process_aux_pdf(src,kind,community,undo_ops,simulate)
                        if _st!="ok":
                            target=os.path.join(cfg["ruta_no_encontrada_folder"],name)
                            if not simulate: target=self.move_error(src,cfg["ruta_no_encontrada_folder"],undo_ops,False)
                    movement_rows.append((name,target,os.path.basename(target) if target else "")); self.log_msg(f"✓ {'SIMULADO: ' if simulate else ''}{kind}: {target}")
                except Exception as e:
                    target=os.path.join(cfg.get("otro_formato_folder",""),name)
                    if not simulate: target=self.move_error(src,cfg.get("otro_formato_folder",""),undo_ops,False)
                    movement_rows.append((name,target,os.path.basename(target))); self.log_msg(f"⚠ Formato no legible: {e}")
            mov_folder=cfg.get("movimientos_folder","")
            if mov_folder:
                if simulate:
                    actual_mov_folder = mov_folder if os.path.isdir(mov_folder) else os.path.dirname(os.path.abspath(CONFIG_FILE))
                    mov_path=os.path.join(actual_mov_folder,f"movimientos{datetime.now().strftime('%Y%m%d_%H%M%S')}_SIMULACION.docx")
                else:
                    self.ensure_folder(mov_folder,undo_ops,False); mov_path=os.path.join(mov_folder,f"movimientos{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")
                doc=Document(); doc.add_heading(f"Movimientos {'SIMULACIÓN ' if simulate else ''}{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",level=1)
                table=doc.add_table(rows=1,cols=3); hdr=table.rows[0].cells; hdr[0].text="Nombre antes"; hdr[1].text="Ruta destino"; hdr[2].text="Nombre ahora"
                for before,dest,now in movement_rows:
                    cells=table.add_row().cells; cells[0].text=before; cells[1].text=dest or ""; cells[2].text=now or ""
                doc.save(mov_path)
                # Deliberadamente NO se añade el Word de movimientos a undo_ops.
                self.log_msg(f"✓ Documento de movimientos {'de simulación ' if simulate else ''}: {mov_path}")
            if not simulate:
                with open(self.last_undo_file,"w",encoding="utf-8") as f: json.dump({"operations":undo_ops},f,ensure_ascii=False,indent=2)
                self.log_msg("\n✓ Proceso terminado. El Word de movimientos NO se eliminará al revertir.")
            else: self.log_msg("\n✓ Simulación terminada. No se han movido archivos ni creado carpetas.")
        except Exception as e: self.log_msg("✗ ERROR: "+repr(e))
        finally: self.set_running(False)

    def undo(self):
        if self.running:return
        if not os.path.isfile(self.last_undo_file): messagebox.showinfo("Revertir","No hay un proceso registrado para revertir."); return
        try:
            with open(self.last_undo_file,encoding="utf-8") as f:data=json.load(f)
            ops=data.get("operations",[])
            if not ops: messagebox.showinfo("Revertir","No hay cambios para revertir."); return
            if not messagebox.askyesno("Revertir cambios","Se desharán los movimientos/cambios del último proceso, pero NO se borrarán los documentos de movimientos. ¿Continuar?"):return
            errors=[]
            for op in reversed(ops):
                try:
                    action=op.get("action"); dest=op.get("destination"); orig=op.get("original")
                    if action=="copy":
                        if dest and os.path.exists(dest):os.remove(dest)
                    elif action=="move":
                        if dest and os.path.exists(dest):
                            back=orig if orig and not os.path.exists(orig) else unique_path(orig) if orig else None
                            if back: os.makedirs(os.path.dirname(back),exist_ok=True); shutil.move(dest,back)
                    elif action=="rmdir":
                        if dest and os.path.isdir(dest) and not os.listdir(dest):os.rmdir(dest)
                except Exception as e: errors.append(str(e))
            self.log_msg(("⚠ Reversión terminada con errores: "+" | ".join(errors)) if errors else "✓ Cambios revertidos. Los documentos de movimientos se conservaron.")
            os.remove(self.last_undo_file)
        except Exception as e: messagebox.showerror("Revertir",str(e))

    def open_settings(self):
        win=tk.Toplevel(self.root); win.title("Configuración y elementos de interfaz"); win.geometry("720x600")
        ttk.Label(win,text="Elementos de la interfaz",font=("Segoe UI",12,"bold")).pack(anchor="w",padx=10,pady=8)
        frame=ttk.Frame(win); frame.pack(fill="both",expand=True,padx=10)
        lb=tk.Listbox(frame,selectmode="extended"); lb.pack(side="left",fill="both",expand=True)
        all_keys=list(self.ui_widgets.keys())
        for k in all_keys: lb.insert("end",k)
        def apply_state():
            sel=list(lb.curselection())
            action=action_var.get()
            if action in ("ocultar","eliminar"):
                keylist=self.cfg.get("hidden_elements",[]) if action=="ocultar" else self.cfg.get("deleted_elements",[])
                for i in sel:
                    k=all_keys[i]
                    if k not in keylist:keylist.append(k)
                self.cfg["hidden_elements"] = keylist if action=="ocultar" else self.cfg.get("hidden_elements",[])
                self.cfg["deleted_elements"] = keylist if action=="eliminar" else self.cfg.get("deleted_elements",[])
            elif action=="visible":
                self.cfg["hidden_elements"]=[x for x in self.cfg.get("hidden_elements",[]) if x not in [all_keys[i] for i in sel]]
                self.cfg["deleted_elements"]=[x for x in self.cfg.get("deleted_elements",[]) if x not in [all_keys[i] for i in sel]]
            self.save(); self.build(); self.load_ui(); win.destroy()
        right=ttk.Frame(frame); right.pack(side="left",fill="y",padx=10)
        action_var=tk.StringVar(value="ocultar")
        for txt,val in [("Ocultar","ocultar"),("Hacer visible","visible"),("Eliminar","eliminar")]: ttk.Radiobutton(right,text=txt,value=val,variable=action_var).pack(anchor="w",pady=3)
        ttk.Button(right,text="Aplicar",command=apply_state).pack(fill="x",pady=8)
        ttk.Separator(right).pack(fill="x",pady=8)
        ttk.Label(right,text="Añadir elemento").pack(anchor="w")
        lab=tk.StringVar(); ttk.Entry(right,textvariable=lab).pack(fill="x")
        def add_custom():
            label=lab.get().strip()
            if not label:return
            key="custom_"+re.sub(r'[^a-z0-9]+','_',label.lower()).strip('_')+"_"+datetime.now().strftime('%H%M%S')
            self.cfg.setdefault("custom_fields",[]).append({"key":key,"label":label})
            self.save(); self.build(); self.load_ui(); win.destroy()
        ttk.Button(right,text="Añadir campo",command=add_custom).pack(fill="x",pady=5)
        ttk.Button(win,text="Cerrar",command=win.destroy).pack(pady=8)

    def open_versions(self):
        os.makedirs(self.version_dir,exist_ok=True); win=tk.Toplevel(self.root); win.title("Control de versiones"); win.geometry("700x500")
        ttk.Label(win,text="Últimas versiones de configuración",font=("Segoe UI",12,"bold")).pack(anchor="w",padx=10,pady=8)
        lb=tk.Listbox(win); lb.pack(fill="both",expand=True,padx=10,pady=5)
        files=sorted([f for f in os.listdir(self.version_dir) if f.endswith('.json')],reverse=True)
        for f in files[:50]:lb.insert("end",f)
        def restore():
            sel=lb.curselection()
            if not sel:return
            path=os.path.join(self.version_dir,lb.get(sel[0]))
            if messagebox.askyesno("Restaurar","¿Restaurar esta configuración?",parent=win):
                with open(path,encoding='utf-8') as f:self.cfg=json.load(f)
                with open(CONFIG_FILE,'w',encoding='utf-8') as f:json.dump(self.cfg,f,ensure_ascii=False,indent=2)
                self.build(); self.load_ui(); win.destroy(); self.log_msg("✓ Configuración restaurada desde "+os.path.basename(path))
        ttk.Button(win,text="Restaurar versión seleccionada",command=restore).pack(pady=5)
        ttk.Button(win,text="Cerrar",command=win.destroy).pack(pady=5)

    def export_config_report(self):
        path=filedialog.asksaveasfilename(title="Guardar informe de configuración",defaultextension=".docx",filetypes=[("Word","*.docx"),("Texto","*.txt")])
        if not path:return
        try:
            if path.lower().endswith('.docx'):
                doc=Document(); doc.add_heading("Informe de configuración - Organizador V18",level=1); doc.add_paragraph(datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
                for k,v in self.cfg.items(): doc.add_paragraph(f"{k}: {v}")
                doc.save(path)
            else:
                with open(path,'w',encoding='utf-8') as f: json.dump(self.cfg,f,ensure_ascii=False,indent=2)
            self.log_msg("✓ Informe de configuración guardado: "+path); messagebox.showinfo("Informe","Informe guardado correctamente.")
        except Exception as e: messagebox.showerror("Informe",str(e))

    def search_log(self,next_match=False):
        if not hasattr(self,'log'):return
        term=self.search_var.get()
        if not term:return
        start=self.log.index('insert') if next_match else '1.0'
        pos=self.log.search(term,start,stopindex='end',nocase=True)
        if not pos and next_match:pos=self.log.search(term,'1.0',stopindex='end',nocase=True)
        self.log.tag_remove('search_hit','1.0','end'); self.log.tag_configure('search_hit',background='lightblue')
        if pos:
            end=f"{pos}+{len(term)}c"; self.log.tag_add('search_hit',pos,end); self.log.mark_set('insert',end); self.log.see(pos)

    def infer_file_from_selection(self):
        if not hasattr(self,'log'):return None
        try:
            idx=self.log.index('sel.first')
        except tk.TclError:return None
        line=int(idx.split('.')[0]);
        for n in range(line,0,-1):
            txt=self.log.get(f'{n}.0',f'{n}.end')
            m=re.search(r'Procesando (?:factura|visita|partes):\s*(.+)$',txt.strip())
            if m:return m.group(1).strip()
        return None

    def on_log_key(self,event):
        if event.char!='º':return
        try:
            selected=self.log.get('sel.first','sel.last').strip()
        except tk.TclError:return
        if not selected:return
        self.log.tag_add('error_selection','sel.first','sel.last')
        original=self.infer_file_from_selection()
        if not original:
            messagebox.showwarning('Error de revisión','No he podido identificar el PDF a partir del fragmento seleccionado.'); return
        comment=simpledialog.askstring('Marcar error','Error seleccionado:\n'+selected+'\n\nEscribe el comentario:',parent=self.root)
        if comment is None:return
        self.handle_review_error(original,selected,comment)
        return 'break'

    def handle_review_error(self,original_name,error_text,comment):
        folder=self.cfg.get('watch_folder',''); pdf_path=os.path.join(folder,original_name)
        if not os.path.isfile(pdf_path):
            rec=find_movement_record(self.cfg.get('movimientos_folder',''),original_name)
            if rec and rec.get('route'): pdf_path=os.path.join(rec['route'],rec.get('after') or original_name)
        if not os.path.isfile(pdf_path):
            messagebox.showerror('Revisión','No encuentro el PDF original ni su copia en la ruta registrada.'); return
        try:
            read_text=extract_text_for_review_once(pdf_path,self.cfg.get('ocr',True))
            communities,_=read_communities(self.cfg.get('excel_file',''),self.cfg.get('sheet',''),self.cfg.get('community_columns','B'),self.cfg.get('id_columns','A'),self.cfg.get('route_column',''),self.cfg.get('alias_id_column','R'),self.cfg.get('alias_name_column','S'))
            community,status,matches=detect_community(read_text,communities)
            rec=find_movement_record(self.cfg.get('movimientos_folder',''),original_name)
            outdir=self.cfg.get('errores_revision_folder','') or os.path.join(self.cfg.get('movimientos_folder','') or os.path.dirname(pdf_path),'Errores revisados')
            os.makedirs(outdir,exist_ok=True); stamp=datetime.now().strftime('%Y%m%d_%H%M%S'); outpath=os.path.join(outdir,f"ERROR_{stamp}_{os.path.basename(original_name)}")
            review_metadata={
                'original': original_name,
                'community': (community.get('community','') if community else (' / '.join([m.get('community','') for m in matches]) if matches else 'No determinada')),
                'id': (community.get('id_fmt','') if community else (' / '.join([m.get('id_fmt','') for m in matches]) if matches else 'No determinado')),
                'route': (rec.get('route','') if rec else ''),
                'after': (rec.get('after','') if rec else ''),
                'error': error_text,
                'comment': comment,
            }
            redacted=create_redacted_review_copy(pdf_path,outpath,read_text,review_metadata)
            report=os.path.splitext(outpath)[0]+'.docx'; doc=Document(); doc.add_heading('Informe de error de revisión',level=1)
            doc.add_paragraph('Archivo original: '+original_name); doc.add_paragraph('Error señalado: '+error_text); doc.add_paragraph('Comentario: '+comment)
            if community:
                doc.add_paragraph(f"Comunidad detectada: {community.get('community','')}")
                doc.add_paragraph(f"ID: {community.get('id_fmt','') or 'no determinado'}")
            elif matches:
                doc.add_paragraph('Comunidades detectadas: ' + ' / '.join([m.get('community','') for m in matches]))
                doc.add_paragraph('IDs: ' + ' / '.join([m.get('id_fmt','') or 'no determinado' for m in matches]))
            else:
                doc.add_paragraph('Comunidad detectada: no determinada')
            if rec:
                doc.add_paragraph('Ruta a la que se movió: '+rec.get('route','')); doc.add_paragraph('Nombre guardado: '+rec.get('after','')); doc.add_paragraph('Word de movimientos: '+rec.get('word',''))
            doc.add_paragraph('Copia revisada: '+outpath); doc.add_paragraph('Redacción CIF/NIF aplicada: '+('sí' if redacted else 'no pudo aplicarse geométricamente'))
            doc.save(report)
            self.error_review_records.append({'original':original_name,'error':error_text,'comment':comment,'review_pdf':outpath,'report':report})
            self.log_msg(f"⚠ ERROR REVISADO: {original_name} | {error_text} | {comment} | copia: {outpath}")
            messagebox.showinfo('Error registrado',f'Se ha creado la copia revisada:\n{outpath}\n\nY el informe:\n{report}')
        except Exception as e: messagebox.showerror('Revisión',str(e))

if __name__ == "__main__":
    root=tk.Tk(); App(root); root.mainloop()