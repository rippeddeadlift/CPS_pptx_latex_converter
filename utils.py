
import subprocess
from pathlib import Path
import sys, re
from collections import Counter
from pptx import Presentation
import json
from extracter.metadata_from_pptx import get_institute_heuristic

RESET = "\033[0m"
RED = "\033[31m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
BLUE = "\033[34m"

def compile_tex_to_pdf(tex_filename, working_dir):
    """
    Kompiliert die .tex Datei.
    WICHTIG: Führt den Befehl IM working_dir aus (cwd), damit relative Bildpfade funktionieren.
    """
    print(f"⚙️ Compiling {tex_filename} in {working_dir}...")
    
    command = [
        "pdflatex",
        "-interaction=nonstopmode",
        tex_filename 
    ]

    try:
        result = subprocess.run(
            command,
            cwd=working_dir,  
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False 
        )

        if result.returncode == 0:
            pdf_name = Path(tex_filename).stem + ".pdf"
            print(f"SUCCESS: PDF generated at {working_dir / pdf_name}")
            return True
        else:
            print("ERROR: PDF compilation failed.")
            print("--- LaTeX Error Log (Last 20 lines) ---")
            lines = result.stdout.splitlines()
            print("\n".join(lines[-20:]))
            return False

    except FileNotFoundError:
        print("CRITICAL ERROR: 'pdflatex' not found.")
        print("Please install a LaTeX distribution (e.g., MiKTeX on Windows, TeX Live on Linux).")
        return False
    except Exception as e:
        print(f"Unexpected error during compilation: {e}")
        return False
    
def get_and_create_next_run_dir(base_dir: Path) -> Path:
    """
    Finds the next available indexed directory (e.g., 'Results/19')
    and creates it.
    Returns the Path to the newly created directory.
    """
    index = 1
    if not base_dir.is_dir():
        print(f"{RED}Error: The base directory '{base_dir}' does not exist.{RESET}")
        sys.exit(1) 

    while True:
        new_dir_path = base_dir / str(index) 
        
        if not new_dir_path.exists():
            break 
        
        index += 1

    try:
        new_dir_path.mkdir()
        print(f"{GREEN}Successfully created new run directory: {new_dir_path}{RESET}")
        return new_dir_path 
    except OSError as e:
        print(f"{RED}ERROR: Could not create directory: {new_dir_path}{RESET}")
        print(f"{RED}Details: {e}{RESET}")
        sys.exit(1)

def extract_metadata(config) -> dict:
    """Extrahiert Metadaten robust aus der PPTX oder nutzt Defaults."""
    try:
        print("Extracting PPTX metadata...")
        prs = Presentation(config.PPTX_INPUT)
        props = prs.core_properties
        
        title_text = props.title if props.title else config.PPTX_INPUT.stem
        author_text = props.author if props.author else "AI Converter"
        
        institute_text = props.category if props.category else ""
        
        if not institute_text:
            print("   -> No metadata 'category' found. Trying to guess from Slide Master...")
            institute_text = get_institute_heuristic(prs, title_text, author_text)
            
        if institute_text:
            institute_text = institute_text.replace('\n', r' \\ ')

        meta = {
            "title": title_text,
            "author": author_text,
            "date": props.created.strftime("%d.%m.%Y") if props.created else r"\today",
            "institute": institute_text
        }
        
        return meta

    except Exception as e:
        print(f"{YELLOW}Could not extract metadata: {e}. Using defaults.{RESET}")
        return {
            "title": "Presentation", 
            "author": "LaTeX Converter", 
            "date": r"\today", 
            "institute": ""
        }

def _calculate_geometry(bbox, page_width, page_height):
    """
    Berechnet relative LaTeX-Koordinaten (0.0-1.0).
    Korrigiert automatisch vertauschte Top/Bottom Werte.
    """
    if not bbox or page_width == 0 or page_height == 0:
        return None
    
    l = bbox.get('l', 0)
    t = bbox.get('t', 0)
    r = bbox.get('r', 0)
    b = bbox.get('b', 0)
    
    width_emu = abs(r - l)
    
    height_emu = abs(b - t)
    
    visual_top = min(t, b)
    
    rel_x = l / page_width
    rel_y = visual_top / page_height
    rel_w = width_emu / page_width
    rel_h = height_emu / page_height

    return {
        "x": round(max(0.0, min(1.0, rel_x)), 3),
        "y": round(max(0.0, min(1.0, rel_y)), 3),
        "w": round(max(0.0, min(1.0, rel_w)), 3),
        "h": round(max(0.0, min(1.0, rel_h)), 3) 
    }
def is_code_line(line):
    # Heuristik: erkenne Java/C-artige Zeilen
    code_tokens = [';', '{', '}', 'int ', 'public ', 'private ', '=', 'while ', 'if ', 'for ']
    return any(token in line for token in code_tokens)

def build_geo_dict(elements):
    geos = {}
    for i, el in enumerate(elements):
        geo = tuple(sorted(el['geometry'].items()))
        geos.setdefault(geo, []).append((i, el))
    return geos

def group_elements(elements):
    grouped = []
    used = set()
    geos = build_geo_dict(elements)
    
    for geo, group in geos.items():
        first_el = group[0][1] 
        y = first_el['geometry']['y']
        
        if y < 0.03:
            items = [(idx, el) for idx, el in group if 'text' in el and idx not in used]
            if items:
                text = "\n".join(el['text'] for _, el in items)
                grouped.append({
                    "type": "header",
                    "geometry": get_union_geometry([el for _, el in items]),
                    "text": text.strip(),
                    "fontsize": "3pt", 
                })
                for idx, _ in items: used.add(idx)
        

        elif y > 0.87:
            items = [(idx, el) for idx, el in group if 'text' in el and idx not in used]
            if items:
                text = "\n".join(el['text'] for _, el in items)
                grouped.append({
                    "type": "footer",
                    "geometry": get_union_geometry([el for _, el in items]),
                    "text": text.strip(),
                    "fontsize": "3pt",
                })
                for idx, _ in items: used.add(idx)

        sure_code_indices = sorted([
            i for i, (idx, el) in enumerate(group)
            if "text" in el and is_code_line(el['text'])
        ])
        
        if len(sure_code_indices) >= 2:
            blocks = []
            current_block = [sure_code_indices[0]]
            for i in range(1, len(sure_code_indices)):
                if sure_code_indices[i] - sure_code_indices[i-1] > 4: 
                    blocks.append(current_block)
                    current_block = []
                current_block.append(sure_code_indices[i])
            blocks.append(current_block)
            
            for blk in blocks:
                if len(blk) < 2: continue
                
                subset = group[blk[0] : blk[-1] + 1]
                code_text = "\n".join(el['text'] for idx, el in subset if 'text' in el)
                union_geo = get_union_geometry([el for idx, el in subset])
                
                grouped.append({
                    "type": "codeblock",
                    "geometry": union_geo,
                    "text": f"\\begin{{lstlisting}}[language=Java]\n{code_text}\n\\end{{lstlisting}}"
                })
                for idx, el in subset: used.add(idx)


        list_like = [(idx, el) for idx, el in group if idx not in used and (
            el['type'] == "list" or el.get("label") in ("list_item", "paragraph", "text"))]
            
        if list_like:
                group_align = "t"
                for _, el in list_like:
                    if el.get("align") == "b":
                        group_align = "b"
                        break

                items = [el['text'] for idx, el in list_like if 'text' in el]
                union_geo = get_union_geometry([el for idx, el in list_like])
                
               
                is_list = False
                
                if len(items) >= 3:
                    if len(items) > 4:
                        is_list = True
                    elif all(len(i) > 20 for i in items):
                        is_list = True

                if is_list:
                    grouped.append({
                        "type": "list",
                        "geometry": union_geo,
                        "items": items,
                        "align": group_align, 
                        "fontsize": "scriptsize"
                    })
                
                else:
                    full_text = "\n".join(items)
                    
                    font_sz = "tiny" if len(full_text) < 20 else "small"
                        
                    grouped.append({
                        "type": "text", 
                        "geometry": union_geo,
                        "text": full_text, 
                        "align": group_align, 
                        "fontsize": font_sz
                    })
                
                for idx, el in list_like: used.add(idx)

        for idx, el in group:
            if idx not in used:
                grouped.append(el)
                used.add(idx)
                
    return grouped


def detect_header_candidate(slides):
    counter = Counter()
    texts = {}
    for slide in slides:
        for el in slide['elements']:
            if el['type'] == 'text' or el.get('label') in ['paragraph', 'header', 'footer']:
                key = (el.get('text').strip(), tuple(sorted(el['geometry'].items())))
                counter[key] += 1
                texts[key] = el.get('text')
    thresh = int(0.7 * len(slides)) 
    if not counter:
        return None
    candidates = [key for key, val in counter.items() if val >= thresh]
    if not candidates:
        return None
    # Nimm den längsten Text als typischen header
    header_key = max(candidates, key=lambda k: len(k[0]))
    header_text = header_key[0]
    header_geometry = header_key[1]
    return header_text, dict(header_geometry)


def remove_auto_header(slides, header_text, header_geometry):
    new_slides = []
    for i, slide in enumerate(slides):
        # Lasse Titelseite unverändert, entferne sonst den Header
        if i == 0:
            new_slides.append(slide)
            continue
        filtered = []
        for el in slide['elements']:
            if (el.get('text', '').strip() == header_text.strip() and
                all(abs(el['geometry'][k] - header_geometry[k]) < 1e-4 for k in header_geometry)):
                continue  # Entferne diese Zeile
            filtered.append(el)
        slide['elements'] = filtered
        new_slides.append(slide)
    return new_slides



def inject_header_to_title_slide(slides, header_text):
    if not slides or not header_text:
        return slides
    title_slide = slides[0]
    # Du kannst es als zusätzlichen "subtitle" eintragen:
    title_slide['elements'].append({
        'type': 'subtitle',
        'text': header_text,
        'geometry': '<optional: gleiche wie vorher, oder speziel für Titelfolie>',
    })
    # oder im LaTeX direkt als \subtitle nutzen, oder als Text-Box
    return slides


def load_slides(json_path):
    from pathlib import Path
    if isinstance(json_path, str):
        json_path = Path(json_path)
    if not json_path.exists():
        print(f"[ERROR] JSON not found at {json_path}")
        return None
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        print(f"[ERROR] JSON file is corrupted: {e}")
        return None
    
def get_slide_dimensions(pptx_path):
    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
        return prs.slide_width, prs.slide_height
    except Exception as e:
        print(f"[WARN] Could not load PPTX dimensions: {e}")
        return 0, 0
    
def enrich_and_group_slides(slides, slide_width, slide_height):
    for slide in slides:
        elements = slide.get('elements', [])
        for el in elements:
            if 'bbox' in el:
                geo = _calculate_geometry(el['bbox'], slide_width, slide_height)
                el['geometry'] = geo
                del el['bbox']
        slide['elements'] = group_elements(elements)
    return slides

def save_json(data, path):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def sanitize_latex(llm_text):
    # Fix possible backspace or control char substitution (hex 08, 09, 0a, etc.) at start of commands
    latex = re.sub(r'([\x00-\x1F]|\/)+begin', r'\\begin', llm_text, flags=re.MULTILINE)
    latex = re.sub(r'([\x00-\x1F]|\/)+end', r'\\end', latex, flags=re.MULTILINE)

    # Fix all standalone '/item' to '\item'
    latex = re.sub(r'([\x00-\x1F]|\/)+item', r'\\item', latex, flags=re.MULTILINE)

    # As safety, replace *any* remaining '\x08' (ASCII 8) anywhere with '\\'
    latex = latex.replace('\x08', '\\')

    # Replace double-backslash back to single if LLM double-escapes
    latex = re.sub(r'\\\\begin', r'\\begin', latex)
    latex = re.sub(r'\\\\end', r'\\end', latex)
    latex = re.sub(r'\\\\item', r'\\item', latex)

    return latex


def get_union_geometry(elements):
    """
    Berechnet die umschließende Box (Union) für eine Gruppe von Elementen.
    Erwartet Elemente mit 'geometry': {'x', 'y', 'w', 'h'}.
    """
    if not elements:
        return None

    # Initialisierung mit Extremwerten
    min_x = float('inf')
    min_y = float('inf')
    max_r = float('-inf') # r = x + w (Rechter Rand)
    max_b = float('-inf') # b = y + h (Unterer Rand)

    for el in elements:
        geo = el.get('geometry', {})
        # Überspringe Elemente ohne Geometrie
        if not geo: 
            continue
            
        x = geo.get('x', 0)
        y = geo.get('y', 0)
        w = geo.get('w', 0)
        h = geo.get('h', 0)
        
        # Min/Max berechnen
        min_x = min(min_x, x)
        min_y = min(min_y, y)
        max_r = max(max_r, x + w)
        max_b = max(max_b, y + h)

    # Die neuen Dimensionen der großen Box
    new_w = max_r - min_x
    new_h = max_b - min_y

    return {
        "x": round(min_x, 3),
        "y": round(min_y, 3),
        "w": round(new_w, 3),
        "h": round(new_h, 3)
    }



def repair_latex_output(latex_code):
    """
    Repariert typische Flüchtigkeitsfehler von kleinen LLMs.
    """
    # 1. Fix: "\paper" statt "\paperheight"
    # Sucht nach \paper, dem KEIN "height" oder "width" folgt
    latex_code = re.sub(r'\\paper(?!(height|width))', r'\\paperheight', latex_code)
    
    # 2. Fix: "\paper]" (passiert oft in minipage definitionen)
    latex_code = latex_code.replace(r'\paper]', r'\paperheight]')
    
    # 3. Fix: Fehlende schließende Klammer bei minipage (Notfall-Fix)
    # Wenn eine Zeile mit \begin{minipage} anfängt, aber am Ende kaputt aussieht
    # (Das ist komplexer, aber der \paper Fix löst meistens 99% der Probleme)
    
    return latex_code