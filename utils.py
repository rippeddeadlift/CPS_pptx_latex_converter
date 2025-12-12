from ast import List, Set
import collections
import json
import os
import re
import zipfile
from pptx import Presentation
from pptx.exc import PackageNotFoundError 
from pptx.enum.shapes import MSO_SHAPE_TYPE
import sys
from pathlib import Path
import subprocess
from unstructured_client import Dict, Tuple, Union
import json
import os
RESET = "\033[0m"
RED = "\033[31m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
BLUE = "\033[34m"

def check_pptx(file_path):
    """
    Tries to open a PowerPoint file.
    If successful, prints 'hello world'.
    """
    try:
        presentation = Presentation(file_path)
        
        print(f"{GREEN}hello world{RESET}")
        
    except PackageNotFoundError:
        print(f"{RED}Error: The file '{file_path}' is not a valid PowerPoint file or was not found.{RESET}")
    except Exception as e:
        print(f"{RED}An error occurred: {e}{RESET}")


def read_pptx_content(file_path):
    """
    Reads the content from a PowerPoint file.
    """
    try:
        presentation = Presentation(file_path)
        print(f"{BLUE}--- Content of '{file_path}' ---{RESET}")

        for i, slide in enumerate(presentation.slides):
            print(f"\n{BLUE}--- Slide {i + 1} ---{RESET}")
            
            for shape in slide.shapes:
                
                if not shape.has_text_frame:
                    continue
                
                text_frame = shape.text_frame
                
                for paragraph in text_frame.paragraphs:
                    
                    full_text = ""
                    for run in paragraph.runs:
                        full_text += run.text
                    
                    if full_text.strip(): 
                        print(full_text)

        print(f"\n{BLUE}--- End of Presentation ---{RESET}")
        
    except Exception as e:
        print(f"{RED}Error reading the file: {e}{RESET}")

def analyze_pptx_slides(file_path):
    """
    Analyzes a PowerPoint file slide by slide.
    Detects text, videos, and counts the used AutoShapes.
    """
    try:
        presentation = Presentation(file_path)
        print(f"{BLUE}--- Analysis of '{file_path}' ---{RESET}")
        analyzed_pptx_data = ""
        
        for i, slide in enumerate(presentation.slides):
            print(f"\n{BLUE}--- Slide {i + 1} ---{RESET}")
            
            shape_counts = collections.defaultdict(int)
            video_found = False
            text_items = []

            for shape in slide.shapes:
                
                if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                    video_found = True
                
                elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    shape_name = shape.auto_shape_type.name
                    shape_counts[shape_name] += 1
                
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    shape_counts["PICTURE (Image)"] += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    shape_counts["TABLE"] += 1

                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        full_text = ""
                        for run in paragraph.runs:
                            full_text += run.text
                        
                        if full_text.strip(): 
                            text_items.append(full_text.strip())

            if video_found:
                analyzed_pptx_data += "video\n"
            
            if shape_counts:
                analyzed_pptx_data += "--- Detected Shapes ---\n"
                for shape_name, count in shape_counts.items():
                    analyzed_pptx_data += f"{shape_name}: {count}\n"
            
            if text_items:
                analyzed_pptx_data += "--- Text Content ---\n"
                for text in sorted(list(set(text_items))): 
                    analyzed_pptx_data += text + "\n"

        return analyzed_pptx_data

    except Exception as e:
        print(f"{RED}An error occurred: {e}{RESET}")
        return "" 

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


def save_latex_to_dir(latex_code: str, output_dir: Path, filename: str):
    """
    Saves the LaTeX code to a output directory.
    """
    try:
        index = output_dir.name
        
        new_filename_string = f"{filename}_{index}.tex"
        final_file_path = output_dir / new_filename_string
        
        final_file_path.write_text(latex_code, encoding="utf-8")
        
        print("="*50)
        print(f"{GREEN}LaTeX file successfully saved!{RESET}")
        print(f"{GREEN}==> {final_file_path}{RESET}")
        print("="*50)
        
        return final_file_path 

    except IOError as e:
        print(f"{RED}ERROR: Could not write file: {final_file_path}{RESET}")
        print(f"{RED}Details: {e}{RESET}")
        sys.exit(1)



def get_slide_number(filename: str) -> Union[int, None]:
    """Extracts the slide number from the slideX.xml.rels filename."""
    match = re.search(r'slide(\d+)\.xml\.rels', filename)
    if match:
        return int(match.group(1))
    return None

def find_prioritized_media_references(pptx_path: str) -> Dict[int, Dict[str, str]]:
    """
    Analyzes .rels files to determine prioritized media references (Video > Image)
    per slide.
    
    Returns: {SlideNumber: {rId: Filename}} - Contains only the prioritized selection.
    """
    
    # 1. Regex to extract rId and Target (corrected)
    # Group 1: rId, Group 2: Target path (e.g., ../media/media3.mp4)
    rel_pattern = re.compile(r'<Relationship Id="(rId\d+)" .*? Target="([^"]+)"')
    temp_media_storage: Dict[int, List[Tuple[str, str, str]]] = {}
    
    try:
        with zipfile.ZipFile(pptx_path, 'r') as presentation_zip:
            for name in presentation_zip.namelist():
                
                if not (name.startswith('ppt/slides/_rels/slide') and name.endswith('.rels')):
                    continue
                    
                slide_num = get_slide_number(name)
                if slide_num is None:
                    continue
                        
                temp_media_storage[slide_num] = []
                
                with presentation_zip.open(name) as rels_file:
                    rels_content = rels_file.read().decode('utf-8')
                            
                    for match in rel_pattern.finditer(rels_content):
                        r_id = match.group(1)
                        target_path = match.group(2)
                        
                        # We only filter for media files in the media/ folder
                        if '/media/' in target_path.lower():
                            filename = os.path.basename(target_path)
                            
                            media_type = 'image'
                            if filename.lower().endswith(('.mp4', '.mov', '.avi')):
                                media_type = 'video'
                            
                            temp_media_storage[slide_num].append((r_id, filename, media_type))
                            
    except Exception as e:
        print(f"{RED} ERROR parsing PPTX relations: {e}{RESET}")
        return {}

    final_media_references: Dict[int, Dict[str, str]] = {}
    
    for slide_num, refs in temp_media_storage.items():
        final_media_references[slide_num] = {}
        for r_id, filename, media_type in refs:
            final_media_references[slide_num][r_id] = filename

    return final_media_references


# Define file types for prioritization
VIDEO_EXTENSIONS = ('.mp4', '.mov', '.avi', '.webm', '.bin') 
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.gif')

def prioritize_and_clean_media_map(
    raw_media_map: Dict[int, Dict[str, str]]
) -> Dict[int, Dict[str, str]]:
    """
    Cleans the collected mapping: If a video (MP4/BIN) exists on the slide,
    all images (PNG/GIF) are eliminated. Redundant rIds are consolidated.
    
    Args:
        raw_media_map: {SlideNumber: {rId: Filename}} - The raw mapping from .rels analysis.
        
    Returns: {SlideNumber: {rId: Filename}} - Cleaned and prioritized.
    """
    
    final_media_references: Dict[int, Dict[str, str]] = {}
    
    for slide_num, refs in raw_media_map.items():
        
        videos = []
        images = []
        
        for r_id, filename in refs.items():
            if filename.lower().endswith(VIDEO_EXTENSIONS):
                videos.append((r_id, filename))
            elif filename.lower().endswith(IMAGE_EXTENSIONS):
                images.append((r_id, filename))
        
        current_slide_results = {}
        
        if not videos:
            for r_id, filename in images:
                current_slide_results[r_id] = filename
        else:
            used_images = set() 
            processed_video_files = set()
            for vid_rid, vid_name in videos:
                if vid_name in processed_video_files:
                    continue 
                processed_video_files.add(vid_name)
                poster_name = None
                base_vid = os.path.splitext(vid_name)[0]
                
                for img_rid, img_name in images:
                    if os.path.splitext(img_name)[0] == base_vid:
                        poster_name = img_name
                        used_images.add(img_rid)
                        break
                
                if not poster_name and len(images) == 1:
                    poster_name = images[0][1]
                    used_images.add(images[0][0])
                
                if poster_name:
                    current_slide_results[vid_rid] = f"{vid_name}|{poster_name}"
                else:
                    current_slide_results[vid_rid] = vid_name
            

        final_media_references[slide_num] = current_slide_results
        
    return final_media_references

def compile_tex_to_pdf(tex_file_path, output_dir):
    tex_path = Path(tex_file_path).resolve()
    out_dir = Path(output_dir).resolve()

    command = [
        "pdflatex",
        "-interaction=nonstopmode",
        f"-output-directory={str(out_dir)}",
        str(tex_path)
    ]

    try:
        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
            text=True
        )
        print(f"SUCCESS: PDF created at {out_dir / tex_path.stem}.pdf")
        return True

    except subprocess.CalledProcessError as e:
        print("ERROR: PDF compilation failed.")
        print("--- LaTeX Error Log (Tail) ---")
        if e.stdout:
            print("\n".join(e.stdout.splitlines()[-20:]))
        return False
    except FileNotFoundError:
        print("ERROR: pdflatex not found. Ensure a LaTeX distribution is installed.")
        return False
    import json

def check_media_completeness(json_path, latex_code):
    """
    Scans the source JSON and checks if all media files (images/videos)
    are present in the generated LaTeX code.
    Returns a list of missing files or error messages.
    """
    missing_items = []
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for slide in data.get('slides', []):
        for element in slide.get('elements', []):
            content = element.get('content', '')
            
            if isinstance(content, str) and ('extracted_media' in content or '.mp4' in content or '.png' in content):
                
                files = content.split('|')
                
                for file_ref in files:
                    clean_ref = file_ref.strip()
                    if not clean_ref: 
                        continue
                        
                    if clean_ref not in latex_code:
                        missing_items.append(f"MISSING MEDIA: {clean_ref} (Slide {slide.get('slide_number')})")

                    if '.mp4' in clean_ref or '.avi' in clean_ref:
                        if '\\includemedia' not in latex_code:
                            missing_items.append(f"MISSING VIDEO PLAYER: {clean_ref} detected, but \\includemedia command is missing.")

    return missing_items


def _get_data_root(docling_data):
    """Entpackt structure_analysis falls nötig"""
    if isinstance(docling_data, dict) and "structure_analysis" in docling_data:
        return docling_data["structure_analysis"]
    return docling_data

def _get_page_dimensions(data_root):
    """Sammelt Höhen/Breiten aller Seiten"""
    dims = {}
    if "pages" in data_root and isinstance(data_root["pages"], dict):
        for pid, pdata in data_root["pages"].items():
            if "size" in pdata:
                dims[int(pid)] = {
                    "width": pdata["size"].get("width", 0),
                    "height": pdata["size"].get("height", 0)
                }
    return dims

def _get_text_items(data_root):
    """Findet die Liste der Texte (egal ob 'texts' oder 'main_text')"""
    if "texts" in data_root: return data_root["texts"]
    if "main_text" in data_root: return data_root["main_text"]
    return []

def _extract_prov_data(item):
    """Holt Seite, BBox und Ursprung aus einem Text-Item"""
    page_no = 1
    bbox = None
    origin = "TOPLEFT" # Default
    
    if "prov" in item and isinstance(item["prov"], list) and len(item["prov"]) > 0:
        prov = item["prov"][0]
        page_no = prov.get("page_no", 1)
        bbox = prov.get("bbox")
        
        if isinstance(bbox, dict):
            origin = bbox.get("coord_origin", "TOPLEFT")
            
    return page_no, bbox, origin

def _determine_zone(bbox, origin, page_height):
    """
    Entscheidet: Header, Title, Content oder Footer?
    
    NEUE LOGIK (Basiert auf deinen Debug-Daten):
    0.00 - 0.10 = HEADER
    0.10 - 0.25 = TITLE
    > 0.89      = FOOTER  (Angehoben von 0.85, da Content bei 0.856 liegt!)
    Rest        = CONTENT
    """
    if not bbox or page_height == 0:
        return "content"

    # Y-Koordinate holen
    current_y = 0
    if isinstance(bbox, dict):
        current_y = bbox.get("t", bbox.get("y", 0))
    elif isinstance(bbox, list):
        current_y = bbox[1]

    # Normalisieren
    if current_y > 10000 and page_height < 10000:
         page_height = 6858000 
    
    relative_y = current_y / page_height

    # --- ZONING ---
    if relative_y < 0.10: 
        return "header"
        
    if relative_y < 0.25: 
        return "title"
        
    if relative_y > 0.9: 
        return "footer"
        
    return "content"

def _assemble_final_json(slides_data, media_map):
    cleaned_slides = []
    max_slide = max(max(slides_data.keys()) if slides_data else 0, len(media_map))
    
    for i in range(max_slide):
        slide_num = i + 1
        data = slides_data[slide_num]
        
        # Alle Felder joinen
        header_text = " ".join(data["header"])  
        title_text = " ".join(data["title"])
        footer_text = " | ".join(data["footer"])
        content_text = "\n".join(data["content"])
        
        # Notfall-Logik anpassen:
        # Wenn wir keinen Title haben, aber Header, könnte der Header der Title sein?
        # Nein, wir lassen das strikt getrennt für das LLM.
        
        slide_obj = {
            "slide_number": slide_num,
            "layout_analysis": {
                "detected_header": header_text, # NEU
                "detected_title": title_text,
                "detected_footer": footer_text
            },
            "text_content": content_text,
            "media_files": media_map.get(i, [])
        }
        cleaned_slides.append(slide_obj)
        
    return {"slides": cleaned_slides}
# Neue Funktion für deine Pipeline / main.py

def prepare_initial_prompt(config):
    import json
    
    # 1. Lade YAML Regeln (Das "Handbuch")
    rules_path = config.BASE_OUTPUT_PATH.parent / config.RULES_FILE 
    with open(rules_path, 'r', encoding='utf-8') as f:
        rules_content = f.read()

    # 2. Lade JSON Daten (Das "Material")
    with open(config.CLEANED_JSON_OUTPUT, 'r', encoding='utf-8') as f:
        presentation_data = json.load(f)

    # 3. Baue den Prompt zusammen
    prompt = f"""
    ### INSTRUCTION MANUAL (RULES) ###
    Apply the following YAML rules strictly to the data below.
    Pay specific attention to the 'layout_logic' for image placement.

    {rules_content}

    ### SOURCE DATA (JSON) ###
    Here is the content you must convert. 
    ID: {presentation_data.get('presentation_meta', {}).get('detected_title', 'Unknown')}

    {json.dumps(presentation_data, indent=2, ensure_ascii=False)}

    ### COMMAND ###
    Generate the LaTeX code now. Adhere to the 'slide_rules' defined above.
    """
    return prompt



def get_single_image_position(geometry):
    """
    Berechnet die Position für EIN einzelnes Bild.
    """
    if not geometry: return "unknown"
    
    x, y, w, h = geometry
    center_x = x + (w / 2)
    
    # 1. Full Width Check
    if w > 0.80:
        if y < 0.35: return "top" 
        if y > 0.7: return "bottom"
        return "center"

    # 2. Spalten Check (Horizontal hat Priorität!)
    if center_x > 0.55: return "right"
    if center_x < 0.45: return "left"

    # 3. Vertical Check (für mittige Bilder)
    # Lockerer Grenzwert: Titel brauchen Platz, also ist y=0.3 immer noch "oben"
    if y < 0.35: 
        return "top"
        
    if y > 0.75: 
        return "bottom"
        
    return "center"

def analyze_layout_structure(slide_data):
    """
    Hauptfunktion: Orchestriert die Layout-Analyse.
    """
    raw_texts = slide_data.get('text_content_raw', []) 
    media_files = slide_data.get('media_files', [])
    
    # Metadaten holen
    layout_analysis = slide_data.get('layout_analysis', {})
    detected_title = layout_analysis.get('detected_title', "")
    detected_footer = layout_analysis.get('detected_footer', "")
    
    # NEU: Echte Maße holen
    page_dims = slide_data.get('page_dimensions') 

    # SCHRITT 1: Bilder analysieren
    _calculate_image_positions(media_files)
    _apply_image_gravity(media_files)

    # SCHRITT 2: Text vorbereiten (mit echten Maßen!)
    if raw_texts:
        normalized_items = _normalize_raw_items(raw_texts, page_dims) # <--- HIER ÜBERGEBEN
    else:
        simple_text = slide_data.get('text_content', "")
        fallback_list = simple_text.split('\n') if isinstance(simple_text, str) else simple_text
        normalized_items = [{"text": t, "bbox": None, "type": "text"} for t in fallback_list if t]

    # ... (Rest der Funktion bleibt IDENTISCH wie vorher: Mischen, Verteilen, Return) ...
    mixed_items = _create_mixed_items(normalized_items, media_files)
    zones, strategy = _distribute_to_zones(mixed_items, detected_title, detected_footer)

    return {
        "strategy": strategy,
        "zones": zones
    }

# --- Helper 1: Bild-Positionen berechnen ---
def _calculate_image_positions(media_files):
    """Berechnet die initiale Position (Left/Right/Top) für jedes Bild."""
    for img in media_files:
        if 'geometry' in img:
            img['layout_pos'] = get_single_image_position(img['geometry'])
        else:
            img['layout_pos'] = "unknown"

# --- Helper 2: Gravity Logic (Gruppenzwang) ---
def _apply_image_gravity(media_files):
    """Zwingt Bilder auf eine Seite, wenn eine Dominanz erkennbar ist."""
    left_imgs = [img for img in media_files if img['layout_pos'] == 'left']
    right_imgs = [img for img in media_files if img['layout_pos'] == 'right']
    
    # 1. Konfliktlösung (Bilder links UND rechts)
    if left_imgs and right_imgs:
        target_side = "right" # Default Winner
        if len(left_imgs) > len(right_imgs) + 1: 
            target_side = "left"
        
        for img in left_imgs + right_imgs:
            img['layout_pos'] = target_side
            
    # 2. "Center"-Bilder in die Spalte ziehen, wenn Spalten existieren
    if right_imgs or (left_imgs and target_side == "left"):
        active_side = "left" if (left_imgs and target_side == "left") else "right"
        for img in media_files:
            if img['layout_pos'] == 'center':
                # Nur Content-Bilder (nicht Header) verschieben
                y = img.get('geometry', [0,0])[1]
                if 0.2 < y < 0.8:
                    img['layout_pos'] = active_side

# --- Helper 3: Text Normalisierung ---
def _normalize_raw_items(raw_texts, page_dims=None):
    """Wandelt Docling-Rohdaten in normalisierte Items um."""
    
    # 1. Page Dimensions bestimmen (Jetzt exakt!)
    if page_dims:
        page_w = page_dims.get('width', 0)
        page_h = page_dims.get('height', 0)
    else:
        page_w, page_h = 0, 0

    # Fallback: Wenn keine echten Maße da sind, müssen wir doch raten (Sicherheit)
    if not page_w or not page_h:
        max_r, max_t = 0, 0
        for item in raw_texts:
            if 'prov' in item and len(item['prov']) > 0:
                b = item['prov'][0]['bbox']
                max_r = max(max_r, b['r'])
                max_t = max(max_t, b['t'])
        page_w = max_r * 1.05 if max_r > 0 else 1000
        page_h = max_t * 1.05 if max_t > 0 else 1000

    normalized = []
    for item in raw_texts:
        is_table = item.get('label') == 'table' or item.get('is_table')
        
        if is_table:
            try: text_content = reconstruct_docling_table(item)
            except: text_content = "DETECTED_TABLE_START"
        else:
            text_content = item.get('text', '').strip()
            
        if not text_content or not item.get('prov'): continue
        
        raw = item['prov'][0]['bbox']
        
        # Berechnung mit echten Werten
        x = raw['l'] / page_w
        y = (page_h - raw['t']) / page_h 
        w = (raw['r'] - raw['l']) / page_w
        h = (raw['t'] - raw['b']) / page_h
        
        normalized.append({
            "text": text_content,
            "bbox": [round(x,3), round(y,3), round(w,3), round(h,3)],
            "type": "table" if is_table else "text"
        })
    return normalized

# --- Helper 4: Mischen & Sortieren ---
def _create_mixed_items(text_items, media_files):
    """Fügt Text und Bild-Platzhalter zusammen und sortiert sie."""
    mixed = []
    
    # Texte
    for item in text_items:
        item["type"] = item.get("type", "text")
        mixed.append(item)
        
    # Bilder
    for img in media_files:
        if 'geometry' in img:
            ix, iy, iw, ih = img['geometry']
            mixed.append({
                "text": f"[[IMAGE: {img.get('filename')}]]", 
                "bbox": [ix, iy, iw, ih],
                "type": "image",
                "forced_pos": img['layout_pos']
            })
            
    # Sortieren: Y (Zeile) vor X (Spalte)
    # Sicheres Sortieren auch wenn bbox fehlen sollte
    def sort_key(k):
        if k.get('bbox'):
            return (round(k['bbox'][1], 2), k['bbox'][0])
        return (999, 999)
        
    mixed.sort(key=sort_key)
    return mixed

# --- Helper 5: Zoning Logic (Der Kern) ---
# --- Helper 5: Zoning Logic (Fix für Tabellen) ---
# --- Helper 5: Zoning Logic (Final Fix) ---
def _distribute_to_zones(mixed_items, detected_title="", detected_footer=""):
    """Verteilt die Items in die 4 Zonen."""
    zones = {"top_content": [], "left_column": [], "right_column": [], "flow_content": []}
    
    # 1. VISUELLE BARRIERE
    # Wo beginnt der "Body" (Bilder/Tabellen)?
    top_barrier_y = 1.0 
    has_layout_elements = False
    
    for item in mixed_items:
        if item['type'] in ['image', 'table']:
            if item.get('bbox'):
                y = item['bbox'][1]
                # Wir ignorieren nur Elemente ganz oben an der Kante
                if y > 0.08: 
                    if y < top_barrier_y: top_barrier_y = y
                    has_layout_elements = True
    
    if not has_layout_elements:
        top_barrier_y = 0.25 

    # 2. FOOTER PARTS VORBEREITEN
    # Wir zerlegen den Footer String ("Datum | Seite | Quelle") in Einzelteile
    footer_parts = []
    if detected_footer:
        # Split am Pipe '|' und bereinigen
        parts = [p.strip().lower() for p in detected_footer.split('|')]
        # Nur Teile aufnehmen, die nicht leer sind
        footer_parts = [p for p in parts if len(p) > 0]

    # Regex für typische Footer-Elemente (als Backup)
    footer_regex = re.compile(r'^(\d{1,2}\.\d{1,2}\.\d{2,4})$|^(seite\s*\d+)$|^(page\s*\d+)$', re.IGNORECASE)

    # 3. VERTEILUNG
    has_columns = False
    
    def clean_str(s): return " ".join(str(s).lower().split())
    clean_title = clean_str(detected_title)
    
    for item in mixed_items:
        text = item['text']
        bbox = item.get('bbox')
        item_type = item.get('type')
        forced_pos = item.get('forced_pos')
        clean_text = clean_str(text)
        
        # --- A. FILTER (Aggressiv) ---
        
        # 1. Titel-Check
        if item_type == "text" and clean_title and clean_text == clean_title:
            continue
            
        # 2. Footer-Check (Smart)
        is_footer = False
        # a) Exakter Match mit einem Footer-Teil (z.B. "seite 2" == "seite 2")
        if clean_text in footer_parts: is_footer = True
        
        # b) Regex Match (für Datum/Seitenzahl egal wo sie stehen)
        if footer_regex.match(text.strip()): is_footer = True
        
        # c) Position Check (ganz unten ist fast immer Footer)
        if bbox and bbox[1] > 0.94: is_footer = True
        
        if is_footer: continue

        # --- B. FALLBACK (Ohne BBox) ---
        if not bbox:
            if forced_pos == "left": zones["left_column"].append(text); has_columns=True
            elif forced_pos == "right": zones["right_column"].append(text); has_columns=True
            elif forced_pos == "top": zones["top_content"].append(text)
            else: zones["flow_content"].append(text)
            continue

        x, y, w, h = bbox
        
        # --- C. IMAGE FORCING ---
        if forced_pos:
            if forced_pos == "top": zones["top_content"].append(text)
            elif forced_pos == "left": zones["left_column"].append(text); has_columns=True
            elif forced_pos == "right": zones["right_column"].append(text); has_columns=True
            elif forced_pos == "center":
                # Nur nach oben, wenn es über der Barriere ist
                if y < top_barrier_y: zones["top_content"].append(text)
                else: zones["flow_content"].append(text)
            else: zones["flow_content"].append(text)
            continue

        # --- D. TOP CONTENT ENTSCHEIDUNG ---
        is_structure_element = (item_type == 'table')
        
        # FIX: Tabellen dürfen NICHT nach oben, außer sie kleben an der Decke (< 5%)
        should_go_top = False
        
        if is_structure_element:
             if y < 0.05: should_go_top = True
        else:
            # Text darf nach oben, wenn er breit ist ODER visuell über der Barriere liegt
            # Wir nutzen 0.85 als Breite
            is_wide = w > 0.85 
            # Puffer: Muss 2% über dem Start der Spalten liegen
            is_above = y < (top_barrier_y - 0.02)
            should_go_top = is_wide or is_above

        if should_go_top:
            zones["top_content"].append(text)
            continue

        # --- E. SPALTEN AUFTEILUNG ---
        center_x = x + (w/2)
        if center_x < 0.50:
            zones["left_column"].append(text)
        else:
            zones["right_column"].append(text)
        
        has_columns = True

    strategy = "columns" if has_columns else "standard_flow"
    return zones, strategy

def reconstruct_docling_table(table_item):
    """
    Wandelt ein Docling Table-Objekt (Grid) zurück in unseren String-Format.
    """
    grid = table_item.get('data', {}).get('grid', [])
    if not grid:
        return "DETECTED_TABLE_START\n(Empty Table)\nDETECTED_TABLE_END"
        
    lines = ["DETECTED_TABLE_START"]
    
    # Header & Rows generieren (Simple CSV Style für das LLM)
    for row in grid:
        # Jede Zelle hat 'text'. Wir joinen sie mit ' | '
        row_text = " | ".join([cell.get('text', '').strip() for cell in row])
        lines.append(row_text)
        
    lines.append("DETECTED_TABLE_END")
    return "\n".join(lines)