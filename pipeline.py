import json
from converters.JSON_into_LaTeX_agent import generate_latex_with_retry
from converters.pptx_into_JSON import convert_pptx_to_json
from converters.raw_JSON_into_cleaned_JSON import clean_and_map_media_elements, save_final_json_for_review
from extracter.media_from_pptx import extract_media_from_pptx
from extracter.metadata_from_pptx import extract_metadata_from_pptx
from utils import analyze_layout_structure, prepare_initial_prompt
from collections import defaultdict
from utils import (
    compile_tex_to_pdf, 
    find_prioritized_media_references, 
    prioritize_and_clean_media_map, 
    save_latex_to_dir, 
    BLUE, GREEN, YELLOW, RESET
)
LAYOUT_DATA_STORAGE = {}
async def step_extract_structure(config):
    print(f"{BLUE}Step 1/5: Extracting structure from {config.PPTX_INPUT}...{RESET}")
    await convert_pptx_to_json(
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.OUTPUT_DIR)
    )

def step_extract_media(config):
    print("{BLUE}Step 2/5: Extracting media (Recursive)...")
    layout_data = extract_media_from_pptx(
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.MEDIA_OUTPUT_DIR)
    )
    # Store globally or return for next step
    config.LAYOUT_DATA_BY_SLIDE = layout_data
    return layout_data

def step_process_and_optimize_data(config):
    print(f"Step 3/5: Processing Data (Mapping Media & Optimizing Layout)...")

    # 1. ROHDATEN LADEN (Hier stecken die BBoxen drin!)
    with open(config.RAW_JSON_INPUT, 'r', encoding='utf-8') as f:
        raw_docling_json = json.load(f)

    # 2. STANDARD CLEANING (Deine existierende Funktion)
    # Erstellt die Slides mit Titel, Header, Footer und einfachem Text
    media_geometry_map = getattr(config, 'LAYOUT_DATA_BY_SLIDE', {})
    global_meta = extract_metadata_from_pptx(str(config.PPTX_INPUT))
    
    merged_data = clean_and_map_media_elements(
        docling_data=raw_docling_json, 
        media_geometry_map=media_geometry_map
    )

   # --- 3. THE MISSING LINK (DEEP SEARCH VERSION) ---
    raw_items_by_page = defaultdict(list)
    
    # Helper: Rekursive Suche nach Items mit 'prov' (Egal wie tief verschachtelt)
    def find_items_recursive(data):
        items = []
        if isinstance(data, dict):
            # Check: Ist das hier ein Item mit Koordinaten?
            if 'prov' in data and isinstance(data['prov'], list) and len(data['prov']) > 0:
                # Optional: Text-Check (leere Items ignorieren)
                if 'text' in data or 'label' in data: 
                    items.append(data)
            
            # Weiter suchen in allen Values
            for k, v in data.items():
                items.extend(find_items_recursive(v))
                
        elif isinstance(data, list):
            # Weiter suchen in allen Elementen
            for element in data:
                items.extend(find_items_recursive(element))
        return items

    # A. Suche starten (Findet Texte UND Tabellen)
    all_source_items = find_items_recursive(raw_docling_json)
    
    # B. Mapping erstellen
    count_found = 0
    for item in all_source_items:
        page_num = item['prov'][0]['page_no']
        raw_items_by_page[page_num].append(item)
        count_found += 1

    print(f"   -> Debug: Found {count_found} raw items via Deep Search.")

    # C. In Slides injizieren
    raw_pages = raw_docling_json.get('pages', {})
    
    for slide in merged_data["slides"]:
        s_num = slide.get('slide_number')
        
        # 1. Text Rohdaten injizieren
        raw_data = raw_items_by_page.get(s_num) or raw_items_by_page.get(str(s_num)) or raw_items_by_page.get(int(s_num) if str(s_num).isdigit() else -1)
        slide['text_content_raw'] = raw_data if raw_data else []
        
        # 2. NEU: Echte Seitenmaße injizieren!
        # Docling nutzt String-Keys "1", "2" für pages
        page_info = raw_pages.get(str(s_num)) or raw_pages.get(s_num)
        if page_info and 'size' in page_info:
            slide['page_dimensions'] = page_info['size'] # {width: ..., height: ...}
        else:
            slide['page_dimensions'] = None
    # --------------------------------------------

    # 4. LAYOUT OPTIMIERUNG
    optimized_slides = []
    
    for i, slide in enumerate(merged_data["slides"]):
        # Jetzt findet die Engine die Daten in 'text_content_raw'!
        layout_instruction = analyze_layout_structure(slide)
        
        slide['llm_layout_instruction'] = layout_instruction
        if 'text_content' in slide:
            del slide['text_content']
        if 'text_content_raw' in slide:
           del slide['text_content_raw']
        # WICHTIG: Wir behalten 'text_content' als Backup, löschen es NICHT.
        optimized_slides.append(slide)

    # 5. FINALES JSON BAUEN
    final_json = {
        "presentation_meta": {
            "file_name": config.PPTX_INPUT.name,
            "detected_title": global_meta["title"],
            "detected_author": global_meta["author"],
            "global_header_text": global_meta.get("global_header_text", ""), 
            "global_footer_text": global_meta.get("global_footer_text", ""),
            "presentation_date": global_meta.get("classified_date", "")
        },
        "slides": optimized_slides
    }

    # 6. SPEICHERN
    save_final_json_for_review(final_json, str(config.CLEANED_JSON_OUTPUT))
    
    # Debug Check
    if optimized_slides:
        raw_count = len(optimized_slides[0].get('text_content_raw', []))
        print(f"   -> Data prepared. Slide 1 has {raw_count} raw items with coordinates.")
        
    return final_json

def step_generate_latex(config):
    print(f"{BLUE}Step 4/5: Generating LaTeX code with Auto-Correction...")
    
    # 1. Prepare the prompt by reading the YAML file + JSON data
    full_prompt = prepare_initial_prompt(config)
    
    # 2. Run the Generation Loop
    final_latex = generate_latex_with_retry(config, full_prompt)
    
    return final_latex

def step_save_and_compile(config, latex_code):
    print(f"{BLUE}Step 5/5: Saving and Compiling...{RESET}")
    
    saved_path = save_latex_to_dir(
        latex_code=latex_code,
        output_dir=config.RESULTS_DIR,
        filename=config.TEX_FILENAME
    )
    print(f"LaTeX saved to: {saved_path}")

    success = compile_tex_to_pdf(
        tex_file_path=saved_path, 
        output_dir=config.RESULTS_DIR
    )
    return success
