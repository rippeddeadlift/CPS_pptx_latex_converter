import json
from generator import LATEX_POSTAMBLE,generate_latex_preamble
from converters.JSON_into_LaTeX_agent import generate_single_slide_latex
from converters.pptx_into_JSON import convert_pptx_to_json
from extracter.media_from_pptx import extract_media_from_pptx
from extracter.metadata import transform_docling_json_to_slides
from utils import (
    compile_tex_to_pdf, 
    extract_metadata,
    BLUE, GREEN, YELLOW, RESET,
    detect_header_candidate,
    enrich_and_group_slides,
    get_slide_dimensions,
    load_slides,
    remove_auto_header,
    sanitize_latex,
    save_json
)
LAYOUT_DATA_STORAGE = {}
async def step_extract_structure(config):
    print(f"{BLUE}Step 1/5: Extracting structure from {config.PPTX_INPUT}...{RESET}")
    

    await convert_pptx_to_json(
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.JSON_OUTPUT_DIR)  
    )

def step_extract_media(config):
    print(f"{BLUE}Step 2/5: Extracting media (Recursive)...{RESET}")
    layout_data = extract_media_from_pptx(
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.MEDIA_OUTPUT_DIR)
    )
    config.LAYOUT_DATA_BY_SLIDE = layout_data
    return layout_data

def step_process_and_optimize_data(config):
    print(f"{BLUE}Step 3/5: Process and Optimize Data...{RESET}")
    
    input_path = config.RAW_JSON_INPUT
    output_path = config.CLEANED_JSON_OUTPUT
    
    # Check, ob Step 1 erfolgreich war
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        print(f"(Did Step 1 save to the wrong folder? Checked: {input_path.parent})")
        return

    try:
        print(f"Loading raw JSON from: {input_path}")
        with open(input_path, 'r', encoding='utf-8') as f:
            raw_data = json.load(f)
            
        print("Grouping and sorting slides by layout...")
        slides_data = transform_docling_json_to_slides(raw_data)
        
        print(f"Saving {len(slides_data)} slides to: {output_path}")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(slides_data, f, indent=2, ensure_ascii=False)
            
        print(f"{GREEN}Data optimization complete.{RESET}")
        
    except Exception as e:
        print(f"Error during processing: {e}")
        import traceback
        traceback.print_exc()

def step_generate_latex(config):
    print(f"\n{BLUE}Step 4/5: step_generate_latex...{RESET}")
    # Step 1: Lade Slides und extrahiere Metadaten
    slides = load_slides(config.CLEANED_JSON_OUTPUT)
    if slides is None: return None

    meta = extract_metadata(config)

    # Step 2: Hole Slide-Dimensionen (für BoundingBox)
    slide_width, slide_height = get_slide_dimensions(config.PPTX_INPUT)

    # Step 3: Rechne Geometrie und gruppiere Elemente
    slides = enrich_and_group_slides(slides, slide_width, slide_height)

    # Step 4: Automatische Header-Erkennung und -Bereinigung
    header_text = None
    header_result = detect_header_candidate(slides)
    if header_result is not None:
        header_text, header_geometry = header_result
        #slides = remove_auto_header(slides, header_text, header_geometry)
        #print(f"[INFO] Detected and removed header: '{header_text}'")
    
    # (Optional: JSON für Debugging speichern)
    save_json(slides, config.CLEANED_JSON_OUTPUT)

    # Step 5: Generiere die LaTeX-Preamble (mit Subtitle!)
    latex_preamble_code = generate_latex_preamble(meta, header_text)

    # Step 6: Für jede Slide LaTeX generieren (slide by slide)
    slide_blocks = []
    total_slides = len(slides)
    for i, slide in enumerate(slides):
        slide_num = slide.get('slide_number', i+1)
        print(f"→ Generiere LaTeX für Slide {slide_num} ({i+1}/{total_slides}) ...")
        latex_code = generate_single_slide_latex(slide, config)
        block = f"\n% --- Slide {slide_num} ---\n{latex_code}\n"
        slide_blocks.append(block)

    # Step 7: Dokument zusammenbauen
    print("Assembling final document...")
    full_body_latex = "".join(slide_blocks)
    final_latex_document = f"{latex_preamble_code}\n{full_body_latex}\n{LATEX_POSTAMBLE}"
    return final_latex_document

def step_save_and_compile(config, latex_code):
    print(f"\n{BLUE}Step 5/5: Saving and Compiling...{RESET}")

    if not latex_code:
        print("Error: No LaTeX code to save.")
        return False
    clean_latex = sanitize_latex(latex_code)  

    output_dir = config.OUTPUT_DIR 
    output_dir.mkdir(parents=True, exist_ok=True)
    
    tex_filename = config.TEX_FILENAME + ".tex"
    tex_path = output_dir / tex_filename
    
    try:
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(clean_latex)       
        print(f"LaTeX saved to: {tex_path}")
    except Exception as e:
        print(f"Error saving .tex file: {e}")
        return False

    success = compile_tex_to_pdf(tex_filename, output_dir)
    return success