import json
import traceback
from helper.generator import LATEX_POSTAMBLE,generate_latex_preamble
from helper.agent import generate_single_slide_latex
from helper.pptx_into_JSON import convert_pptx_to_json
from helper.media_from_pptx import extract_media_from_pptx
from helper.metadata import transform_docling_json_to_slides
from helper.metadata_from_pptx import extract_metadata
from helper.utils import (
    get_text_alignment_map,
    compile_tex_to_pdf, 
    RED, BLUE, GREEN, YELLOW, RESET,
    detect_header_candidate,
    enrich_and_group_slides,
    get_slide_dimensions,
    inject_group_metrics,
    load_slides,
    sanitize_latex,
    save_json
)
LAYOUT_DATA_STORAGE = {}

async def step_extract_structure(config):
    """
    Executes the first pipeline step: structural extraction.

    Logs the initialization and asynchronously converts the source PowerPoint file 
    into a structured JSON representation within the configured output directory.
    """
    print(f"{BLUE}Step 1/5: Extracting structure from {config.PPTX_INPUT}...{RESET}")
    
    await convert_pptx_to_json(
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.JSON_OUTPUT_DIR)
    )

def step_extract_media(config) -> dict:
    """
    Executes the second pipeline step: media extraction.

    Recursively extracts images and media assets from the PowerPoint file to the 
    configured output directory. The resulting layout and media mapping data is 
    stored in the config object and returned.
    """
    print(f"{BLUE}Step 2/5: Extracting media...{RESET}")
    
    layout_data = extract_media_from_pptx(
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.MEDIA_OUTPUT_DIR)
    )
    
    config.LAYOUT_DATA_BY_SLIDE = layout_data
    return layout_data

def step_process_and_optimize_data(config) -> None:
    """
    Executes the third pipeline step: data processing and optimization.
    
    Loads the raw JSON, applies layout overrides from the PPTX, and merges 
    previously extracted media assets into the slide structure. The fully 
    integrated dataset is then saved as a cleaned JSON file.
    """
    print(f"{BLUE}Step 3/5: Process and Optimize Data...{RESET}")
    
    input_path = config.RAW_JSON_INPUT
    output_path = config.CLEANED_JSON_OUTPUT
    
    if not input_path.exists():
        print(f"{RED}Error: Input file not found: {input_path}{RESET}")
        print(f"{YELLOW}(Did Step 1 save to the wrong folder? Checked: {input_path.parent}){RESET}")
        return

    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            raw_data = json.load(f)
            
        align_map = get_text_alignment_map(str(config.PPTX_INPUT))
        slides_data = transform_docling_json_to_slides(raw_data, align_map)
        media_storage = getattr(config, 'LAYOUT_DATA_BY_SLIDE', {})
        
        if media_storage:
            for i, slide in enumerate(slides_data):
                if i in media_storage:
                    if 'elements' not in slide:
                        slide['elements'] = []
                    slide['elements'].extend(media_storage[i])

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(slides_data, f, indent=2, ensure_ascii=False)
            
        print(f"{GREEN}Data optimization complete.{RESET}")
        
    except Exception as e:
        print(f"{RED}Error during processing: {e}{RESET}")
        traceback.print_exc()

def step_generate_latex(config) -> str | None:
    """
    Executes the fourth pipeline step: LaTeX generation.

    Loads processed slide data, enriches it with geometry and grouping, and detects global headers.
    Generates the LaTeX preamble and iterates through all slides to produce individual slide code.
    Finally, assembles and returns the complete LaTeX document string.
    """
    print(f"\n{BLUE}Step 4/5: Generating LaTeX code...{RESET}")
    
    slides = load_slides(config.CLEANED_JSON_OUTPUT)
    if slides is None: 
        return None

    meta = extract_metadata(config)
    slide_width, slide_height = get_slide_dimensions(config.PPTX_INPUT)
    slides = inject_group_metrics(slides, config.PPTX_INPUT)
    slides = enrich_and_group_slides(slides, slide_width, slide_height)

    header_text = None
    header_result = detect_header_candidate(slides)
    if header_result is not None:
        header_text, header_geometry = header_result
    
    save_json(slides, config.CLEANED_JSON_OUTPUT)

    latex_preamble_code = generate_latex_preamble(meta, header_text)

    slide_blocks = []
    total_slides = len(slides)
    
    for i, slide in enumerate(slides):
        slide_num = slide.get('slide_number', i+1)
        print(f" -> Generating LaTeX for Slide {slide_num} ({i+1}/{total_slides})...")
        
        latex_code = generate_single_slide_latex(slide, config)
        block = f"\n% --- Slide {slide_num} ---\n{latex_code}\n"
        slide_blocks.append(block)

    print("Assembling final document...")
    full_body_latex = "".join(slide_blocks)
    final_latex_document = f"{latex_preamble_code}\n{full_body_latex}\n{LATEX_POSTAMBLE}"
    
    return final_latex_document

def step_save_and_compile(config, latex_code: str) -> bool:
    """
    Executes the final pipeline step: saving the LaTeX code and compiling it to PDF.

    Sanitizes the generated LaTeX string, writes it to a .tex file in the configured
    output directory, and triggers the PDF compilation process. Returns True if
    the PDF is successfully generated.
    """
    print(f"\n{BLUE}Step 5/5: Saving and Compiling...{RESET}")

    if not latex_code:
        print(f"{RED}No LaTeX code to save.{RESET}")
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
        print(f"{RED}Error saving .tex file: {e}{RESET}")
        return False

    success = compile_tex_to_pdf(tex_filename, output_dir)
    return success