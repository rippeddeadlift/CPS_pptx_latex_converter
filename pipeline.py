import json
from converters.JSON_into_LaTeX_agent import generate_latex_with_retry
from converters.pptx_into_JSON import convert_pptx_to_json
from converters.raw_JSON_into_image_JSON import clean_and_map_media_elements, save_final_json_for_review
from extracter.media_from_pptx import extract_media_from_pptx
from utils import prepare_initial_prompt

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
    print("Step 2/5: Extracting media and geometry...")
    extracted_files, layout_data = extract_media_from_pptx( 
        pptx_path=str(config.PPTX_INPUT),
        output_dir=str(config.MEDIA_OUTPUT_DIR)
    )
    global LAYOUT_DATA_STORAGE
    LAYOUT_DATA_STORAGE = layout_data 
    
    return extracted_files

def step_map_and_clean_data(config):
    print("Step 3/5: Merging Text and Geometry...")
    
    media_map = find_prioritized_media_references(str(config.PPTX_INPUT))
    final_media_map = prioritize_and_clean_media_map(media_map)
    
    with open(config.RAW_JSON_INPUT, 'r', encoding='utf-8') as f:
        raw_json = json.load(f)

    final_json = clean_and_map_media_elements(
        raw_json, 
        final_media_map, 
        layout_data=LAYOUT_DATA_STORAGE 
    )
    
    save_final_json_for_review(final_json, str(config.CLEANED_JSON_OUTPUT))
    return final_json

def step_generate_latex(config):
    print(f"{BLUE}Step 4/5: Generating LaTeX code with Auto-Correction...{RESET}")
    

    prompt_text = prepare_initial_prompt(
        rules_file_path=config.RULES_FILE,
        json_file_path=str(config.CLEANED_JSON_OUTPUT)
    )
    final_latex = generate_latex_with_retry(config, prompt_text)
    
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
