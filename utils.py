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

import json
import os

def prepare_initial_prompt(rules_file_path, json_file_path):
    """
    Reads system rules and enriched Docling Markdown.
    Constructs a specific prompt instructing the LLM to use Layout Data.
    """
    if not os.path.exists(rules_file_path):
        raise FileNotFoundError(f"Rules file not found: {rules_file_path}")
    if not os.path.exists(json_file_path):
        raise FileNotFoundError(f"JSON file not found: {json_file_path}")

    with open(rules_file_path, 'r', encoding='utf-8') as f:
        rules_content = f.read()

    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    markdown_content = data.get('content_markdown', '')
    
    if not markdown_content:
        markdown_content = json.dumps(data, indent=2, ensure_ascii=False)

    prompt = (
        f"Role: You are an expert LaTeX Beamer developer.\n"
        f"Objective: Convert the following Markdown content into a compilable LaTeX Beamer presentation.\n\n"
        f"--- SYSTEM RULES ---\n"
        f"{rules_content}\n\n"
        f"--- INPUT CONTENT (Markdown + Layout Data) ---\n"
        f"{markdown_content}\n\n"
        f"--- INSTRUCTIONS ---\n"
        f"1. Analyze the 'LAYOUT & MEDIA DATA' section at the bottom of the input.\n"
        f"2. Map images/videos to their slides based on the Slide Number.\n"
        f"3. Use the [GEOMETRY] hints to decide on layout (e.g., use \\begin{{columns}} if x > 0.5).\n"
        f"4. Output ONLY the LaTeX code."
    )
    
    return prompt