import json
from collections import defaultdict
from typing import List, Dict, Any
import re

def get_bbox_sort_key(item: Dict[str, Any]):
    """Sortiert Elemente vertikal (Top -> Down)."""
    prov = item.get("prov", [{}])[0]
    bbox = prov.get("bbox", {})
    # -t sortiert absteigend (Top ist bei Bottom-Left Origin der höchste Wert)
    return (-bbox.get("t", 0), bbox.get("l", 0))

def simplify_table_data(table_item: Dict[str, Any]) -> List[List[str]]:
    """Wandelt Grid in Matrix um."""
    if "data" not in table_item or "grid" not in table_item["data"]:
        return []
    simple_rows = []
    grid = table_item["data"]["grid"]
    for row in grid:
        simple_row = [cell.get("text", "").strip() for cell in row]
        if any(simple_row): 
            simple_rows.append(simple_row)
    return simple_rows

def transform_docling_json_to_slides(raw_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    # 1. Choose Data Source
    if "structure_analysis" in raw_data:
        source_data = raw_data["structure_analysis"]
    else:
        source_data = raw_data

    slides_buckets = defaultdict(list)
    global_image_counter = 1 
    
    content_keys = ["texts", "tables", "pictures"]
    
    for key in content_keys:
        if key in source_data:
            items = source_data[key]
            
            for item in items:
                provs = item.get("prov", [])
                if not provs: continue
                
                page_no = provs[0].get("page_no")
                
                # Filter Text Elements
                if "text" in item:
                    text_content = item["text"].strip()
                    if not text_content: continue 
                
                # Create Element
                element = {
                    "type": key[:-1], # texts -> text
                    "label": item.get("label", "unknown"),
                    "bbox": {k: int(v) for k, v in provs[0].get("bbox", {}).items() if isinstance(v, (int, float))}
                }
                
                if "text" in item: 
                    element["text"] = item["text"].strip()
                
                if key == "tables": 
                    element["table_rows"] = simplify_table_data(item)
                
                if key == "pictures":
                    filename = f"image_{global_image_counter}.png"
                    element["image_path"] = f"extracted_media/{filename}"
                    global_image_counter += 1
                
                slides_buckets[page_no].append(element)

    # Build Final List
    final_slides = []
    for page_num in sorted(slides_buckets.keys()):
        raw_items = slides_buckets[page_num]
        
        # Sort Top-to-Bottom
        sorted_items = sorted(raw_items, key=get_bbox_sort_key)
        
        slide_obj = {
            "slide_number": page_num,
            "elements": sorted_items # Keep elements as is (LLM handles rest)
        }
        final_slides.append(slide_obj)
        
    return final_slides

# def is_generic_noise(text_content: str) -> bool:
#     """
#     Detects Date, Page Numbers, and specific footer noise.
#     Returns True if the text should be deleted.
#     """
#     text = text_content.strip()
    
#     # 1. Date (dd.mm.yyyy or yyyy-mm-dd)
#     if re.match(r'^\d{2}\.\d{2}\.\d{4}$', text): return True
#     if re.match(r'^\d{4}-\d{2}-\d{2}$', text): return True
        
#     # 2. Page Numbers & Footer Digits
#     # Matches "Seite 5", "Page 5", "5/20", or just "5"
#     if re.match(r'^(Seite|Page)\s+\d+$', text, re.IGNORECASE): return True
#     if re.match(r'^\d+\s*/\s*\d+$', text): return True 
#     if re.match(r'^\d+$', text): return True 
    
#     return False