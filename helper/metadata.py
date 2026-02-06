from collections import defaultdict
from typing import List, Dict, Any


def get_bbox_sort_key(item: Dict[str, Any]) -> tuple[float, float]:
    """
    Generates a sort key for reading order (Top-Down, Left-Right).

    Extracts bounding box data from the item's provenance. Assumes a coordinate 
    system where Y increases upwards (PDF style), using negation on the vertical 
    component to sort from visual top to bottom.
    """
    prov_list = item.get("prov", [])
    prov = prov_list[0] if prov_list else {}
    
    bbox = prov.get("bbox", {})
    
    return (-bbox.get("t", 0.0), bbox.get("l", 0.0))


def simplify_table_data(table_item: Dict[str, Any]) -> list[list[str]]:
    """
    Converts a complex table grid into a simple text matrix.
    
    Extracts the text content from each cell in the grid structure, stripping whitespace.
    Rows that are completely empty are filtered out to ensure a clean output.
    """
    if "data" not in table_item or "grid" not in table_item["data"]:
        return []

    simple_rows = []
    grid = table_item["data"]["grid"]

    for row in grid:
        simple_row = [cell.get("text", "").strip() for cell in row]
        
        if any(simple_row): 
            simple_rows.append(simple_row)

    return simple_rows

def transform_docling_json_to_slides(raw_data: Dict[str, Any], alignment_map: Dict = None) -> List[Dict[str, Any]]:
    """
    Converts raw document analysis JSON into a structured, slide-based format.

    Groups content by page, simplifies tables, assigns image paths, and sorts elements 
    vertically. Applies alignment overrides if text matches the provided alignment map.
    """
    if alignment_map is None: alignment_map = {}
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
                
                if "text" in item:
                    text_content = item["text"].strip()
                    if not text_content: continue 
                
                element = {
                    "type": key[:-1], 
                    "label": item.get("label", "unknown"),
                    "bbox": {k: int(v) for k, v in provs[0].get("bbox", {}).items() if isinstance(v, (int, float))}
                }
                
                if "text" in item: 
                    element["text"] = item["text"].strip()
                
                if "items" in item:
                    element["items"] = item["items"]

                check_text = ""
                
                if element.get("text"):
                    check_text = element["text"]
                elif element.get("items"):
                    for it in element["items"]:
                        if isinstance(it, str) and it.strip():
                            check_text = it
                            break
                
                if check_text and page_no in alignment_map:
                    lookup_key = "".join(check_text.split()).lower()[:50]
                    
                    if lookup_key in alignment_map[page_no]:
                        element["align"] = "b" 
                
                if key == "tables": 
                    element["table_rows"] = simplify_table_data(item)
                
                if key == "pictures":
                    filename = f"image_{global_image_counter}.png"
                    element["image_path"] = f"extracted_media/{filename}"
                    global_image_counter += 1
                
                slides_buckets[page_no].append(element)
                

    final_slides = []
    for page_num in sorted(slides_buckets.keys()):
        raw_items = slides_buckets[page_num]
        
        sorted_items = sorted(raw_items, key=get_bbox_sort_key)
        
        slide_obj = {
            "slide_number": page_num,
            "elements": sorted_items 
        }
        final_slides.append(slide_obj)
        
    return final_slides
