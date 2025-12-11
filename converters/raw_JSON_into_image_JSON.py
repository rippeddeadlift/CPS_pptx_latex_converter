import json
import utils
from collections import defaultdict


def clean_and_map_media_elements(docling_data, media_geometry_map):
    print("   -> Running Semantic Zoning (Refactored)...")
    
    # 1. SETUP & DATEN FINDEN
    data_root = utils._get_data_root(docling_data)
    page_dimensions = utils._get_page_dimensions(data_root)
    text_items = utils._get_text_items(data_root)

    # 2. ANALYSE & ZONING (Hier passiert die Magie)
    slides_data = defaultdict(lambda: {"header": [], "title": [], "content": [], "footer": []})

    for item in text_items:
            text_content = item.get("text", "").strip()
            if not text_content: continue

            page_no, bbox, origin = utils._extract_prov_data(item)
            page_height = page_dimensions.get(page_no, {}).get("height", 720)
            
            # 1. Determine Geometric Zone
            raw_zone = utils._determine_zone(bbox, origin, page_height)
            

            # Sort into the corrected bucket
            slides_data[page_no][raw_zone].append(text_content)

    # 3. ZUSAMMENBAU (Das JSON erstellen)
    return utils._assemble_final_json(slides_data, media_geometry_map)

def _recursive_remove_bits(node):
    """
    Internal helper: Walks through JSON and removes keys containing Base64 data.
    """
    if isinstance(node, dict):
        clean_dict = {}
        for key, value in node.items():
            # The Blocklist: Keys that Docling/Unstructured use for heavy data
            if key in ["bitmap", "image", "data", "uri", "base64", "binary"]:
                continue 
            
            clean_dict[key] = _recursive_remove_bits(value)
        return clean_dict
        
    elif isinstance(node, list):
        return [_recursive_remove_bits(item) for item in node]
        
    return node

def save_final_json_for_review(data, output_path):
    """
    Saves the cleaned JSON to disk.
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        print(f"   -> JSON Analysis saved to: {output_path}")
    except Exception as e:
        print(f"   -> Error saving JSON: {e}")