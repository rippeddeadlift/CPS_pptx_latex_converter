import json

def clean_and_map_media_elements(raw_data, media_map, layout_data=None):
    if isinstance(raw_data, dict) and raw_data.get('type') == 'docling_converted':
        
        original_markdown = raw_data.get('content_markdown', '')
        
        media_appendix = "\n\n" + "="*50 + "\n"
        media_appendix += "--- LAYOUT & MEDIA DATA ---\n"
        media_appendix += "Use this data to position images correctly.\n"
        media_appendix += "Coordinates are percentages (0.0 to 1.0). x=0.0 (Left), y=0.0 (Top)\n\n"
        
        sorted_slides = sorted(media_map.keys(), key=lambda x: int(x) if str(x).isdigit() else 999)
        
        for slide_num in sorted_slides:
            raw_ref = media_map[slide_num]
            
            # 1. Filename extrahieren
            filename = None
            if isinstance(raw_ref, str): filename = raw_ref
            elif isinstance(raw_ref, list) and len(raw_ref) > 0: filename = raw_ref[0]
            elif isinstance(raw_ref, dict) and len(raw_ref) > 0: filename = next(iter(raw_ref.values()))
            
            if not filename or not isinstance(filename, str): continue
            
            # 2. Lookup Namen vorbereiten
            lookup_primary = filename
            lookup_secondary = None
            
            # Splitten bei "media1.mp4|image3.png"
            if "|" in filename:
                parts = filename.split("|")
                lookup_primary = parts[0].strip()   # media1.mp4
                if len(parts) > 1:
                    lookup_secondary = parts[1].strip() # image3.png

            # 3. Geometrie suchen (Primary ODER Secondary!)
            geo_info = ""
            g = None
            
            if layout_data and lookup_primary in layout_data:
                g = layout_data[lookup_primary]
            elif layout_data and lookup_secondary and lookup_secondary in layout_data:
                g = layout_data[lookup_secondary]
                geo_info += "   [NOTE]: Using geometry from preview image.\n"

            if g:
                geo_info += (f"   [GEOMETRY]: "
                             f"Pos_X={g['x']}, Pos_Y={g['y']} | "
                             f"Width={g['w']}, Height={g['h']}")
                
                # Hints für das LLM
                if g['w'] > 0.9: geo_info += " -> (Use full width)"
                elif g['x'] > 0.5: geo_info += " -> (Place on RIGHT side)"
                elif g['x'] < 0.5 and g['w'] < 0.5: geo_info += " -> (Place on LEFT side)"

            media_appendix += f"Slide {slide_num}: {filename}\n{geo_info}\n"
            media_appendix += "-"*30 + "\n"

        media_appendix += "="*50 + "\n"
        raw_data['content_markdown'] = original_markdown + media_appendix
        return raw_data

    return raw_data

def save_final_json_for_review(data, output_path):
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)