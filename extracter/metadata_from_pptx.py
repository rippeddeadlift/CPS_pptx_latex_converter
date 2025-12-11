import re
from pptx import Presentation

def extract_metadata_from_pptx(pptx_path):
    print("   -> Mining Global Metadata (Geometry + Keywords)...")
    prs = Presentation(pptx_path)
    slide_height = prs.slide_height # Total height in EMUs
    
    metadata = {
        "title": prs.core_properties.title or "",
        "author": prs.core_properties.author or "",
        "subject": prs.core_properties.subject or "",
        
        # Two distinct buckets
        "global_header_text": [], 
        "global_footer_text": [], 
        
        "classified_date": "",
        "raw_background_context": []
    }

    collected_shapes = [] # Store (text, y_position)

    # --- 1. SCAN MASTER & LAYOUTS ---
    scan_targets = list(prs.slide_masters)
    for m in prs.slide_masters: scan_targets.extend(m.slide_layouts)

    seen_texts = set()

    for layer in scan_targets:
        for shape in layer.shapes:
            if shape.has_text_frame and shape.text.strip():
                text = shape.text.strip()
                
                # Deduplication
                if text in seen_texts: continue
                seen_texts.add(text)
                
                # Filter Garbage (Blocklist)
                if not _is_valid_meta_text(text): continue
                
                # Store with Geometry
                # shape.top is distance from top edge
                collected_shapes.append({
                    "text": text,
                    "y": shape.top, 
                    "height": shape.height
                })

    # --- 2. CLASSIFY BASED ON POSITION & CONTENT ---
    for item in collected_shapes:
        text = item["text"]
        y = item["y"]
        text_low = text.lower()
        
        # A. DATE DETECTION (Priority)
        if re.search(r'\d{2}\.\d{2}\.\d{4}', text):
            metadata["classified_date"] = text
            continue

        # B. KEYWORD MAGNETS (Force specific content into buckets)
        # Institute/Prof keywords -> FORCE FOOTER
        footer_keywords = ["prof.", "dr.", "fakultät", "faculty", "institut", "university", "gmbh", "©"]
        if any(k in text_low for k in footer_keywords):
            # We treat this as footer info (Institute), regardless of position
            metadata["global_footer_text"].append(text)
            continue

        # C. GEOMETRY ZONING (For everything else)
        relative_y = y / slide_height
        
        if relative_y < 0.20:
            # Top 20% -> Header
            metadata["global_header_text"].append(text)
        elif relative_y > 0.80:
            # Bottom 20% -> Footer
            metadata["global_footer_text"].append(text)
        else:
            # Middle -> Context
            metadata["raw_background_context"].append(text)

    # Convert lists to strings for JSON
    metadata["global_header_text"] = " | ".join(metadata["global_header_text"])
    metadata["global_footer_text"] = " | ".join(metadata["global_footer_text"])

    return metadata

def _is_valid_meta_text(txt):
    """Helper to kill placeholder noise"""
    if len(txt) < 3: return False
    txt_low = txt.lower()
    noise = ["text", "ebene", "level", "titel", "date", "datum", 
             "footer", "fußzeile", "nr.", "<nr>", "#", "click to edit", 
             "masterformat", "bild durch klicken", "symbol hinzufügen", 
             "hier text eingeben", "überschrift"]
    
    if any(n in txt_low for n in noise): return False
    if txt.replace(".", "").replace("/", "").isdigit(): return False # Pure page numbers
    return True