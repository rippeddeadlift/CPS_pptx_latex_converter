import json

def check_media_completeness(json_path, latex_code):
    """
    Scans the source JSON and checks if all media files (images/videos)
    are present in the generated LaTeX code.
    Returns a list of missing files or error messages.
    """
    missing_items = []
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Validator Warning: Could not read JSON to check media: {e}")
        return []

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