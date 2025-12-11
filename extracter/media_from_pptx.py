import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def extract_media_from_pptx(pptx_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    prs = Presentation(pptx_path)
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    extracted_files = []
    layout_data = {} 

    image_count = 1
    video_count = 1

    print(f"Analyzing layout and mining media for {len(prs.slides)} slides...")

    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        
        for shape in slide.shapes:
            is_captured = False

            # --- CASE 1: BILDER & PLACEHOLDER ---
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                if hasattr(shape, "image") and shape.image:
                    
                    # 1. ALWAYS extract the image (Thumbnail/Poster)
                    image = shape.image
                    # Use generic naming or try to keep extension
                    ext = image.ext
                    image_filename = f"image{image_count}.{ext}"
                    filepath = os.path.join(output_dir, image_filename)
                    
                    with open(filepath, "wb") as f:
                        f.write(image.blob)
                    
                    # Save Geometry for Image
                    layout_data[image_filename] = _get_geo(shape, slide_width, slide_height, "image")
                    extracted_files.append(image_filename)
                    image_count += 1
                    is_captured = True
                    
                    # 2. CHECK FOR HIDDEN VIDEO (in the same shape)
                    # If it looks like a video placeholder, try to rescue the video file too
                    if "media" in shape.name.lower() or "demonstrated" in shape.name.lower():
                        video_filename = f"media{video_count}.mp4"
                        found_video = _rescue_video_from_rels(slide, output_dir, video_filename)
                        
                        if found_video:
                            print(f"  -> Match: Slide {slide_num} has Image '{image_filename}' AND Video '{video_filename}'")
                            # Save Geometry for Video (same as image)
                            layout_data[video_filename] = _get_geo(shape, slide_width, slide_height, "video")
                            extracted_files.append(video_filename)
                            video_count += 1

            # --- CASE 2: LEGACY MEDIA OBJECTS ---
            elif shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                filename = f"media{video_count}.mp4"
                layout_data[filename] = _get_geo(shape, slide_width, slide_height, "video")
                _create_placeholder(output_dir, filename) 
                extracted_files.append(filename)
                video_count += 1
                is_captured = True

    return extracted_files, layout_data

def _rescue_video_from_rels(slide, output_dir, filename):
    """
    Scans the slide relationships to find hidden video files.
    """
    try:
        for rel in slide.part.rels.values():
            if "video" in rel.target_part.content_type or "media" in rel.target_part.content_type:
                video_blob = rel.target_part.blob
                filepath = os.path.join(output_dir, filename)
                
                with open(filepath, "wb") as f:
                    f.write(video_blob)
                
                return True
    except Exception:
        pass
    
    return False

def _get_geo(shape, sw, sh, type_label):
    return {
        "file": "N/A", 
        "type": type_label,
        "x": round(shape.left / sw, 3),
        "y": round(shape.top / sh, 3),
        "w": round(shape.width / sw, 3),
        "h": round(shape.height / sh, 3)
    }

def _create_placeholder(output_dir, filename):
    filepath = os.path.join(output_dir, filename)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as f: 
            f.write("Placeholder")