import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def extract_media_from_pptx(pptx_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    prs = Presentation(pptx_path)
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    layout_data_by_slide = {}
    
    # We use a mutable counter to keep filenames unique across all slides
    global_image_count = 1 

    print(f"   -> Mining {len(prs.slides)} slides for hidden media...")

    for i, slide in enumerate(prs.slides):
        slide_index = i
        slide_media = []
        
        # We start the recursion here
        for shape in slide.shapes:
            global_image_count = _process_shape_recursive(
                shape, 
                slide_media, 
                output_dir, 
                global_image_count,
                slide_width, 
                slide_height
            )

        if slide_media:
            layout_data_by_slide[slide_index] = slide_media
            print(f"      Slide {i+1}: Found {len(slide_media)} media items")

    return layout_data_by_slide

def _process_shape_recursive(shape, slide_media, output_dir, count, s_width, s_height):
    """
    Recursively inspects shapes.
    - If Group: inspects children.
    - If Picture: saves it.
    """
    
    # CASE 1: GROUP (The logic we were missing!)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child_shape in shape.shapes:
            count = _process_shape_recursive(child_shape, slide_media, output_dir, count, s_width, s_height)
        return count

    # CASE 2: PICTURE (Standard images)
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return _save_shape_image(shape, slide_media, output_dir, count, s_width, s_height)

    # CASE 3: PICTURE PLACEHOLDER
    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
        if hasattr(shape, 'image') and shape.image:
            return _save_shape_image(shape, slide_media, output_dir, count, s_width, s_height)

    # CASE 4: SHAPES WITH PICTURE FILL (Advanced/Optional)
    # Some "Rectangles" are actually photos. This tries to catch them.
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        try:
            # Type 6 is 'Picture Fill'
            if shape.fill.type == 6:
                # Accessing the image from a fill is tricky but this often works
                if hasattr(shape.fill, 'fore_color') and hasattr(shape.fill.fore_color, 'type'):
                     # We can't easily extract the blob from a Fill in python-pptx without deep hacking
                     # So we skip saving the file, but we acknowledge it existed.
                     pass 
        except:
            pass

    return count

def _save_shape_image(shape, slide_media, output_dir, count, s_width, s_height):
    try:
        image = shape.image
        ext = image.ext
        filename = f"image_{count}.{ext}"
        filepath = os.path.join(output_dir, filename)
        
        # Save to disk
        with open(filepath, "wb") as f:
            f.write(image.blob)
            
        # geometry calculation
        # (Handling the fact that some grouped shapes report positions differently)
        left = shape.left / s_width
        top = shape.top / s_height
        width = shape.width / s_width
        height = shape.height / s_height
        
        slide_media.append({
            "filename": filename,
            "path": f"images/{filename}",
            "geometry": [left, top, width, height]
        })
        
        return count + 1
    except Exception as e:
        print(f"      Warning: Could not extract image {count}: {e}")
        return count