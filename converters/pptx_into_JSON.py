import json
from pathlib import Path
from docling.document_converter import DocumentConverter

async def convert_pptx_to_json(pptx_path: str, output_dir: str):
    """
    Uses IBM Docling to convert PPTX to a structured representation.
    Saves the output as a JSON file containing the Markdown representation
    and structured dictionary for further processing.
    """
    input_path = Path(pptx_path)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    
    json_output_path = out_path / (input_path.stem + ".json")
    
    print(f"Docling: Parsing {input_path.name} locally...")

    try:
        converter = DocumentConverter()
        result = converter.convert(input_path)
        
        markdown_content = result.document.export_to_markdown()
        structured_dict = result.document.export_to_dict()
        
        final_data = {
            "filename": input_path.name,
            "type": "docling_converted",
            "content_markdown": markdown_content,
            "structure_analysis": structured_dict
        }

        with open(json_output_path, 'w', encoding='utf-8') as f:
            json.dump(final_data, f, indent=2, ensure_ascii=False)

        print(f"Docling conversion finished: {json_output_path}")
        return str(json_output_path)

    except Exception as e:
        print(f"Docling Error: {e}")
        raise e