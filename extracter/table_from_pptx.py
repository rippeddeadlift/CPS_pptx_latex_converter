def _extract_tables_from_docling(docling_data):
    """
    Extracts table data from the raw JSON 'tables' node.
    Converts the structured grid into a text representation 
    that mimics a Markdown/CSV format.
    """
    items = []
    
    tables = docling_data.get("tables", [])
    
    for tbl in tables:
        if not tbl.get("prov") or not tbl.get("data"): 
            continue
            
        lines = ["DETECTED_TABLE_START"]
        
        grid = tbl["data"].get("grid", [])
        
        for row in grid:
            row_texts = []
            for cell in row:
                cell_text = cell.get("text", "").strip().replace("\n", " ")
                row_texts.append(cell_text)
            
            lines.append(" | ".join(row_texts))
            
        lines.append("DETECTED_TABLE_END")
        
        full_table_text = "\n".join(lines)

        items.append({
            "text": full_table_text,
            "prov": tbl["prov"] # Passes bbox and page_no to the zoning logic
        })
        
    print(f"   -> Extracted {len(items)} tables from raw JSON.")
    return items