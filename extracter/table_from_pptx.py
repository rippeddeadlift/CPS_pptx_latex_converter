def _extract_tables_from_docling(docling_data):
    """
    Extracts table data from the raw JSON 'tables' node.
    Converts the structured grid into a text representation 
    that mimics a Markdown/CSV format.
    """
    items = []
    
    # 1. Access the tables list
    # Note: 'docling_data' might be the root dict. 
    tables = docling_data.get("tables", [])
    
    for tbl in tables:
        # Safety check: We need provenance to know the page number
        if not tbl.get("prov") or not tbl.get("data"): 
            continue
            
        # 2. Build the Text Representation
        # We start with a marker so the LLM knows: "STOP! This is a table."
        lines = ["DETECTED_TABLE_START"]
        
        # Access the grid (Rows -> Cells)
        # Based on your snippet: data -> grid -> [ [cell, cell], [cell, cell] ]
        grid = tbl["data"].get("grid", [])
        
        for row in grid:
            row_texts = []
            for cell in row:
                # Extract text from the cell object
                # Replace newlines in cells with spaces to keep the row on one line
                cell_text = cell.get("text", "").strip().replace("\n", " ")
                row_texts.append(cell_text)
            
            # Join columns with a pipe character
            # Format: "Description | O-Notation | Runtime..."
            lines.append(" | ".join(row_texts))
            
        lines.append("DETECTED_TABLE_END")
        
        full_table_text = "\n".join(lines)

        # 3. Create a standard Item object
        # This matches the structure of normal text items so we can merge them easily.
        items.append({
            "text": full_table_text,
            "prov": tbl["prov"] # Passes bbox and page_no to the zoning logic
        })
        
    print(f"   -> Extracted {len(items)} tables from raw JSON.")
    return items