import ollama
import re
from pathlib import Path
from utils import compile_tex_to_pdf
from validators import check_media_completeness



def extract_latex_content(text):
    """
    Verbesserte Extraktion: Entfernt Markdown-Wrapper robuster
    und sucht präziser nach dem Dokumenten-Rumpf.
    """
    # 1. Entferne Code-Block Marker (egal ob ```latex, ```tex oder nur ```)
    pattern = r"```(?:latex|tex)?\s*([\s\S]*?)\s*```"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        text = match.group(1)
    
    # 2. Suche Start und Ende
    # Manche LLMs schreiben Text VOR \documentclass, das muss weg.
    start_pattern = re.compile(r"(\\documentclass.*)", re.IGNORECASE | re.DOTALL)
    end_pattern = re.compile(r"(\\end\s*\{\s*document\s*\})", re.IGNORECASE)
    
    start_match = start_pattern.search(text)
    end_match = end_pattern.search(text)
    
    if start_match:
        text = text[start_match.start():]
    
    if end_match:
        text = text[:end_match.end()]
        
    return text.strip()

def parse_latex_log_errors(log_path):
    """
    Extrahiert NUR die relevanten Fehlerzeilen aus dem Log.
    LaTeX Logs sind sehr lang, das LLM braucht nur die Zeilen, die mit '!' beginnen.
    """
    errors = []
    if not log_path.exists():
        return "Log file missing."
        
    with open(log_path, 'r', encoding='latin-1', errors='replace') as f:
        lines = f.readlines()
        
    for i, line in enumerate(lines):
        if line.startswith('!'):
            context = "".join(lines[i:i+2]).strip()
            errors.append(context)
        elif "Error:" in line:
            errors.append(line.strip())

    if not errors:
        return "\n".join(lines[-15:])
    
    return "\n".join(errors[:5])

def generate_latex_with_retry(config, prompt_initial):
    system_prompt = (
    "You are a strict LaTeX Beamer code generator machine. "
    "OBJECTIVE: Convert the provided JSON data into a compiling LaTeX Beamer presentation."
    "CRITICAL RULES:\n"
    "1. Output ONLY valid LaTeX code. No explanations, no markdown blocks.\n"
    "2. Start strictly with \\documentclass{beamer} and end with \\end{document}.\n"
    "3. Use [fragile] for ANY frame containing code listings.\n"
    "4. Do NOT hallucinate content. Use strictly the provided JSON text and images."
)

    messages = [
        {'role': 'system', 'content': system_prompt},
        {'role': 'user', 'content': prompt_initial}
    ]

    print(f"Starting Agentic Generation (Model: {config.AGENT_LLM_MODEL})...")
    latex_code = ""

    for attempt in range(1, config.AGENT_MAX_RETRIES + 1):
        print(f"\n--- Attempt {attempt}/{config.AGENT_MAX_RETRIES} ---")
        
        try:
            # 1. Call LLM
            response = ollama.chat(model=config.AGENT_LLM_MODEL, messages=messages)
            raw_content = response['message']['content']
            
            # 2. Sanitize
            latex_code = extract_latex_content(raw_content)
            
            if not latex_code:
                print("Warning: Received empty or invalid LaTeX content.")
                messages.append({'role': 'assistant', 'content': raw_content})
                messages.append({'role': 'user', 'content': "Error: No valid \\documentclass found. Please output ONLY the LaTeX code."})
                continue

            # 3. Validate Logic (Images)
            missing_errors = check_media_completeness(config.CLEANED_JSON_OUTPUT, latex_code)
            
            if missing_errors:
                print(f"Logical Error: Missing {len(missing_errors)} images.")
                feedback = (
                    f"Logic Check Failed: Your code is missing required media files.\n"
                    f"MISSING FILES:\n" + "\n".join(missing_errors) + "\n"
                    f"Please regenerate the code and ensure ALL images are included via \\includegraphics."
                )
                messages.append({'role': 'assistant', 'content': raw_content})
                messages.append({'role': 'user', 'content': feedback})
                continue 

            # 4. Validate Syntax (Compilation)
            temp_tex = config.RESULTS_DIR / f"temp_attempt_{attempt}.tex"
            with open(temp_tex, "w", encoding="utf-8") as f:
                f.write(latex_code)
            
            print(f"   -> Compiling syntax check...")
            # Wichtig: compile_tex_to_pdf muss True/False zurückgeben
            is_valid_syntax = compile_tex_to_pdf(temp_tex, config.RESULTS_DIR)
            
            if is_valid_syntax:
                print("SUCCESS: PDF compiled and logic verified!")
                return latex_code 
                
            else:
                print("Syntax Error: Compilation failed.")
                log_file = config.RESULTS_DIR / f"temp_attempt_{attempt}.log"
                
                # Hier rufen wir die verbesserte Log-Parsing Funktion auf
                error_snippet = parse_latex_log_errors(log_file)
                
                feedback = (
                    f"Compilation Failed. Fix the LaTeX syntax errors based on this log output:\n"
                    f"```\n{error_snippet}\n```\n"
                    f"Common fixes: Escape special characters (%, _, &), check closing brackets, ensure environment names are correct."
                )
                
                messages.append({'role': 'assistant', 'content': raw_content})
                messages.append({'role': 'user', 'content': feedback})
                continue

        except Exception as e:
            print(f"Critical Error: {e}")
            break

    print("Max retries reached. Returning best effort (may be broken).")
    return latex_code