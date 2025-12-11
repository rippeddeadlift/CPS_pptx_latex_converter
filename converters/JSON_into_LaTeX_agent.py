import ollama
import re
from utils import compile_tex_to_pdf
from validators import check_media_completeness

MAX_RETRIES = 3
LLM_MODEL = 'qwen3:8b' 

def extract_latex_content(text):
    """
    Isoliert den LaTeX-Code zwischen \documentclass und \end{document}.
    Entfernt Markdown-Blöcke und Chat-Antworten.
    """
    # 1. Markdown-Code-Syntax entfernen
    text = text.replace("```latex", "").replace("```", "").strip()
    
    # 2. Suche nach dem echten Start (\documentclass)
    # Wir nutzen Regex, um auch Whitespace-Variationen zu fangen
    start_pattern = r"\\documentclass"
    end_pattern = r"\\end\{document\}"
    
    start_match = re.search(start_pattern, text)
    end_match = re.search(end_pattern, text)
    
    if start_match:
        # Schneide alles vor \documentclass weg
        text = text[start_match.start():]
    else:
        # Fallback: Wenn kein documentclass gefunden wurde, ist der Code eh kaputt,
        # aber wir geben das Beste zurück, was wir haben.
        print("Warning: No \\documentclass found in LLM output.")

    if end_match:
        # Schneide alles nach \end{document} weg
        text = text[:end_match.end()]
        
    return text

def generate_latex_with_retry(config, prompt_initial):
    messages = [
        {'role': 'user', 'content': prompt_initial}
    ]
    
    print(f"Starting generation loop (Max retries: {MAX_RETRIES})...")

    for attempt in range(1, MAX_RETRIES + 1):
        print(f"--- Attempt {attempt}/{MAX_RETRIES} ---")
        
        try:
            response = ollama.chat(
                model=LLM_MODEL,
                messages=messages
            )
            raw_content = response['message']['content']
            
            # --- HIER PASSIERT DIE MAGIE ---
            latex_code = extract_latex_content(raw_content)
            # -------------------------------

            # Validator Check (Missing Media)
            missing_errors = check_media_completeness(config.CLEANED_JSON_OUTPUT, latex_code)
            
            if missing_errors:
                print(f"Validation Failed: Found {len(missing_errors)} missing items.")
                error_msg = "\n".join(missing_errors)
                
                feedback = (
                    f"The generated code is incomplete. You ignored strict rules.\n"
                    f"ERRORS FOUND:\n{error_msg}\n"
                    f"RULE REMINDER: If you see .mp4, you MUST use the \\includemedia block defined in the rules.\n"
                    f"Please regenerate the FULL LaTeX code correcting these errors."
                )
                
                # Update History
                messages.append({'role': 'assistant', 'content': raw_content}) # Wir geben dem LLM seine rohe Antwort zurück
                messages.append({'role': 'user', 'content': feedback})
                
                print("Feedback sent to LLM: Missing Media.")
                continue 

            # Syntax Check (Compiler)
            temp_tex = config.RESULTS_DIR / "temp_debug.tex"
            with open(temp_tex, "w", encoding="utf-8") as f:
                f.write(latex_code)
                
            is_valid_syntax = compile_tex_to_pdf(temp_tex, config.RESULTS_DIR)
            
            if is_valid_syntax:
                print("Validation Passed: Syntax is correct.")
                return latex_code
            else:
                print("Validation Failed: Compilation Error.")
                
                log_file = config.RESULTS_DIR / "temp_debug.log"
                log_content = "Log not found."
                if log_file.exists():
                    with open(log_file, 'r', encoding='latin-1') as f:
                        lines = f.readlines()
                        log_content = "\n".join(lines[-20:])
                
                feedback = (
                    f"The code has syntax errors and cannot compile.\n"
                    f"LATEX COMPILER LOG:\n{log_content}\n"
                    f"Please fix the syntax errors and regenerate the code.\n"
                    f"CRITICAL: Output ONLY valid LaTeX code starting with \\documentclass."
                )
                
                messages.append({'role': 'assistant', 'content': raw_content})
                messages.append({'role': 'user', 'content': feedback})
                
                continue

        except Exception as e:
            print(f"Error during LLM call: {e}")
            import traceback
            traceback.print_exc()
            break

    print("Max retries reached. Returning last generated result.")
    return latex_code # Im schlimmsten Fall geben wir den letzten Versuch zurück