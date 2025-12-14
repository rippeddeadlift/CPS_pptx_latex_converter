import json
import ollama
import re
import yaml 
from pathlib import Path

from utils import repair_latex_output

def extract_latex_content(text):
    """Entfernt Markdown ```latex Wrapper"""
    pattern = r"```(?:latex)?\s*(.*?)\s*```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1)
    return text.strip()
       
def load_conversion_rules():
    return r"""
You are a specialized LaTeX Beamer Generator.
You convert a provided JSON structure of a presentation slide into valid, compilable LaTeX code using the 'textpos' package for absolute positioning.

INPUT DATA:
You receive a JSON object representing a SINGLE slide with a list of "elements".
Each element contains:
- "type": (text, list, codeblock, table, picture, header, footer, etc.)
- "geometry": { "x", "y", "w", "h" } (Normalized coordinates 0.0-1.0)
- Content fields: "text", "items", "table_rows", "image_path", etc.

OUTPUT FORMAT RULES (STRICTLY FOLLOW):
1. **Frame Structure:**
   - Start with `\begin{frame}[fragile]`. End with `\end{frame}`.
   - NO frame title argument.

2. **Positioning (The Container):**
   - For EACH element, generate a textblock: `\begin{textblock}{<w>}(<x>, <y>) ... \end{textblock}`.
    - If "fontsize" is "3pt": Write exactly \fontsize{3}{3.3}\selectfont before the text.
3. **Content Layout (The Inner Box):**
   - Inside EVERY textblock, wrap content in a minipage.
   - Syntax:
     ```latex
     \begin{minipage}[<ALIGN>][<h>\paperheight]{\linewidth}
        <CONTENT>
     \end{minipage}
     ```
   - **CRITICAL: ALIGNMENT LOGIC (<ALIGN>):**
     - **"table", "list", "picture", "codeblock"**: ALWAYS use **[t]** (Top).
       *Explanation: Even if the geometry height (h) is large, the content must start at the top (y).*
     - **"text"**: Use **[t]** (Top) by default. Only use **[b]** if the element is strictly a label at the bottom of its box.
     - **"title", "header"**: Use **[b]** (Bottom) or **[c]** (Center).
     - **"footer"**: ALWAYS use **[b]** (Bottom) AND add `\raggedright`.

4. **Element-Specific Rendering:**
   - **"title", "header", "text"**: Output text. Use `\textbf{...}` for titles.
   - **"list"**: `\begin{itemize} \item ... \end{itemize}`. Single item -> plain text (no bullet).
   - **"codeblock"**: `\begin{lstlisting}[language=Java, basicstyle=\ttfamily\scriptsize] ... \end{lstlisting}`.
   - **"table"**:
     - Generate a standard `tabular`.
     - **IMPORTANT:** Wrap the tabular inside `\resizebox{\linewidth}{!}{ ... }` to fit width.
   - **"picture"**: `\includegraphics[width=\linewidth, height=\textheight, keepaspectratio]{...}`.
   - **Fontsize**: If "fontsize" exists, apply it INSTANTLY inside the minipage (e.g., `{\tiny ...}`).

5. **Sanitization:**
   - Escape special LaTeX chars (%, &, $, #, _) in text, but NOT in codeblocks or math ($...$).
EXAMPLES:

Input 1 (Footer - requires [b] and \raggedright):
{
  "type": "footer",
  "geometry": {"x": 0.56, "y": 0.90, "w": 0.23, "h": 0.03},
  "text": "Quelle: University of Washington",
  "fontsize": "tiny"
}

Output 1:
\begin{textblock}{0.23}(0.56, 0.90)
  \begin{minipage}[b][0.03\paperheight]{\linewidth}
    \raggedright
    {\tiny Quelle: University of Washington}
  \end{minipage}
\end{textblock}


Input 2 (List - [t] or [b]):
{
  "type": "list",
  "geometry": {"x": 0.1, "y": 0.2, "w": 0.8, "h": 0.6},
  "items": ["Point A", "Point B"],
  "fontsize": "scriptsize"
}

Output 2:
\begin{textblock}{0.8}(0.1, 0.2)
  \begin{minipage}[t][0.6\paperheight]{\linewidth}
    {\scriptsize
    \begin{itemize}
      \item Point A
      \item Point B
    \end{itemize}
    }
  \end{minipage}
\end{textblock}


Input 3 (Title - [t]):
{
  "type": "title",
  "geometry": {"x": 0.1, "y": 0.05, "w": 0.8, "h": 0.1},
  "text": "My Presentation Title"
}

Output 3:
\begin{textblock}{0.8}(0.1, 0.05)
  \begin{minipage}[t][0.1\paperheight]{\linewidth}
    \textbf{My Presentation Title}
  \end{minipage}
\end{textblock}


Input 4 (Table - requires [t] and resizebox):
{
  "type": "table",
  "geometry": {"x": 0.1, "y": 0.3, "w": 0.5, "h": 0.4},
  "table_rows": [["Col1", "Col2"], ["Val1", "Val2"]]
}

Output 4:
\begin{textblock}{0.5}(0.1, 0.3)
  \begin{minipage}[t][0.4\paperheight]{\linewidth}
    \resizebox{\linewidth}{!}{
      \begin{tabular}{|l|l|}
        Col1 & Col2 \\
        Val1 & Val2 \\
      \end{tabular}
    }
  \end{minipage}
\end{textblock}
Input 5 (Footer/Header with fixed 3pt font size):
{
  "type": "footer",
  "geometry": {"x": 0.56, "y": 0.90, "w": 0.23, "h": 0.03},
  "text": "Quelle: University of Washington",
  "fontsize": "3pt"
}

Output 5:
\begin{textblock}{0.23}(0.56, 0.90)
  \begin{minipage}[b][0.03\paperheight]{\linewidth}
    \raggedright
    \fontsize{3}{3.3}\selectfont Quelle: University of Washington
  \end{minipage}
\end{textblock}
"""

# --- 2. WORKER FUNKTION ---
def generate_single_slide_latex(slide_data, config):
    slide_num = slide_data.get('slide_number', '?')
    
    # KORREKTUR: Wir nutzen strikt den Pfad aus dem Config-Objekt!
    # Keine Defaults, kein Raten.
    rules_block = load_conversion_rules()

    # System Prompt
    system_prompt = (
        "You are a strictly constrained LaTeX Beamer generator. "
        "You do not speak. You do not explain. You only output code."
    )

    # User Prompt
    user_prompt = f"""
    TASK: Convert the following JSON slide data into a LaTeX Beamer Frame using ONLY the syntax shown below.
    
    {rules_block}
    
    INPUT DATA (Slide {slide_num}):
    {json.dumps(slide_data, indent=2, ensure_ascii=False)}
    """

    messages = [
        {'role': 'system', 'content': system_prompt},
        {'role': 'user', 'content': user_prompt}
    ]

    try:
        response = ollama.chat(model=config.AGENT_LLM_MODEL, messages=messages)
        content = response['message']['content']
        content = repair_latex_output(content)
        return extract_latex_content(content)
    except Exception as e:
        print(f"Error generating Slide {slide_num}: {e}")
        return f"% ERROR Slide {slide_num}\n\\begin{{frame}}{{Error}}\nGeneration failed.\n\\end{{frame}}"
