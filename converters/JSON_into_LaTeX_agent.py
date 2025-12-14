import json
import ollama
import re
import yaml 
from pathlib import Path

def extract_latex_content(text):
    """Entfernt Markdown ```latex Wrapper"""
    pattern = r"```(?:latex)?\s*(.*?)\s*```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1)
    return text.strip()
       
def load_conversion_rules():
    return r"""
IMPORTANT RULES. STRICTLY FOLLOW, NO EXCEPTIONS!

- Output LaTeX for a SINGLE slide. Start with \begin{frame}[fragile] (if there is a codeblock; else just \begin{frame}), end with \end{frame}.
- DO NOT use a frame title argument (no \begin{frame}{...}).
- For each element with a geometry field, create exactly one \begin{textblock}{WIDTH}(X, Y) ... \end{textblock}.
- WIDTH, X, Y are decimal fractions from the geometry field ("w", "x", "y").
- All content (text, lists, code, images, tables) must be INSIDE their respective textblock, with no wrapper outside.
- NEVER use itemize, tabular, images, code, or math outside a textblock. NEVER nest textblocks.

SPECIAL CASES:
- For the title ("type": "title"): Always output as the FIRST textblock, using its geometry, as \textbf{...} (bold). DO NOT set a frame title argument.
- If there is a codeblock, add [fragile] to \begin{frame}, e.g. \begin{frame}[fragile]. If none, use \begin{frame} only.
- For a "list": single item → output as plain text within the textblock (no bullet). Multiple items → use \begin{itemize}...\end{itemize} (inside the textblock).
- For "codeblock": output a single textblock with \begin{lstlisting}...\end{lstlisting}.
- For "table": inside the textblock, wrap the tabular inside \resizebox{\textwidth}{!}{...}.
- For "picture": use \includegraphics[width=\textwidth]{...} inside the textblock.
- If an element contains a field "fontsize" (e.g., "fontsize": "scriptsize"), wrap the entire content inside the textblock in curly braces with the respective LaTeX font size command, e.g. {\scriptsize ... }. This must come immediately inside the textblock, before the content, and end after the content.

EXAMPLE INPUT:
[
  {"type":"list", "geometry":{"x":0.2,"y":0.15,"w":0.7}, "items":["a","b","c","d","e","f","g","h","i","j","k"], "fontsize":"scriptsize"},
  {"type":"title","text":"Mein Titel","geometry":{"x":0.1,"y":0.05,"w":0.6}},
  {"type":"text","text":"Hallo!","geometry":{"x":0.2,"y":0.1,"w":0.5}},
  {"type":"list", "items":["Foo","Bar"],"geometry":{"x":0.5,"y":0.2,"w":0.4}},
  {"type":"codeblock", "text":"int i = 0;\ni++;","geometry":{"x":0.3,"y":0.5,"w":0.2}},
  {"type":"table", "text":"<full tabular LaTeX>", "geometry":{"x":0.1,"y":0.10,"w":0.7}}
]

EXAMPLE OUTPUT:
\begin{frame}[fragile]
\begin{textblock}{0.2}(0.2, 0.15)
{\scriptsize
\begin{itemize}
\item a
\item b
...
\item k
\end{itemize}
}
\end{textblock}
\begin{textblock}{0.6}(0.1, 0.05)
\textbf{Mein Titel}
\end{textblock}
\begin{textblock}{0.5}(0.2, 0.1)
Hallo!
\end{textblock}
\begin{textblock}{0.4}(0.5, 0.2)
\begin{itemize}
\item Foo
\item Bar
\end{itemize}
\end{textblock}
\begin{textblock}{0.2}(0.3, 0.5)
\begin{lstlisting}
int i = 0;
i++;
\end{lstlisting}
\end{textblock}
\begin{textblock}{0.7}(0.1, 0.10)
\resizebox{\textwidth}{!}{%
<full tabular LaTeX>
}
\end{textblock}
\end{frame}

NO frame title argument, only textblock for title!
Geometry must be used EXACTLY as given.
If these rules are not followed, the output is invalid.
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
        return extract_latex_content(content)
    except Exception as e:
        print(f"Error generating Slide {slide_num}: {e}")
        return f"% ERROR Slide {slide_num}\n\\begin{{frame}}{{Error}}\nGeneration failed.\n\\end{{frame}}"
