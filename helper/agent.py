import json
import ollama
import re
from helper.utils import RED,RESET, repair_latex_output

def cut_latex_content(text: str) -> str:
    """
    Removes Markdown code block wrappers to extract raw LaTeX content.

    Searches for content enclosed in triple backticks (optionally tagged with 'latex') 
    and returns the inner text. If no wrapper is found, returns the original text 
    stripped of leading/trailing whitespace.
    """
    pattern = r"```(?:latex)?\s*(.*?)\s*```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1)
    return text.strip()
       
def load_conversion_rules() -> str:
    """
    Returns the strict system prompt rules for converting JSON slide elements into LaTeX.

    The returned string defines the required LaTeX structure (frames, textblocks), 
    alignment logic per element type, and specific rendering syntax for media, 
    tables, and code blocks.
    """
    return r"""
You are a specialized LaTeX Beamer Generator.
You convert a provided JSON structure of a presentation slide into valid, compilable LaTeX code using the 'textpos' package for absolute positioning.

INPUT DATA:
You receive a JSON object representing a SINGLE slide with a list of "elements".
Each element contains:
- "type": (text, list, codeblock, table, picture, video, header, footer, etc.)
- "geometry": { "x", "y", "w", "h" } (Normalized coordinates 0.0-1.0)
- Content fields: "text", "items", "table_rows", "image_path", "path", etc.

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
     - **"table", "list", "picture", "codeblock", "video"**: ALWAYS use **[t]** (Top).
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
- **"video"**:
     - Generate exactly: 
       `\includemedia[width=\linewidth, height=\textheight, activate=pageopen, addresource=<path>, flashvars={source=<path> &autoPlay=true &loop=true}]{\includegraphics[width=\linewidth,height=\textheight]{<poster_path>}}{VPlayer.swf}`
     - **IMPORTANT:** 1. Use `path` for `addresource` and `flashvars`.
       2. Use `poster_path` inside `\includegraphics`. 
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

Input 2 (List - [b]):
{
  "type": "list",
  "geometry": {"x": 0.1, "y": 0.2, "w": 0.8, "h": 0.6},
  "items": ["Point A", "Point B"],
  "align": "b",
  "fontsize": "small"
}

Output 2:
\begin{textblock}{0.8}(0.1, 0.2)
  % WICHTIG: [b] sorgt hier dafür, dass der Text am unteren Rand der Box klebt
  \begin{minipage}[b][0.6\paperheight]{\linewidth}
    {\small
    \begin{itemize}
      \item Summary Point 1
      \item Summary Point 2
    \end{itemize}
    }
  \end{minipage}
\end{textblock}

Input 3 (List - [t]):
{
  "type": "list",
  "geometry": {"x": 0.1, "y": 0.2, "w": 0.8, "h": 0.6},
  "items": ["Point A", "Point B"],
  "align": "t",
  "fontsize": "small"
}

Output 3:
\begin{textblock}{0.8}(0.1, 0.2)
  \begin{minipage}[t][0.6\paperheight]{\linewidth}
    {\small
    \begin{itemize}
      \item Point A
      \item Point B
    \end{itemize}
    }
  \end{minipage}
\end{textblock}

Input 4 (Title - [t]):
{
  "type": "title",
  "geometry": {"x": 0.1, "y": 0.05, "w": 0.8, "h": 0.1},
  "text": "My Presentation Title"
}

Output 4:
\begin{textblock}{0.8}(0.1, 0.05)
  \begin{minipage}[t][0.1\paperheight]{\linewidth}
    \textbf{My Presentation Title}
  \end{minipage}
\end{textblock}


Input 5 (Table - requires [t] and resizebox):
{
  "type": "table",
  "geometry": {"x": 0.1, "y": 0.3, "w": 0.5, "h": 0.4},
  "table_rows": [["Col1", "Col2"], ["Val1", "Val2"]]
}

Output 5:
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
Input 6 (Footer/Header with fixed 3pt font size):
{
  "type": "footer",
  "geometry": {"x": 0.56, "y": 0.90, "w": 0.23, "h": 0.03},
  "text": "Quelle: University of Washington",
  "fontsize": "3pt"
}

Output 6:
\begin{textblock}{0.23}(0.56, 0.90)
  \begin{minipage}[b][0.03\paperheight]{\linewidth}
    \raggedright
    \fontsize{3}{3.3}\selectfont Quelle: University of Washington
  \end{minipage}
\end{textblock}

Input 7 (Codeblock):
{
  "type": "codeblock",
  "geometry": {"x": 0.1, "y": 0.4, "w": 0.8, "h": 0.3},
  "text": "\\begin{lstlisting}[language=Java]\nfor (i = 0; i < n; i++) {\na[i] = 1;\nb[i] = 2;\n}\n\\end{lstlisting}"
}

Output 7:
\begin{textblock}{0.8}(0.1, 0.4)
  \begin{minipage}[t][0.3\paperheight]{\linewidth}
    \begin{lstlisting}[language=Java, basicstyle=\ttfamily\scriptsize]
for (i = 0; i < n; i++) {
a[i] = 1;
b[i] = 2;
}
    \end{lstlisting}
  \end{minipage}
\end{textblock}

Input 8 (Video with Poster):
{
  "type": "video",
  "geometry": {"x": 0.2, "y": 0.2, "w": 0.6, "h": 0.4},
  "path": "extracted_media/video.mp4",
  "poster_path": "extracted_media/video_poster.png"
}

Output 8:
\begin{textblock}{0.6}(0.2, 0.2)
  \begin{minipage}[t][0.4\paperheight]{\linewidth}
    \includemedia[
       width=\linewidth, 
       height=0.4\paperheight,
       activate=pageopen, 
       addresource=extracted_media/video.mp4, 
       flashvars={
          source=extracted_media/video.mp4 
          &autoPlay=true 
          &loop=true
       }
    ]{\includegraphics[width=\linewidth,height=\textheight]{extracted_media/video_poster.png}}{VPlayer.swf}
  \end{minipage}
\end{textblock}
"""

import json
import ollama

def generate_single_slide_latex(slide_data: dict, config) -> str:
    """
    Generates LaTeX code for a single slide using an LLM.

    Constructs a prompt based on strict conversion rules and the provided JSON slide data.
    Calls the configured LLM model to generate the LaTeX code, then applies post-processing 
    repairs and cleanup to ensure valid output.
    """
    slide_num = slide_data.get('slide_number', '?')
    
    rules_block = load_conversion_rules()

    system_prompt = (
        "You are a strictly constrained LaTeX Beamer generator. "
        "You do not explain. You only output code."
    )

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
        return cut_latex_content(content)
        
    except Exception as e:
        print(f"{RED}Error generating Slide {slide_num}: {e}{RESET}")
        return f"% ERROR Slide {slide_num}\n\\begin{{frame}}{{Error}}\nGeneration failed.\n\\end{{frame}}"