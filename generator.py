import re


def generate_latex_preamble(metadata, detected_header=None):
    # Standardwerte
    date = metadata.get("date", r"\today")
    raw_title = metadata.get("title", "Presentation")
    title = normalize_title(raw_title)
    author = metadata.get("author", "LaTeX Converter") 
    institute_blob = metadata.get("institute", "") 
    author_line = f"\\author[{author}]{{{author}}}"
    institute_line = f"\\institute{{{institute_blob}}}"


    return rf"""
\documentclass[aspectratio=169]{{beamer}}

\usepackage[utf8]{{inputenc}}
\usepackage[T1]{{fontenc}}
\usepackage[ngerman]{{babel}}
\usepackage{{graphicx}}
\usepackage{{booktabs}}
\usepackage{{listings}}
\usepackage{{xcolor}}
\usepackage{{textgreek}}
\usepackage{{tikz}}
\usepackage[absolute, overlay]{{textpos}}
\usepackage{{hyperref}}
\setlength{{\TPHorizModule}}{{\paperwidth}}
\setlength{{\TPVertModule}}{{\paperheight}}
\TPGrid{{1}}{{1}}
\usetheme{{Madrid}}
\setbeamertemplate{{frametitle}}{{}}
\usecolortheme{{default}}

\definecolor{{codegreen}}{{rgb}}{{0,0.6,0}}
\definecolor{{codegray}}{{rgb}}{{0.5,0.5,0.5}}
\definecolor{{codepurple}}{{rgb}}{{0.58,0,0.82}}
\lstdefinestyle{{mystyle}}{{%
    commentstyle=\color{{codegreen}},
    keywordstyle=\color{{magenta}},
    numberstyle=\tiny\color{{codegray}},
    stringstyle=\color{{codepurple}},
    basicstyle=\ttfamily\footnotesize,
    breakatwhitespace=false,
    breaklines=true,
    captionpos=b,
    keepspaces=true,
    numbers=left,
    numbersep=5pt,
    showspaces=false,
    showstringspaces=false,
    showtabs=false,
    tabsize=2
}}
\lstset{{style=mystyle}}



% --- METADATA ---
\title{{{title}}}
{author_line}
{institute_line}
\date{{{date}}}

\begin{{document}}

\setbeamertemplate{{footline}}{{}} 
\begin{{frame}}
  \titlepage
\end{{frame}}
"""


LATEX_POSTAMBLE = r"""
\end{document}
"""



def normalize_title(title):
    # Ersetze _ durch Leerzeichen
    title = title.replace('_', ' ')
    # Füge Leerzeichen zwischen kleinGroß (CamelCase)
    title = re.sub(r'(?<=[a-zäöüß])(?=[A-ZÄÖÜ])', ' ', title)
    # Optional: Mehrere Leerzeichen auf eines reduzieren
    title = re.sub(r'\s+', ' ', title)
    return title.strip()