from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


INLINE_MATH_PATTERN = re.compile(r"(\\\((.+?)\\\)|\$(?!\$)(.+?)(?<!\\)\$)")
DISPLAY_MATH_PATTERN = re.compile(r"^\s*(?:\\\[(.+?)\\\]|\$\$(.+?)\$\$)\s*$")

LATEX_REPLACEMENTS = {
    r"\leq": "≤",
    r"\geq": "≥",
    r"\neq": "≠",
    r"\ne": "≠",
    r"\pm": "±",
    r"\mp": "∓",
    r"\times": "×",
    r"\cdot": "·",
    r"\div": "÷",
    r"\infty": "∞",
    r"\approx": "≈",
    r"\sim": "∼",
    r"\equiv": "≡",
    r"\parallel": "∥",
    r"\perp": "⊥",
    r"\angle": "∠",
    r"\triangle": "△",
    r"\Delta": "Δ",
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\theta": "θ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\pi": "π",
    r"\varphi": "φ",
    r"\phi": "φ",
    r"\omega": "ω",
    r"\sin": "sin",
    r"\cos": "cos",
    r"\tan": "tan",
    r"\cot": "cot",
    r"\log": "log",
    r"\ln": "ln",
    r"\sqrt": "√",
    r"\vec": "→",
}

SUPERSCRIPT_MAP = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")
SUBSCRIPT_MAP = str.maketrans("0123456789+-=()n", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₙ")


def normalize_latex_math(value: str) -> str:
    text = value.strip()
    text = text.replace("\\left", "").replace("\\right", "")
    text = re.sub(r"\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}", r"(\1)/(\2)", text)
    text = re.sub(r"\\dfrac\s*\{([^{}]+)\}\s*\{([^{}]+)\}", r"(\1)/(\2)", text)
    text = re.sub(r"\\sqrt\s*\{([^{}]+)\}", r"√(\1)", text)
    text = re.sub(r"\\overrightarrow\s*\{([^{}]+)\}", r"\1⃗", text)
    text = re.sub(r"\\vec\s*\{([^{}]+)\}", r"\1⃗", text)

    for source, target in LATEX_REPLACEMENTS.items():
        text = text.replace(source, target)

    text = re.sub(
        r"\^\{([0-9+\-=()n]+)\}",
        lambda match: match.group(1).translate(SUPERSCRIPT_MAP),
        text,
    )
    text = re.sub(
        r"_\{([0-9+\-=()n]+)\}",
        lambda match: match.group(1).translate(SUBSCRIPT_MAP),
        text,
    )
    text = re.sub(r"\^([0-9+\-=()n])", lambda match: match.group(1).translate(SUPERSCRIPT_MAP), text)
    text = re.sub(r"_([0-9+\-=()n])", lambda match: match.group(1).translate(SUBSCRIPT_MAP), text)
    text = text.replace("{", "").replace("}", "")
    text = text.replace("\\", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def contains_math_markup(text: str) -> bool:
    return bool(INLINE_MATH_PATTERN.search(text) or DISPLAY_MATH_PATTERN.search(text))


def _append_text_run(paragraph: Paragraph, text: str) -> None:
    if text:
        paragraph.add_run(text)


def _append_omml_equation(paragraph: Paragraph, latex: str) -> None:
    equation_text = normalize_latex_math(latex)
    o_math = OxmlElement("m:oMath")
    run = OxmlElement("m:r")
    run_properties = OxmlElement("m:rPr")
    normal_text = OxmlElement("m:nor")
    run_properties.append(normal_text)
    text_element = OxmlElement("m:t")
    text_element.set(qn("xml:space"), "preserve")
    text_element.text = equation_text
    run.append(run_properties)
    run.append(text_element)
    o_math.append(run)
    paragraph._p.append(o_math)


def add_math_aware_paragraph(paragraph: Paragraph, text: str) -> None:
    display_match = DISPLAY_MATH_PATTERN.match(text)
    if display_match:
        _append_omml_equation(paragraph, display_match.group(1) or display_match.group(2) or "")
        return

    position = 0
    for match in INLINE_MATH_PATTERN.finditer(text):
        _append_text_run(paragraph, text[position:match.start()])
        _append_omml_equation(paragraph, match.group(2) or match.group(3) or "")
        position = match.end()
    _append_text_run(paragraph, text[position:])
