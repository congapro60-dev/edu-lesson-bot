from __future__ import annotations

import re
from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


INLINE_MATH_PATTERN = re.compile(r"(\\\((.+?)\\\)|\$(?!\$)(.+?)(?<!\\)\$)")
DISPLAY_MATH_PATTERN = re.compile(r"^\s*(?:\\\[(.+?)\\\]|\$\$(.+?)\$\$)\s*$")
COMMAND_PATTERN = re.compile(r"\\([A-Za-z]+)")

LATEX_REPLACEMENTS = {
    r"\leq": "≤",
    r"\le": "≤",
    r"\geq": "≥",
    r"\ge": "≥",
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
    r"\sum": "∑",
    r"\int": "∫",
    r"\lim": "lim",
    r"\sin": "sin",
    r"\cos": "cos",
    r"\tan": "tan",
    r"\cot": "cot",
    r"\log": "log",
    r"\ln": "ln",
    r"\mid": "∣",
    r"\to": "→",
    r"\rightarrow": "→",
    r"\Leftarrow": "⇐",
    r"\Rightarrow": "⇒",
    r"\Leftrightarrow": "⇔",
}


SCRIPT_COMMANDS = {"frac", "dfrac", "tfrac", "sqrt", "vec", "overrightarrow"}


def contains_math_markup(text: str) -> bool:
    return bool(INLINE_MATH_PATTERN.search(text) or DISPLAY_MATH_PATTERN.search(text))


def _math_run(text: str) -> OxmlElement:
    run = OxmlElement("m:r")
    run_properties = OxmlElement("m:rPr")
    normal_text = OxmlElement("m:nor")
    run_properties.append(normal_text)
    text_element = OxmlElement("m:t")
    text_element.set(qn("xml:space"), "preserve")
    text_element.text = text
    run.append(run_properties)
    run.append(text_element)
    return run


def _container(tag: str, nodes: list[OxmlElement] | None = None) -> OxmlElement:
    element = OxmlElement(tag)
    for node in nodes or []:
        element.append(node)
    return element


def _append_nodes(parent: OxmlElement, nodes: list[OxmlElement]) -> None:
    for node in nodes:
        parent.append(deepcopy(node))


def _fraction(num_nodes: list[OxmlElement], den_nodes: list[OxmlElement]) -> OxmlElement:
    frac = OxmlElement("m:f")
    frac_properties = OxmlElement("m:fPr")
    frac.append(frac_properties)
    numerator = OxmlElement("m:num")
    denominator = OxmlElement("m:den")
    _append_nodes(numerator, num_nodes)
    _append_nodes(denominator, den_nodes)
    frac.append(numerator)
    frac.append(denominator)
    return frac


def _radical(nodes: list[OxmlElement]) -> OxmlElement:
    radical = OxmlElement("m:rad")
    radical.append(OxmlElement("m:radPr"))
    radical.append(OxmlElement("m:deg"))
    expression = OxmlElement("m:e")
    _append_nodes(expression, nodes)
    radical.append(expression)
    return radical


def _script(base: OxmlElement, sup_nodes: list[OxmlElement] | None, sub_nodes: list[OxmlElement] | None) -> OxmlElement:
    if sup_nodes and sub_nodes:
        element = OxmlElement("m:sSubSup")
        base_container = OxmlElement("m:e")
        base_container.append(base)
        sub_container = OxmlElement("m:sub")
        sup_container = OxmlElement("m:sup")
        _append_nodes(sub_container, sub_nodes)
        _append_nodes(sup_container, sup_nodes)
        element.append(base_container)
        element.append(sub_container)
        element.append(sup_container)
        return element

    if sup_nodes:
        element = OxmlElement("m:sSup")
        base_container = OxmlElement("m:e")
        sup_container = OxmlElement("m:sup")
        base_container.append(base)
        _append_nodes(sup_container, sup_nodes)
        element.append(base_container)
        element.append(sup_container)
        return element

    if sub_nodes:
        element = OxmlElement("m:sSub")
        base_container = OxmlElement("m:e")
        sub_container = OxmlElement("m:sub")
        base_container.append(base)
        _append_nodes(sub_container, sub_nodes)
        element.append(base_container)
        element.append(sub_container)
        return element

    return base


def _read_braced_group(text: str, position: int) -> tuple[str, int]:
    while position < len(text) and text[position].isspace():
        position += 1
    if position >= len(text):
        return "", position
    if text[position] != "{":
        command_match = COMMAND_PATTERN.match(text, position)
        if command_match:
            return text[position:command_match.end()], command_match.end()
        return text[position], position + 1

    depth = 0
    start = position + 1
    for index in range(position, len(text)):
        char = text[index]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start:index], index + 1
    return text[start:], len(text)


def _command_to_text(command: str) -> str:
    source = f"\\{command}"
    if source in LATEX_REPLACEMENTS:
        return LATEX_REPLACEMENTS[source]
    return command


def _read_atom(text: str, position: int) -> tuple[OxmlElement | None, int]:
    char = text[position]
    if char.isspace():
        return _math_run(" "), position + 1
    if char in "{}":
        return None, position + 1

    command_match = COMMAND_PATTERN.match(text, position)
    if command_match:
        command = command_match.group(1)
        position = command_match.end()
        if command in {"left", "right"}:
            return None, position
        if command in {"frac", "dfrac", "tfrac"}:
            numerator, position = _read_braced_group(text, position)
            denominator, position = _read_braced_group(text, position)
            return _fraction(_parse_math_nodes(numerator), _parse_math_nodes(denominator)), position
        if command == "sqrt":
            radicand, position = _read_braced_group(text, position)
            return _radical(_parse_math_nodes(radicand)), position
        if command in {"vec", "overrightarrow"}:
            content, position = _read_braced_group(text, position)
            return _math_run(f"{_plain_math_text(content)}⃗"), position
        return _math_run(_command_to_text(command)), position

    return _math_run(char), position + 1


def _parse_script_nodes(text: str, position: int) -> tuple[list[OxmlElement], int]:
    content, position = _read_braced_group(text, position)
    return _parse_math_nodes(content), position


def _parse_math_nodes(text: str) -> list[OxmlElement]:
    nodes: list[OxmlElement] = []
    position = 0
    while position < len(text):
        atom, position = _read_atom(text, position)
        if atom is None:
            continue

        sup_nodes: list[OxmlElement] | None = None
        sub_nodes: list[OxmlElement] | None = None
        while position < len(text) and text[position] in "^_":
            marker = text[position]
            position += 1
            script_nodes, position = _parse_script_nodes(text, position)
            if marker == "^":
                sup_nodes = script_nodes
            else:
                sub_nodes = script_nodes

        nodes.append(_script(atom, sup_nodes, sub_nodes))
    return nodes


def _plain_math_text(value: str) -> str:
    text = value.strip().replace("\\left", "").replace("\\right", "")
    for source, target in LATEX_REPLACEMENTS.items():
        text = text.replace(source, target)
    text = re.sub(r"\\(?:frac|dfrac|tfrac)\s*\{([^{}]+)\}\s*\{([^{}]+)\}", r"(\1)/(\2)", text)
    text = re.sub(r"\\sqrt\s*\{([^{}]+)\}", r"√(\1)", text)
    text = re.sub(r"\\(?:vec|overrightarrow)\s*\{([^{}]+)\}", r"\1⃗", text)
    text = text.replace("{", "").replace("}", "").replace("\\", "")
    return re.sub(r"\s+", " ", text).strip()


def normalize_latex_math(value: str) -> str:
    """Return a readable fallback text for logs/tests; DOCX uses OMML nodes."""
    return _plain_math_text(value)


def _append_text_run(paragraph: Paragraph, text: str) -> None:
    if text:
        paragraph.add_run(text)


def _append_omml_equation(paragraph: Paragraph, latex: str) -> None:
    o_math = OxmlElement("m:oMath")
    nodes = _parse_math_nodes(latex.strip()) or [_math_run(_plain_math_text(latex))]
    _append_nodes(o_math, nodes)
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
