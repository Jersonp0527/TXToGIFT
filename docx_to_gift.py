#!/usr/bin/env python3
"""
Convierte un archivo .docx con preguntas de quiz al formato GIFT de Moodle.

Tipos de pregunta soportados:
  - Selección múltiple (una respuesta correcta)
  - Emparejamiento (matching)
  - Verdadero / Falso

Uso:
    python3 docx_to_gift.py input.docx
    python3 docx_to_gift.py input.docx output.txt
"""

import re
import sys
from docx import Document


# ---------------------------------------------------------------------------
# Helpers de detección de líneas
# ---------------------------------------------------------------------------

OPTION_PATTERN = re.compile(r'^([A-D])[\.\)]\s+(.+)', re.DOTALL)
ANSWER_PATTERN = re.compile(r'^ANSWER\s*:\s*([A-D]?)\s*$', re.IGNORECASE)
MATCHING_ARROW = re.compile(r'\s*(?:→|->)\s*')  # acepta → (unicode) y ->


def _is_option(text: str) -> bool:
    return bool(OPTION_PATTERN.match(text))


def _is_answer(text: str) -> bool:
    return bool(ANSWER_PATTERN.match(text))


def _parse_option(text: str) -> str:
    """Devuelve el texto de la opción sin la letra inicial (A., B., …)."""
    m = OPTION_PATTERN.match(text)
    return m.group(2).strip() if m else text.strip()


def _answer_letter(text: str) -> str:
    """Extrae la letra de respuesta (puede ser vacía para emparejamiento)."""
    m = ANSWER_PATTERN.match(text)
    return m.group(1).strip() if m else ''


def _is_matching_option(text: str) -> bool:
    return bool(MATCHING_ARROW.search(text))


def _normalize_arrow(text: str) -> str:
    """Unifica → y -> en la flecha estándar de GIFT: ' -> '."""
    return MATCHING_ARROW.sub(' -> ', text)


# ---------------------------------------------------------------------------
# Parser del .docx
# ---------------------------------------------------------------------------

def parse_docx(filepath: str) -> list[dict]:
    """
    Lee el .docx y devuelve una lista de preguntas con la estructura:
        { 'text': str, 'options': [str, ...], 'answer': str }
    """
    doc = Document(filepath)
    questions: list[dict] = []

    current_text: str | None = None
    current_options: list[str] = []

    def _flush(answer: str):
        nonlocal current_text, current_options
        if current_text and current_options:
            questions.append({
                'text': current_text,
                'options': current_options,
                'answer': answer,
            })
        current_text = None
        current_options = []

    for para in doc.paragraphs:
        line = para.text.strip()
        if not line:
            continue

        if _is_answer(line):
            _flush(_answer_letter(line))
        elif _is_option(line):
            current_options.append(_parse_option(line))
        else:
            # Texto de pregunta: si ya había una pregunta abierta, concatenar
            if current_text is None:
                current_text = line
            else:
                current_text += ' ' + line

    return questions


# ---------------------------------------------------------------------------
# Detección del tipo de pregunta
# ---------------------------------------------------------------------------

def _is_true_false(options: list[str]) -> bool:
    if len(options) != 2:
        return False
    lowered = {o.lower() for o in options}
    return 'verdadero' in lowered and 'falso' in lowered


def _is_matching(options: list[str], answer: str) -> bool:
    return (not answer) and all(_is_matching_option(o) for o in options)


# ---------------------------------------------------------------------------
# Conversión a GIFT
# ---------------------------------------------------------------------------

def _gift_true_false(question: str, options: list[str], answer: str) -> str:
    # answer='A' → Verdadero → TRUE ; answer='B' → Falso → FALSE
    is_true = answer.upper() == 'A'
    value = 'TRUE' if is_true else 'FALSE'
    return f"{question} {{{value}}}"


def _gift_matching(question: str, options: list[str]) -> str:
    lines = [f"{question} {{"]
    for opt in options:
        lines.append(f"    ={_normalize_arrow(opt)}")
    lines.append("}")
    return '\n'.join(lines)


def _gift_multiple_choice(question: str, options: list[str], answer: str) -> str:
    correct_idx = ord(answer.upper()) - ord('A')
    lines = [f"{question} {{"]
    for i, opt in enumerate(options):
        prefix = '=' if i == correct_idx else '~'
        lines.append(f"    {prefix}{opt}")
    lines.append("}")
    return '\n'.join(lines)


def convert_to_gift(questions: list[dict]) -> str:
    blocks: list[str] = []

    for q in questions:
        text = q['text']
        opts = q['options']
        ans  = q['answer']

        if _is_true_false(opts):
            block = _gift_true_false(text, opts, ans)
        elif _is_matching(opts, ans):
            block = _gift_matching(text, opts)
        elif ans:
            block = _gift_multiple_choice(text, opts, ans)
        else:
            # Pregunta sin respuesta definida: se omite con advertencia
            print(f"  [ADVERTENCIA] Pregunta omitida (sin respuesta): {text[:60]}…",
                  file=sys.stderr)
            continue

        blocks.append(block)

    return '\n\n'.join(blocks) + '\n'


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Uso: python3 docx_to_gift.py input.docx [output.txt]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) >= 3 else (
        re.sub(r'\.docx$', '', input_path, flags=re.IGNORECASE) + '.txt'
    )

    print(f"Leyendo: {input_path}")
    questions = parse_docx(input_path)
    print(f"  → {len(questions)} pregunta(s) encontrada(s)")

    gift_text = convert_to_gift(questions)

    with open(output_path, 'w', encoding='utf-8') as fh:
        fh.write(gift_text)

    print(f"Archivo GIFT generado: {output_path}")


if __name__ == '__main__':
    main()
