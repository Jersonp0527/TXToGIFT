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

OPTION_PATTERN = re.compile(r'^([A-Z])[\.\)]\s+(.+)', re.DOTALL)
ANSWER_PATTERN = re.compile(r'^ANSWER\s*:\s*([A-Z]?)\s*$', re.IGNORECASE)
MATCHING_ARROW = re.compile(r'\s*(?:→|->)\s*')  # acepta → (unicode) y ->
BULLET_PATTERN = re.compile(r'^[\u2022\u25E6\u25AA\u25CF\u00B7]\s*')  # • ◦ ▪ ● ·
GIFT_ESCAPE_RE = re.compile(r'([\\~=#{}])')


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


def _is_bullet(line: str) -> bool:
    return bool(BULLET_PATTERN.match(line))


def _strip_bullet(line: str) -> str:
    return BULLET_PATTERN.sub('', line).strip()


def _gift_escape(text: str) -> str:
    """Escapa los caracteres especiales de GIFT: \\ ~ = # { }."""
    return GIFT_ESCAPE_RE.sub(r'\\\1', text)


def _format_question_text(text_lines: list[str]) -> tuple[str, bool]:
    """
    Devuelve (texto_gift, es_html).

    Si alguno de los párrafos empieza con viñeta (•, ◦, ▪…), se emite HTML con
    lista <ul><li> para que Moodle renderice el enunciado en varias líneas en
    lugar de colapsar los ítems en una sola.
    """
    if not any(_is_bullet(ln) for ln in text_lines):
        return _gift_escape(' '.join(text_lines)), False

    parts: list[str] = []
    in_list = False
    for ln in text_lines:
        if _is_bullet(ln):
            if not in_list:
                parts.append('<ul>')
                in_list = True
            parts.append(f'<li>{_gift_escape(_strip_bullet(ln))}</li>')
        else:
            if in_list:
                parts.append('</ul>')
                in_list = False
            parts.append(f'<p>{_gift_escape(ln)}</p>')
    if in_list:
        parts.append('</ul>')
    return '[html]' + ''.join(parts), True


# ---------------------------------------------------------------------------
# Parser del .docx
# ---------------------------------------------------------------------------

def _resolve_block(block_lines: list[str], answer: str, default_options: int) -> dict | None:
    """
    Convierte un bloque de líneas (entre dos ANSWER) en una pregunta.

    Soporta dos modos:
      - Lettered: si alguna línea matchea `A. …`, `B. …`, etc. Esas líneas son
        opciones; lo anterior es el texto de la pregunta.
      - Unlettered: ninguna línea tiene letra. Las últimas N líneas son opciones
        (N=2 si son Verdadero/Falso, si no `default_options`); el resto es el
        texto de la pregunta.
    """
    if not block_lines:
        return None

    lettered_idx = [i for i, ln in enumerate(block_lines) if _is_option(ln)]

    if lettered_idx:
        first_opt = lettered_idx[0]
        text_lines = block_lines[:first_opt]
        options = [_parse_option(block_lines[i]) for i in lettered_idx]
    else:
        n_options = default_options
        if len(block_lines) >= 2:
            last_two = {block_lines[-2].lower().rstrip('.'), block_lines[-1].lower().rstrip('.')}
            if last_two == {'verdadero', 'falso'}:
                n_options = 2
        if len(block_lines) <= n_options:
            return None
        text_lines = block_lines[:-n_options]
        options = [ln.rstrip('.').strip() if ln.lower().rstrip('.') in ('verdadero', 'falso') else ln
                   for ln in block_lines[-n_options:]]

    if not text_lines or not options:
        return None

    text, _ = _format_question_text(text_lines)
    return {
        'text': text,
        'options': options,
        'answer': answer,
    }


def parse_docx(filepath: str, default_options: int = 4) -> list[dict]:
    """
    Lee el .docx y devuelve una lista de preguntas con la estructura:
        { 'text': str, 'options': [str, ...], 'answer': str }
    """
    doc = Document(filepath)
    questions: list[dict] = []
    block_lines: list[str] = []

    for para in doc.paragraphs:
        line = para.text.strip()
        if not line:
            continue

        if _is_answer(line):
            q = _resolve_block(block_lines, _answer_letter(line), default_options)
            if q is not None:
                questions.append(q)
            else:
                preview = ' '.join(block_lines)[:60]
                print(
                    f"  [ADVERTENCIA] Bloque descartado (no se pudo resolver): {preview}…",
                    file=sys.stderr,
                )
            block_lines = []
        else:
            block_lines.append(line)

    if block_lines:
        print(
            f"  [ADVERTENCIA] Líneas sobrantes sin ANSWER descartadas: "
            f"{' '.join(block_lines)[:60]}…",
            file=sys.stderr,
        )

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
    args = [a for a in sys.argv[1:]]
    default_options = 4
    if '--options' in args:
        i = args.index('--options')
        default_options = int(args[i + 1])
        del args[i:i + 2]

    if not args:
        print("Uso: python3 docx_to_gift.py [--options N] input.docx [output.txt]")
        sys.exit(1)

    input_path = args[0]
    output_path = args[1] if len(args) >= 2 else (
        re.sub(r'\.docx$', '', input_path, flags=re.IGNORECASE) + '.txt'
    )

    print(f"Leyendo: {input_path}")
    questions = parse_docx(input_path, default_options=default_options)
    print(f"  → {len(questions)} pregunta(s) encontrada(s)")

    gift_text = convert_to_gift(questions)

    with open(output_path, 'w', encoding='utf-8') as fh:
        fh.write(gift_text)

    print(f"Archivo GIFT generado: {output_path}")


if __name__ == '__main__':
    main()
