# TXToGIFT

Convierte exámenes en formato Word (`.docx`) al formato **GIFT** de Moodle, listos para importar como banco de preguntas.

## Tipos de pregunta soportados

- **Opción múltiple** con opciones etiquetadas (`A. …`, `B. …`, …, hasta `Z`).
- **Opción múltiple** con opciones sin etiqueta (una por párrafo).
- **Verdadero / Falso** (detectadas automáticamente cuando las opciones son `Verdadero` y `Falso`).
- **Emparejamiento** con flechas `→` o `->`.

Cada pregunta del `.docx` debe terminar con una línea `ANSWER: <letra>` (por ejemplo `ANSWER: C`) que indica la respuesta correcta.

## Instalación

### Linux / macOS / Git Bash

```bash
./install.sh
```

### Windows (CMD o PowerShell)

```bat
install.bat
```

Ambos scripts instalan el paquete en modo editable (`pip install -e .`) junto con su dependencia `python-docx`, y registran el comando global `txt-to-gift`. Requiere Python 3.9 o superior.

## Uso

Una vez instalado, puedes ejecutar el comando desde cualquier directorio:

```bash
txt-to-gift "Examen Leccion 3.docx"
```

Esto genera `Examen Leccion 3.txt` en el mismo directorio con el contenido en formato GIFT.

### Especificar archivo de salida

```bash
txt-to-gift "Examen Leccion 3.docx" salida.txt
```

### Opciones sin letra: ajustar cantidad de alternativas

Si tu `.docx` usa opciones sin letra (solo el texto) y el número de opciones por pregunta no es 4, usa el flag `--options`:

```bash
txt-to-gift --options 5 "Examen.docx"
```

Las preguntas Verdadero/Falso se detectan automáticamente sin importar este valor.

### Windows: caracteres Unicode en la salida de consola

Si ves un `UnicodeEncodeError` al ejecutar desde `cmd.exe`, exporta la codificación antes:

```bat
set PYTHONIOENCODING=utf-8
txt-to-gift "Examen Leccion 3.docx"
```

## Formato esperado del `.docx`

```
Texto de la pregunta (puede ocupar varios párrafos).
A. Primera opción
B. Segunda opción
C. Tercera opción
D. Cuarta opción
ANSWER: C
```

Para Verdadero/Falso:

```
Enunciado de la afirmación.
Verdadero
Falso
ANSWER: B
```

Para emparejamiento:

```
Relaciona cada elemento con su definición.
Concepto A → Definición A
Concepto B → Definición B
ANSWER:
```

## Salida GIFT

El archivo de salida puede importarse directamente en Moodle:
**Administración del curso → Banco de preguntas → Importar → Formato GIFT**.

## Desarrollo

El código está en `docx_to_gift.py`. Las pruebas básicas se encuentran en `test_conversion.py`:

```bash
python test_conversion.py
```
