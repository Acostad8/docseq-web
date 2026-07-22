"""
DocSeq — Numeración automática para documentos Word
====================================================
Interfaz gráfica para numerar ilustraciones y tablas en documentos .docx
mediante campos SEQ de Word.

Uso:
    python docseq.py            ← Abre la interfaz gráfica
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.app import launch

if __name__ == "__main__":
    launch()
