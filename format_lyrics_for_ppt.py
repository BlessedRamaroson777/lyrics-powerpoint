# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR
import os

def get_unique_filename(base="paroles", ext=".pptx"):
    counter = 0
    while True:
        filename = f"{base}{ext}" if counter == 0 else f"{base}{counter}{ext}"
        if not os.path.exists(filename):
            return filename
        counter += 1

def split_block_recursive(block, max_lines=12):
    """
    Divise récursivement un bloc de lignes jusqu'à ce que
    chaque bloc contienne max_lines ou moins.
    """
    if len(block) <= max_lines:
        return [block]

    mid = len(block) // 2
    part1 = block[:mid]
    part2 = block[mid:]
    return split_block_recursive(part1, max_lines) + split_block_recursive(part2, max_lines)

# Lis les paroles depuis un fichier texte
with open("paroles.txt", "r", encoding="utf-8") as f:
    lines = f.readlines()

prs = Presentation()
block = []

for line in lines:
    line = line.strip()
    if line == "":
        if block:
            small_blocks = split_block_recursive(block, max_lines=12)
            for b in small_blocks:
                slide = prs.slides.add_slide(prs.slide_layouts[1])  # Titre + contenu
                # Mettre le texte dans le contenu, pas le titre
                content_placeholder = slide.placeholders[1]
                content_placeholder.text = "\n".join(b)
                content_placeholder.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            block = []
    else:
        block.append(line)

# Dernier bloc
if block:
    small_blocks = split_block_recursive(block, max_lines=12)
    for b in small_blocks:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        content_placeholder = slide.placeholders[1]
        content_placeholder.text = "\n".join(b)
        content_placeholder.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

# Nom de fichier unique
filename = get_unique_filename("paroles", ".pptx")
prs.save(filename)
print(f"PowerPoint créé : {filename}")