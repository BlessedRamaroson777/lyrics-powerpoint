# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR

# ---------- Fonctions Utilitaires ----------
def get_unique_filename(base="paroles", ext=".pptx"):
    counter = 0
    while True:
        filename = f"{base}{ext}" if counter == 0 else f"{base}{counter}{ext}"
        if not os.path.exists(filename):
            return filename
        counter += 1

def split_block_recursive(block, max_lines=12):
    if len(block) <= max_lines:
        return [block]
    mid = len(block) // 2
    return split_block_recursive(block[:mid], max_lines) + split_block_recursive(block[mid:], max_lines)

def generate_pptx_from_lines(lines, max_lines=12):
    prs = Presentation()
    block = []

    for line in lines:
        line = line.strip()
        if line == "":
            if block:
                small_blocks = split_block_recursive(block, max_lines)
                for b in small_blocks:
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    content_placeholder = slide.placeholders[1]
                    content_placeholder.text = "\n".join(b)
                    content_placeholder.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                block = []
        else:
            block.append(line)

    if block:
        small_blocks = split_block_recursive(block, max_lines)
        for b in small_blocks:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            content_placeholder = slide.placeholders[1]
            content_placeholder.text = "\n".join(b)
            content_placeholder.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    filename = get_unique_filename("paroles", ".pptx")
    prs.save(filename)
    return filename

# ---------- Actions ----------
def choose_file():
    filepath = filedialog.askopenfilename(
        title="Sélectionner un fichier de paroles",
        filetypes=[("Fichiers texte", "*.txt")]
    )
    if filepath:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                lines = f.readlines()
            output = generate_pptx_from_lines(lines, max_lines=12)
            messagebox.showinfo("Succès", f"PowerPoint créé :\n{output}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

def generate_from_text():
    text = text_box.get("1.0", tk.END).strip()
    if not text:
        messagebox.showwarning("Attention", "Veuillez coller ou écrire du texte.")
        return
    lines = text.split("\n")
    try:
        output = generate_pptx_from_lines(lines, max_lines=12)
        messagebox.showinfo("Succès", f"PowerPoint créé :\n{output}")
    except Exception as e:
        messagebox.showerror("Erreur", str(e))

# ---------- Interface ----------
root = tk.Tk()
root.title("🎤 Créateur de PowerPoint de Paroles")
root.geometry("600x500")
root.configure(bg="#f8c8dc")  # Rose pastel

# Style moderne
style = ttk.Style()
style.configure("TButton", font=("Arial", 12, "bold"), padding=10)

# Titre
label = tk.Label(root, text="🎶 Générateur de Paroles en PowerPoint 🎶",
                 font=("Arial", 16, "bold"), bg="#f8c8dc", fg="#5a004f")
label.pack(pady=15)

# Bouton pour choisir un fichier
btn_file = ttk.Button(root, text="📂 Sélectionner un fichier .txt", command=choose_file)
btn_file.pack(pady=10)

# Zone de texte
label2 = tk.Label(root, text="Ou collez vos paroles ci-dessous :", 
                  font=("Arial", 12), bg="#f8c8dc", fg="#3a0033")
label2.pack(pady=5)

text_box = tk.Text(root, wrap="word", width=60, height=15, font=("Arial", 11))
text_box.pack(padx=20, pady=10)

# Bouton pour générer depuis la zone de texte
btn_text = ttk.Button(root, text="✨ Générer à partir du texte collé", command=generate_from_text)
btn_text.pack(pady=15)

# Lancement
root.mainloop()
