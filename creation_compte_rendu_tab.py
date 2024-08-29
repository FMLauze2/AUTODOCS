import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Inches
from datetime import datetime
import os

def create_creation_compte_rendu_tab(tab_creation_compte_rendu):
    # Créer un canvas pour le scroll
    canvas = tk.Canvas(tab_creation_compte_rendu)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Ajouter un scrollbar vertical lié au canvas
    scrollbar = tk.Scrollbar(tab_creation_compte_rendu, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Créer un frame à l'intérieur du canvas pour y ajouter les widgets
    scrollable_frame = tk.Frame(canvas)

    # Configurer le canvas pour utiliser le scrollbar
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # S'assurer que le canvas réagit à la molette de la souris
    scrollable_frame.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

    # Ajouter les widgets dans scrollable_frame
    label_cr = tk.Label(scrollable_frame, text="Création de Compte Rendu")
    label_cr.pack(padx=10, pady=10)

    # Augmenter la largeur des widgets pour correspondre à la nouvelle taille
    label_titre = tk.Label(scrollable_frame, text="Titre du compte rendu :")
    label_titre.pack(padx=10, pady=5, anchor="w")
    entry_titre = tk.Entry(scrollable_frame, width=70)
    entry_titre.pack(padx=10, pady=5)

    label_date = tk.Label(scrollable_frame, text="Date et heure :")
    label_date.pack(padx=10, pady=5, anchor="w")
    entry_date = tk.Entry(scrollable_frame, width=70)
    entry_date.insert(0, datetime.now().strftime("%Y-%m-%d %H:%M"))
    entry_date.pack(padx=10, pady=5)

    label_participants = tk.Label(scrollable_frame, text="Participants :")
    label_participants.pack(padx=10, pady=5, anchor="w")
    entry_participants = tk.Entry(scrollable_frame, width=70)
    entry_participants.pack(padx=10, pady=5)

    label_contact = tk.Label(scrollable_frame, text="Infos de contact (client ou correspondant) :")
    label_contact.pack(padx=10, pady=5, anchor="w")
    entry_contact = tk.Entry(scrollable_frame, width=70)
    entry_contact.pack(padx=10, pady=5)

    label_objet = tk.Label(scrollable_frame, text="Objet / Sujet :")
    label_objet.pack(padx=10, pady=5, anchor="w")
    entry_objet = tk.Entry(scrollable_frame, width=70)
    entry_objet.pack(padx=10, pady=5)

    label_notes = tk.Label(scrollable_frame, text="Notes / Détails :")
    label_notes.pack(padx=10, pady=5, anchor="w")
    text_notes = tk.Text(scrollable_frame, wrap='word', height=10, width=70)
    text_notes.pack(padx=10, pady=5)

    label_actions = tk.Label(scrollable_frame, text="Actions à suivre :")
    label_actions.pack(padx=10, pady=5, anchor="w")
    text_actions = tk.Text(scrollable_frame, wrap='word', height=5, width=70)
    text_actions.pack(padx=10, pady=5)

    label_images = tk.Label(scrollable_frame, text="Images (Screenshots) :")
    label_images.pack(padx=10, pady=5, anchor="w")
    listbox_images = tk.Listbox(scrollable_frame, height=5, width=70, selectmode=tk.SINGLE)
    listbox_images.pack(padx=10, pady=5)

    def ajouter_image():
        fichiers = filedialog.askopenfilenames(title="Sélectionnez des images", filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp")])
        for fichier in fichiers:
            listbox_images.insert(tk.END, fichier)

    btn_ajouter_image = tk.Button(scrollable_frame, text="Ajouter Image", command=ajouter_image)
    btn_ajouter_image.pack(padx=10, pady=5)

    def generer_document():
        titre = entry_titre.get().strip()
        date_heure = entry_date.get().strip()
        participants = entry_participants.get().strip()
        contact_info = entry_contact.get().strip()
        objet = entry_objet.get().strip()
        notes = text_notes.get("1.0", tk.END).strip()
        actions = text_actions.get("1.0", tk.END).strip()

        if not titre or not date_heure:
            messagebox.showerror("Erreur", "Le titre et la date/heure sont obligatoires.")
            return

        doc = Document()
        doc.add_heading(titre, 0)

        doc.add_paragraph(f"Date et Heure : {date_heure}")
        if participants:
            doc.add_paragraph(f"Participants : {participants}")
        if contact_info:
            doc.add_paragraph(f"Contact : {contact_info}")
        if objet:
            doc.add_paragraph(f"Objet / Sujet : {objet}")

        if notes:
            doc.add_heading("Notes / Détails", level=1)
            doc.add_paragraph(notes)

        if actions:
            doc.add_heading("Actions à suivre", level=1)
            doc.add_paragraph(actions)

        if listbox_images.size() > 0:
            doc.add_heading("Images", level=1)
            for image_path in listbox_images.get(0, tk.END):
                doc.add_picture(image_path, width=Inches(5.0))

        fichier = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Document Word", "*.docx")])
        if fichier:
            doc.save(fichier)
            messagebox.showinfo("Succès", "Le compte rendu a été généré avec succès.")
            os.startfile(fichier)

    btn_generer = tk.Button(scrollable_frame, text="Générer et Ouvrir Document Word", command=generer_document)
    btn_generer.pack(padx=10, pady=20)

# Dans le fichier principal `main.py`, configurez la taille de la fenêtre principale

