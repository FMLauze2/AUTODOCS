import tkinter as tk
import json
import os
import sys
import pyperclip  # Pour copier dans le presse-papiers

# Chemin absolu pour les fichiers JSON et signature
chemin_reponses = os.path.join(os.path.dirname(__file__), 'reponses.json')
chemin_signature = os.path.join(os.path.dirname(__file__), 'signature.txt')

def resource_path(relative_path):
    """Obtenir le chemin absolu vers la ressource, fonctionne pour le développement et pour PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def charger_reponses():
    chemin_reponses = resource_path('reponses.json')
    with open(chemin_reponses, 'r', encoding='utf-8') as f:
        return json.load(f)
    
def create_utilisation_tab(tab_main):
    reponses = charger_reponses()

    # Charger la signature depuis un fichier texte
    with open(chemin_signature, 'r', encoding='utf-8') as f:
        signature = f.read()

    label_nom = tk.Label(tab_main, text="Nom du client :")
    label_nom.grid(row=0, column=0, padx=5, pady=5, sticky="w")
    entry_nom = tk.Entry(tab_main, width=30)
    entry_nom.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    frame_buttons = tk.Frame(tab_main)
    frame_buttons.grid(row=1, column=0, padx=10, pady=10, sticky="nw")

    frame_viewer = tk.Frame(tab_main)
    frame_viewer.grid(row=1, column=1, padx=10, pady=10, sticky="ne")

    text_viewer = tk.Text(frame_viewer, wrap='word', height=20, width=50)
    text_viewer.pack()

    def afficher_reponse(reponse):
        nom_client = entry_nom.get()
        if nom_client:
            reponse = reponse.replace("[NomClient]", nom_client)
        reponse += "\n\n" + signature  # Ajout de la signature à la fin de la réponse
        text_viewer.delete(1.0, tk.END)
        text_viewer.insert(tk.END, reponse)

    def copier_texte_visionneuse():
        texte = text_viewer.get("1.0", tk.END)
        pyperclip.copy(texte)
        tk.messagebox.showinfo("Succès", "Texte copié dans le presse-papiers.")

    def rafraichir_boutons():
        for widget in frame_buttons.winfo_children():
            widget.destroy()

        nonlocal reponses
        reponses = charger_reponses()

        for titre, contenu in reponses.items():
            button = tk.Button(frame_buttons, text=titre, width=20, command=lambda c=contenu: afficher_reponse(c))
            button.pack(pady=5)

    # Bouton pour actualiser les réponses
    btn_actualiser = tk.Button(tab_main, text="Actualiser", command=rafraichir_boutons)
    btn_actualiser.grid(row=2, column=0, padx=5, pady=5, sticky="w")

    # Bouton pour copier le texte de la visionneuse
    btn_copier = tk.Button(tab_main, text="Copier", command=copier_texte_visionneuse)
    btn_copier.grid(row=2, column=1, padx=5, pady=5, sticky="e")

    rafraichir_boutons()
