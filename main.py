import tkinter as tk
import sys, os, json
from tkinter import ttk
from utilisation_tab import create_utilisation_tab
from creation_compte_rendu_tab import create_creation_compte_rendu_tab
from options_tab import create_options_tab

def main():
    # Créer l'interface graphique avec Tkinter
    root = tk.Tk()
    root.title("Gestionnaire de Réponses Email")

    # Définir la taille initiale de la fenêtre
    root.geometry("1024x768")

    # Utiliser un Notebook pour les onglets
    notebook = ttk.Notebook(root)
    notebook.pack(padx=10, pady=10, fill='both', expand=True)

    # Créer l'onglet "Utilisation"
    tab_main = tk.Frame(notebook)
    create_utilisation_tab(tab_main)
    notebook.add(tab_main, text="Utilisation")

    # Créer l'onglet "Création compte rendu"
    tab_creation_compte_rendu = tk.Frame(notebook)
    create_creation_compte_rendu_tab(tab_creation_compte_rendu)
    notebook.add(tab_creation_compte_rendu, text="Création compte rendu")

    # Créer l'onglet "Options"
    tab_options = tk.Frame(notebook)
    create_options_tab(tab_options)
    notebook.add(tab_options, text="Options")

    # Lancer l'application
    root.mainloop()

if __name__ == "__main__":
    main()
