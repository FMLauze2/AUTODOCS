import tkinter as tk
from tkinter import ttk
from utilisation_tab import create_utilisation_tab
from creation_compte_rendu_tab import create_creation_compte_rendu_tab
from options_tab import create_options_tab
from creation_ticket_installation_tab import create_creation_ticket_installation_tab
from creation_contrat_services_tab import create_contrat_services_tab
from liste_contrats_tab import create_liste_contrats_tab
import os
import sys

# Fonction pour obtenir le chemin des ressources empaquetées avec PyInstaller
def resource_path(relative_path):
    """Obtenir le chemin absolu vers la ressource, fonctionne pour PyInstaller."""
    try:
        # PyInstaller crée un répertoire temporaire pour stocker les fichiers
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def main():
    # Créer l'interface graphique avec Tkinter
    root = tk.Tk()
    root.title("Gestionnaire de Réponses Email")

    # Définir la taille initiale de la fenêtre
    root.geometry("1000x800")

    # Utiliser un Notebook pour les onglets
    notebook = ttk.Notebook(root)
    notebook.pack(padx=10, pady=10, fill='both', expand=True)

    # Créer l'onglet "Utilisation"
    tab_main = tk.Frame(notebook)
    create_utilisation_tab(tab_main)
    notebook.add(tab_main, text="Réponses Mail")

    # Créer l'onglet "Création compte rendu"
    tab_creation_compte_rendu = tk.Frame(notebook)
    create_creation_compte_rendu_tab(tab_creation_compte_rendu)
    notebook.add(tab_creation_compte_rendu, text="Création compte rendu")

    # Créer l'onglet "Création ticket ASITEK"
    tab_creation_ticket_asitek = tk.Frame(notebook)
    create_creation_ticket_installation_tab(tab_creation_ticket_asitek)
    notebook.add(tab_creation_ticket_asitek, text="Création ticket ASITEK")

    # Créer l'onglet "Contrat de Services"
    tab_contrat_services = tk.Frame(notebook)
    create_contrat_services_tab(tab_contrat_services)
    notebook.add(tab_contrat_services, text="Création contrat de Services")

    # Créer un nouvel onglet "Liste des contrats"
    tab_liste_contrats = tk.Frame(notebook)
    create_liste_contrats_tab(tab_liste_contrats)
    notebook.add(tab_liste_contrats, text="Liste des contrats")

    # Créer l'onglet "Options"
    tab_options = tk.Frame(notebook)
    create_options_tab(tab_options)
    notebook.add(tab_options, text="Options")

    # Lancer l'application
    root.mainloop()

if __name__ == "__main__":
    main()
