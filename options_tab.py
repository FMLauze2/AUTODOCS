import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import json
import os

# Chemin absolu pour les fichiers JSON
chemin_reponses = os.path.join(os.path.dirname(__file__), 'reponses.json')

def sauvegarder_reponses(reponses):
    with open(chemin_reponses, 'w', encoding='utf-8') as f:
        json.dump(reponses, f, ensure_ascii=False, indent=4)

def charger_reponses():
    with open(chemin_reponses, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_options_tab(tab_options):
    reponses = charger_reponses()
    modele_actuel = tk.StringVar(value="")

    # Fonction pour mettre à jour la liste des modèles de réponse
    def mettre_a_jour_liste():
        combo_models['values'] = list(reponses.keys())
        if modele_actuel.get() in reponses:
            combo_models.set(modele_actuel.get())
        else:
            combo_models.current(0)  # Sélectionner le premier modèle

    # Fonction pour charger un modèle dans les champs d'édition
    def charger_modele(event=None):
        titre = combo_models.get()
        contenu = reponses.get(titre, "")
        entry_titre.delete(0, tk.END)
        entry_titre.insert(0, titre)
        text_contenu.delete(1.0, tk.END)
        text_contenu.insert(tk.END, contenu)
        modele_actuel.set(titre)

    # Fonction pour ajouter un nouveau modèle
    def ajouter_modele():
        titre = entry_titre.get().strip()
        contenu = text_contenu.get(1.0, tk.END).strip()

        if titre and contenu:
            if titre in reponses:
                messagebox.showerror("Erreur", "Un modèle avec ce titre existe déjà. Choisissez un autre titre.")
            else:
                reponses[titre] = contenu
                sauvegarder_reponses(reponses)
                modele_actuel.set(titre)
                mettre_a_jour_liste()
                messagebox.showinfo("Succès", "Le nouveau modèle a été ajouté avec succès.")
        else:
            messagebox.showerror("Erreur", "Le titre et le contenu ne peuvent pas être vides.")

    # Fonction pour sauvegarder un modèle existant
    def sauvegarder_modele():
        titre = entry_titre.get().strip()
        contenu = text_contenu.get(1.0, tk.END).strip()

        if titre and contenu:
            if modele_actuel.get() and modele_actuel.get() != titre:
                # Si le titre a changé, supprimer l'ancien modèle
                del reponses[modele_actuel.get()]
            reponses[titre] = contenu
            sauvegarder_reponses(reponses)
            modele_actuel.set(titre)
            mettre_a_jour_liste()
            messagebox.showinfo("Succès", "Le modèle a été sauvegardé avec succès.")
        else:
            messagebox.showerror("Erreur", "Le titre et le contenu ne peuvent pas être vides.")

    # Fonction pour supprimer un modèle
    def supprimer_modele():
        titre = combo_models.get()

        if titre in reponses:
            del reponses[titre]
            sauvegarder_reponses(reponses)
            entry_titre.delete(0, tk.END)
            text_contenu.delete(1.0, tk.END)
            modele_actuel.set("")
            mettre_a_jour_liste()
            messagebox.showinfo("Succès", "Le modèle a été supprimé avec succès.")
        else:
            messagebox.showerror("Erreur", "Le modèle sélectionné n'existe pas.")

    # Interface pour gérer les modèles de réponse
    label_models = tk.Label(tab_options, text="Modèles de réponse :")
    label_models.grid(row=0, column=0, padx=5, pady=5, sticky="w")

    combo_models = ttk.Combobox(tab_options, state="readonly")
    combo_models.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    combo_models.bind("<<ComboboxSelected>>", charger_modele)

    label_titre = tk.Label(tab_options, text="Titre du modèle :")
    label_titre.grid(row=1, column=0, padx=5, pady=5, sticky="w")

    entry_titre = tk.Entry(tab_options, width=50)
    entry_titre.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    label_contenu = tk.Label(tab_options, text="Contenu du modèle :")
    label_contenu.grid(row=2, column=0, padx=5, pady=5, sticky="nw")

    text_contenu = tk.Text(tab_options, wrap='word', height=10, width=50)
    text_contenu.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    btn_ajouter = tk.Button(tab_options, text="Ajouter", command=ajouter_modele)
    btn_ajouter.grid(row=3, column=0, padx=5, pady=2, sticky="w")

    btn_sauvegarder = tk.Button(tab_options, text="Sauvegarder", command=sauvegarder_modele)
    btn_sauvegarder.grid(row=4, column=0, padx=5, pady=2, sticky="w")

    btn_supprimer = tk.Button(tab_options, text="Supprimer", command=supprimer_modele)
    btn_supprimer.grid(row=5, column=0, padx=5, pady=2, sticky="w")

    # Initialiser la liste des modèles
    mettre_a_jour_liste()
    charger_modele()  # Charger le premier modèle par défaut
