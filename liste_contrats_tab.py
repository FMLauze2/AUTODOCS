import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog  # Ajout de simpledialog

# Chemin vers le fichier CSV
chemin_contrats_csv = os.path.join(os.path.dirname(__file__), 'contrats.csv')

# Fonction pour charger et afficher les contrats
def charger_contrats(tree, filtre=""):
    for item in tree.get_children():
        tree.delete(item)

    if os.path.exists(chemin_contrats_csv):
        with open(chemin_contrats_csv, newline='', encoding='utf-8') as fichier_csv:
            reader = csv.reader(fichier_csv)
            try:
                next(reader)  # Sauter l'en-tête
            except StopIteration:
                return  # Si le fichier est vide

            for row in reader:
                if len(row) == 4:
                    nom_cabinet, date_realisation, fichier_contrat, type_fichier = row
                elif len(row) == 3:
                    nom_cabinet, date_realisation, fichier_contrat = row
                    type_fichier = "PDF" if fichier_contrat.endswith(".pdf") else "Word"

                if filtre.lower() in nom_cabinet.lower():
                    tree.insert("", "end", values=(nom_cabinet, date_realisation, fichier_contrat, type_fichier))

# Fonction pour ajouter un contrat dans le fichier CSV
def ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat):
    fichier_existe = os.path.exists(chemin_contrats_csv)

    with open(chemin_contrats_csv, mode='a', newline='', encoding='utf-8') as fichier_csv:
        writer = csv.writer(fichier_csv)
        if not fichier_existe:
            writer.writerow(['Nom du Cabinet', 'Date de Réalisation', 'Chemin du Fichier', 'Type Fichier'])
        type_fichier = "PDF" if fichier_contrat.endswith(".pdf") else "Word"
        writer.writerow([nom_cabinet, date_realisation, fichier_contrat, type_fichier])

# Fonction pour supprimer un contrat du fichier CSV et de l'interface
def supprimer_contrat(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Erreur", "Veuillez sélectionner un contrat à supprimer.")
        return

    # Obtenir les détails du contrat sélectionné
    contrat_selectionne = tree.item(selected_item, 'values')
    nom_cabinet, date_realisation, fichier_contrat, type_fichier = contrat_selectionne

    confirmation = messagebox.askyesno("Confirmation", f"Voulez-vous vraiment supprimer le contrat pour {nom_cabinet} ?")
    if not confirmation:
        return

    # Lire tous les contrats dans une liste sauf celui à supprimer
    contrats_restants = []
    if os.path.exists(chemin_contrats_csv):
        with open(chemin_contrats_csv, newline='', encoding='utf-8') as fichier_csv:
            reader = csv.reader(fichier_csv)
            en_tete = next(reader)  # Garder l'en-tête
            for row in reader:
                if row[0] != nom_cabinet or row[1] != date_realisation or row[2] != fichier_contrat:
                    contrats_restants.append(row)

    # Réécrire le fichier CSV sans le contrat supprimé
    with open(chemin_contrats_csv, mode='w', newline='', encoding='utf-8') as fichier_csv:
        writer = csv.writer(fichier_csv)
        writer.writerow(en_tete)  # Réécrire l'en-tête
        writer.writerows(contrats_restants)

    # Supprimer l'entrée de la liste affichée
    tree.delete(selected_item)
    messagebox.showinfo("Succès", "Le contrat a été supprimé avec succès.")

# Fonction pour ajouter un seul contrat
def ajouter_un_contrat(tree):
    fichier_contrat = filedialog.askopenfilename(
        title="Sélectionnez un contrat historique (Word ou PDF)",
        filetypes=[("Documents Word ou PDF", "*.docx;*.pdf")]
    )
    
    if fichier_contrat:
        # Extraire le nom du cabinet à partir du nom du fichier (sans l'extension)
        nom_cabinet = os.path.splitext(os.path.basename(fichier_contrat))[0]

        # Demander la date de réalisation
        date_realisation = simpledialog.askstring("Date de Réalisation", f"Entrez la date de réalisation pour {nom_cabinet} (YYYY-MM-DD) :")

        if not date_realisation:
            messagebox.showerror("Erreur", "La date de réalisation est obligatoire.")
            return

        # Ajouter dans le CSV
        ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat)

        # Recharger la liste des contrats
        charger_contrats(tree)
        messagebox.showinfo("Succès", "Le contrat a été ajouté avec succès.")

# Fonction pour créer l'onglet des contrats
def create_liste_contrats_tab(tab_liste_contrats):
    def rechercher_contrat():
        filtre = entry_recherche.get().strip()
        charger_contrats(tree, filtre)

    # Champ de recherche
    label_recherche = tk.Label(tab_liste_contrats, text="Rechercher par nom de cabinet :")
    label_recherche.pack(pady=5)
    entry_recherche = tk.Entry(tab_liste_contrats, width=50)
    entry_recherche.pack(pady=5)

    # Bouton de recherche
    btn_rechercher = tk.Button(tab_liste_contrats, text="Rechercher", command=rechercher_contrat)
    btn_rechercher.pack(pady=5)

    # Tableau (Treeview)
    colonnes = ('Nom du Cabinet', 'Date de Réalisation', 'Chemin du Fichier', 'Type de Fichier')
    tree = ttk.Treeview(tab_liste_contrats, columns=colonnes, show='headings')
    tree.heading('Nom du Cabinet', text='Nom du Cabinet')
    tree.heading('Date de Réalisation', text='Date de Réalisation')
    tree.heading('Chemin du Fichier', text='Chemin du Fichier')
    tree.heading('Type de Fichier', text='Type de Fichier')
    tree.pack(pady=10, padx=10, fill='both', expand=True)

    # Bouton pour ajouter un contrat
    btn_ajouter_contrat = tk.Button(tab_liste_contrats, text="Ajouter Contrat", command=lambda: ajouter_un_contrat(tree))
    btn_ajouter_contrat.pack(pady=10)

    # Bouton pour supprimer un contrat sélectionné
    btn_supprimer_contrat = tk.Button(tab_liste_contrats, text="Supprimer Contrat", command=lambda: supprimer_contrat(tree))
    btn_supprimer_contrat.pack(pady=10)

    # Bouton pour rafraîchir la liste des contrats
    btn_rafraichir = tk.Button(tab_liste_contrats, text="Rafraîchir la liste", command=lambda: charger_contrats(tree))
    btn_rafraichir.pack(pady=10)

    # Charger les contrats existants au démarrage
    charger_contrats(tree)
