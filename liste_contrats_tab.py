import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

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
                # Assurer que chaque ligne ait bien 6 colonnes
                while len(row) < 6:
                    row.append("")  # Ajouter des colonnes vides pour les lignes incomplètes

                nom_cabinet, date_realisation, fichier_contrat, type_fichier, date_envoi, date_reception = row

                if filtre.lower() in nom_cabinet.lower():
                    tree.insert("", "end", values=(nom_cabinet, date_realisation, fichier_contrat, type_fichier, date_envoi, date_reception))

# Fonction pour ajouter ou mettre à jour un contrat dans le fichier CSV
def ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat, date_envoi="", date_reception=""):
    fichier_existe = os.path.exists(chemin_contrats_csv)

    # Lire l'existant
    contrats = []
    if fichier_existe:
        with open(chemin_contrats_csv, newline='', encoding='utf-8') as fichier_csv:
            reader = csv.reader(fichier_csv)
            en_tete = next(reader, None)  # Récupérer l'en-tête si elle existe
            contrats = list(reader)

    # Vérifier si le contrat existe déjà
    contrat_existant = False
    for i, row in enumerate(contrats):
        if row[0] == nom_cabinet and row[1] == date_realisation:
            while len(row) < 6:
                row.append("")
            row[4] = date_envoi if date_envoi else row[4]
            row[5] = date_reception if date_reception else row[5]
            contrat_existant = True
            break

    if not contrat_existant:
        type_fichier = "PDF" if fichier_contrat.endswith(".pdf") else "Word"
        contrats.append([nom_cabinet, date_realisation, fichier_contrat, type_fichier, date_envoi, date_reception])

    # Écriture dans le fichier CSV
    try:
        with open(chemin_contrats_csv, mode='w', newline='', encoding='utf-8') as fichier_csv:
            writer = csv.writer(fichier_csv)
            writer.writerow(['Nom du Cabinet', 'Date de Réalisation', 'Chemin du Fichier', 'Type Fichier', 'Date d\'envoi', 'Date de réception'])
            writer.writerows(contrats)
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la mise à jour du CSV : {e}")

# Fonction pour ajouter un contrat
def ajouter_un_contrat(tree):
    fichier_contrat = filedialog.askopenfilename(
        title="Sélectionnez un contrat historique (Word ou PDF)",
        filetypes=[("Documents Word ou PDF", "*.docx;*.pdf")]
    )
    
    if fichier_contrat:
        nom_cabinet = os.path.splitext(os.path.basename(fichier_contrat))[0]
        date_realisation = simpledialog.askstring("Date de Réalisation", f"Entrez la date de réalisation pour {nom_cabinet} (YYYY-MM-DD) :")

        if not date_realisation:
            messagebox.showerror("Erreur", "La date de réalisation est obligatoire.")
            return

        ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat)
        charger_contrats(tree)
        messagebox.showinfo("Succès", "Le contrat a été ajouté avec succès.")

# Fonction pour ajouter une date d'envoi ou de réception à un contrat sélectionné
def ajouter_date(tree, type_date):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Erreur", "Veuillez sélectionner un contrat.")
        return

    contrat_selectionne = tree.item(selected_item, 'values')
    nom_cabinet, date_realisation, fichier_contrat, type_fichier, date_envoi, date_reception = contrat_selectionne

    nouvelle_date = simpledialog.askstring(f"Date de {type_date}", f"Entrez la date de {type_date} pour {nom_cabinet} (YYYY-MM-DD) :")

    if not nouvelle_date:
        messagebox.showerror("Erreur", f"La date de {type_date} est obligatoire.")
        return

    if type_date == "envoi":
        ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat, date_envoi=nouvelle_date)
    elif type_date == "réception":
        ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat, date_reception=nouvelle_date)

    charger_contrats(tree)
    messagebox.showinfo("Succès", f"La date de {type_date} a été mise à jour pour {nom_cabinet}.")

# Fonction pour supprimer un contrat du fichier CSV et de l'interface
def supprimer_contrat(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Erreur", "Veuillez sélectionner un contrat à supprimer.")
        return

    # Obtenir les détails du contrat sélectionné
    contrat_selectionne = tree.item(selected_item, 'values')
    nom_cabinet, date_realisation, fichier_contrat, type_fichier, date_envoi, date_reception = contrat_selectionne

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
                # Comparer les détails pour s'assurer que c'est bien le contrat à supprimer
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

# Fonction pour créer l'onglet des contrats
def create_liste_contrats_tab(tab_liste_contrats):
    def rechercher_contrat():
        filtre = entry_recherche.get().strip()
        charger_contrats(tree, filtre)

    label_recherche = tk.Label(tab_liste_contrats, text="Rechercher par nom de cabinet :")
    label_recherche.pack(pady=5)
    entry_recherche = tk.Entry(tab_liste_contrats, width=50)
    entry_recherche.pack(pady=5)

    btn_rechercher = tk.Button(tab_liste_contrats, text="Rechercher", command=rechercher_contrat)
    btn_rechercher.pack(pady=5)

    colonnes = ('Nom du Cabinet', 'Date de Réalisation', 'Chemin du Fichier', 'Type de Fichier', 'Date d\'envoi', 'Date de réception')
    tree = ttk.Treeview(tab_liste_contrats, columns=colonnes, show='headings')
    tree.heading('Nom du Cabinet', text='Nom du Cabinet')
    tree.heading('Date de Réalisation', text='Date de Réalisation')
    tree.heading('Chemin du Fichier', text='Chemin du Fichier')
    tree.heading('Type de Fichier', text='Type de Fichier')
    tree.heading('Date d\'envoi', text='Date d\'envoi')
    tree.heading('Date de réception', text='Date de réception')
    tree.pack(pady=10, padx=10, fill='both', expand=True)

    btn_ajouter_contrat = tk.Button(tab_liste_contrats, text="Ajouter Contrat", command=lambda: ajouter_un_contrat(tree))
    btn_ajouter_contrat.pack(pady=10)

    btn_supprimer_contrat = tk.Button(tab_liste_contrats, text="Supprimer Contrat", command=lambda: supprimer_contrat(tree))
    btn_supprimer_contrat.pack(pady=10)

    btn_ajouter_date_envoi = tk.Button(tab_liste_contrats, text="Ajouter Date d'Envoi", command=lambda: ajouter_date(tree, "envoi"))
    btn_ajouter_date_envoi.pack(pady=10)

    btn_ajouter_date_reception = tk.Button(tab_liste_contrats, text="Ajouter Date de Réception", command=lambda: ajouter_date(tree, "réception"))
    btn_ajouter_date_reception.pack(pady=10)

    btn_rafraichir = tk.Button(tab_liste_contrats, text="Rafraîchir la liste", command=lambda: charger_contrats(tree))
    btn_rafraichir.pack(pady=10)

    charger_contrats(tree)
