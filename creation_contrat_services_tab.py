import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from docx import Document
from datetime import datetime, timedelta
from docx2pdf import convert

# Chemin vers le modèle de contrat
chemin_modele_contrat = os.path.join(os.path.dirname(__file__), 'CONTRAT SERVICE.docx')

# Chemin vers le fichier CSV où sont stockés les contrats
chemin_contrats_csv = os.path.join(os.path.dirname(__file__), 'contrats.csv')

# Fonction pour ajouter un contrat dans le fichier CSV
def ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_contrat, date_envoi="", date_reception=""):
    fichier_existe = os.path.exists(chemin_contrats_csv)

    try:
        with open(chemin_contrats_csv, mode='a', newline='', encoding='utf-8') as fichier_csv:
            writer = csv.writer(fichier_csv)
            if not fichier_existe:
                writer.writerow(['Nom du Cabinet', 'Date de Réalisation', 'Chemin du Fichier', 'Type Fichier', 'Date d\'envoi', 'Date de réception'])
            type_fichier = "PDF" if fichier_contrat.endswith(".pdf") else "Word"
            writer.writerow([nom_cabinet, date_realisation, fichier_contrat, type_fichier, date_envoi, date_reception])
    except Exception as e:
        print(f"Erreur lors de l'ajout du contrat dans le CSV : {e}")

# Fonction pour générer le contrat en PDF
def generer_contrat_pdf():
    nom_cabinet = entry_nom_cabinet.get().strip()
    adresse_cabinet = entry_adresse_cabinet.get().strip()
    cp_cabinet = entry_cp_cabinet.get().strip()
    ville_cabinet = entry_ville_cabinet.get().strip()
    nombre_medecins = entry_nombre_medecins.get().strip()
    prix = entry_prix.get().strip()
    praticiens = listbox_praticiens.get(0, tk.END)
    date_realisation = entry_date_realisation.get().strip()

    if not nom_cabinet or not nombre_medecins or not prix or not date_realisation:
        messagebox.showerror("Erreur", "Tous les champs sont obligatoires.")
        return

    try:
        date_debut = datetime.strptime(date_realisation, "%Y-%m-%d")
        date_fin = date_debut + timedelta(days=365)
        debut_contrat = date_debut.strftime("%Y-%m-%d")
        fin_contrat = date_fin.strftime("%Y-%m-%d")
    except ValueError:
        messagebox.showerror("Erreur", "La date de réalisation doit être au format YYYY-MM-DD.")
        return

    # Charger le modèle Word
    doc = Document(chemin_modele_contrat)

    # Dictionnaire de balises et de valeurs de remplacement
    balises_remplacement = {
        "[NOM_CABINET]": nom_cabinet,
        "[ADRESSE_CABINET]": adresse_cabinet,
        "[CP_CABINET]": cp_cabinet,
        "[VILLE_CABINET]": ville_cabinet,
        "[NOMBRE_MEDECINS]": nombre_medecins,
        "[PRIX]": prix,
        "[DEBUT_CONTRAT]": debut_contrat,
        "[FIN_CONTRAT]": fin_contrat,
        "[DATE_REALISATION]": date_realisation,
        "[PRATICIENS]": ', '.join(praticiens)
    }

    # Remplacer les balises dans chaque paragraphe
    for paragraphe in doc.paragraphs:
        for balise, remplacement in balises_remplacement.items():
            if balise in paragraphe.text:
                paragraphe.text = paragraphe.text.replace(balise, remplacement)

    # Sauvegarder le contrat en Word
    fichier_sortie = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Document Word", "*.docx")])
    if fichier_sortie:
        doc.save(fichier_sortie)

        # Convertir en PDF avec docx2pdf
        fichier_pdf = fichier_sortie.replace('.docx', '.pdf')
        convert(fichier_sortie, fichier_pdf)

        # Ajouter le contrat dans le CSV
        ajouter_contrat_dans_csv(nom_cabinet, date_realisation, fichier_pdf)

        messagebox.showinfo("Succès", "Le contrat de services a été généré en PDF et ajouté à la liste.")
        os.startfile(fichier_pdf)

# Fonction pour créer le formulaire de génération de contrat dans l'onglet
def create_contrat_services_tab(tab_contrat_services):
    # Créer un canvas pour le scroll
    canvas = tk.Canvas(tab_contrat_services)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Ajouter un scrollbar vertical lié au canvas
    scrollbar = tk.Scrollbar(tab_contrat_services, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Créer un frame à l'intérieur du canvas pour y ajouter les widgets
    scrollable_frame = tk.Frame(canvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # Widgets du formulaire
    label_nom_cabinet = tk.Label(scrollable_frame, text="Nom du cabinet :")
    label_nom_cabinet.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    global entry_nom_cabinet
    entry_nom_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_nom_cabinet.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    label_adresse_cabinet = tk.Label(scrollable_frame, text="Adresse du cabinet :")
    label_adresse_cabinet.grid(row=1, column=0, padx=10, pady=5, sticky="e")
    global entry_adresse_cabinet
    entry_adresse_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_adresse_cabinet.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    label_cp_cabinet = tk.Label(scrollable_frame, text="Code postal du cabinet :")
    label_cp_cabinet.grid(row=2, column=0, padx=10, pady=5, sticky="e")
    global entry_cp_cabinet
    entry_cp_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_cp_cabinet.grid(row=2, column=1, padx=10, pady=5, sticky="w")

    label_ville_cabinet = tk.Label(scrollable_frame, text="Ville du cabinet :")
    label_ville_cabinet.grid(row=3, column=0, padx=10, pady=5, sticky="e")
    global entry_ville_cabinet
    entry_ville_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_ville_cabinet.grid(row=3, column=1, padx=10, pady=5, sticky="w")

    label_nombre_medecins = tk.Label(scrollable_frame, text="Nombre de médecins :")
    label_nombre_medecins.grid(row=4, column=0, padx=10, pady=5, sticky="e")
    global entry_nombre_medecins
    entry_nombre_medecins = tk.Entry(scrollable_frame, width=50)
    entry_nombre_medecins.grid(row=4, column=1, padx=10, pady=5, sticky="w")

    label_prix = tk.Label(scrollable_frame, text="Prix du contrat :")
    label_prix.grid(row=5, column=0, padx=10, pady=5, sticky="e")
    global entry_prix
    entry_prix = tk.Entry(scrollable_frame, width=50)
    entry_prix.grid(row=5, column=1, padx=10, pady=5, sticky="w")

    label_liste_praticiens = tk.Label(scrollable_frame, text="Liste des praticiens (Nom, Prénom) :")
    label_liste_praticiens.grid(row=6, column=0, padx=10, pady=5, sticky="ne")
    global listbox_praticiens
    listbox_praticiens = tk.Listbox(scrollable_frame, height=5, width=50)
    listbox_praticiens.grid(row=6, column=1, padx=10, pady=5, sticky="w")

    btn_ajouter_praticien = tk.Button(scrollable_frame, text="Ajouter Praticien", command=lambda: ajouter_praticien(listbox_praticiens))
    btn_ajouter_praticien.grid(row=7, column=1, padx=10, pady=5, sticky="w")

    label_date_realisation = tk.Label(scrollable_frame, text="Date de réalisation :")
    label_date_realisation.grid(row=8, column=0, padx=10, pady=5, sticky="e")
    global entry_date_realisation
    entry_date_realisation = tk.Entry(scrollable_frame, width=50)
    entry_date_realisation.insert(0, datetime.now().strftime("%Y-%m-%d"))
    entry_date_realisation.grid(row=8, column=1, padx=10, pady=5, sticky="w")

    # Bouton pour générer le contrat en PDF
    btn_generer_pdf = tk.Button(scrollable_frame, text="Générer le PDF", command=generer_contrat_pdf)
    btn_generer_pdf.grid(row=9, column=1, padx=10, pady=20)

    # Bouton pour rafraîchir le formulaire
    btn_rafraichir_formulaire = tk.Button(scrollable_frame, text="Rafraîchir le formulaire", command=rafraichir_formulaire)
    btn_rafraichir_formulaire.grid(row=10, column=1, padx=10, pady=10)

# Fonction pour rafraîchir le formulaire après la génération du contrat
def rafraichir_formulaire():
    entry_nom_cabinet.delete(0, tk.END)
    entry_adresse_cabinet.delete(0, tk.END)
    entry_cp_cabinet.delete(0, tk.END)
    entry_ville_cabinet.delete(0, tk.END)
    entry_nombre_medecins.delete(0, tk.END)
    entry_prix.delete(0, tk.END)
    listbox_praticiens.delete(0, tk.END)
    entry_date_realisation.delete(0, tk.END)
    entry_date_realisation.insert(0, datetime.now().strftime("%Y-%m-%d"))

# Fonction pour ajouter un praticien à la liste
def ajouter_praticien(listbox_praticiens):
    praticien = simpledialog.askstring("Praticien", "Entrez le nom et prénom du praticien :")
    if praticien:
        listbox_praticiens.insert(tk.END, praticien)
