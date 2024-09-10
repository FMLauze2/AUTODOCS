import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Chemin vers le fichier CSV où sont stockées les fiches prospects
chemin_prospect_csv = os.path.join(os.path.dirname(__file__), 'fiches_prospects.csv')

# Fonction pour ajouter une fiche prospect dans le fichier CSV
def ajouter_fiche_prospect(nom_contact, ville, logiciel_metier, nb_praticiens, nb_postes, besoin_lecteur, email, telephone, date_ouverture):
    fichier_existe = os.path.exists(chemin_prospect_csv)

    with open(chemin_prospect_csv, mode='a', newline='', encoding='utf-8') as fichier_csv:
        writer = csv.writer(fichier_csv)
        if not fichier_existe:
            writer.writerow(['Nom Contact', 'Ville', 'Logiciel Métier', 'Nombre Praticiens', 'Nombre Postes', 'Besoin Lecteur', 'Email', 'Téléphone', 'Date d\'Ouverture'])
        writer.writerow([nom_contact, ville, logiciel_metier, nb_praticiens, nb_postes, besoin_lecteur, email, telephone, date_ouverture])

# Fonction pour mettre à jour la visionneuse de texte
def mettre_a_jour_visionneuse():
    nom_contact = entry_nom_contact.get()
    ville = entry_ville.get()
    logiciel_metier = entry_logiciel_metier.get()
    nb_praticiens = entry_nb_praticiens.get()
    nb_postes = entry_nb_postes.get()
    besoin_lecteur = entry_besoin_lecteur.get()
    email = entry_email.get()
    telephone = entry_telephone.get()
    date_ouverture = entry_date_ouverture.get()

    texte_visionneuse = f"""
    Nom du Contact : {nom_contact}
    Ville : {ville}
    Logiciel Métier : {logiciel_metier}
    Nombre de Praticiens : {nb_praticiens}
    Nombre de Postes : {nb_postes}
    Besoin de Lecteur : {besoin_lecteur}
    Email : {email}
    Téléphone : {telephone}
    Date d'Ouverture Envisagée : {date_ouverture}
    """

    # Mettre à jour la zone de texte avec le contenu du formulaire
    texte_visionneuse_widget.config(state=tk.NORMAL)
    texte_visionneuse_widget.delete(1.0, tk.END)
    texte_visionneuse_widget.insert(tk.END, texte_visionneuse)
    texte_visionneuse_widget.config(state=tk.DISABLED)

# Fonction pour créer l'onglet "Création fichier prospect"
def create_fichier_prospect_tab(tab_prospect):
    global entry_nom_contact, entry_ville, entry_logiciel_metier, entry_nb_praticiens, entry_nb_postes, entry_besoin_lecteur, entry_email, entry_telephone, entry_date_ouverture, texte_visionneuse_widget
    
    # Frame pour contenir la visionneuse et les champs de saisie côte à côte
    frame_main = tk.Frame(tab_prospect)
    frame_main.pack(fill=tk.BOTH, expand=True)

    # Frame pour le formulaire de saisie
    frame_form = tk.Frame(frame_main)
    frame_form.pack(side=tk.LEFT, padx=10, pady=10, anchor='n')

    # Widgets du formulaire
    tk.Label(frame_form, text="Nom Contact :").grid(row=0, column=0, sticky="e")
    entry_nom_contact = tk.Entry(frame_form)
    entry_nom_contact.grid(row=0, column=1)
    entry_nom_contact.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Ville :").grid(row=1, column=0, sticky="e")
    entry_ville = tk.Entry(frame_form)
    entry_ville.grid(row=1, column=1)
    entry_ville.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Logiciel Métier :").grid(row=2, column=0, sticky="e")
    entry_logiciel_metier = tk.Entry(frame_form)
    entry_logiciel_metier.grid(row=2, column=1)
    entry_logiciel_metier.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Nombre Praticiens :").grid(row=3, column=0, sticky="e")
    entry_nb_praticiens = tk.Entry(frame_form)
    entry_nb_praticiens.grid(row=3, column=1)
    entry_nb_praticiens.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Nombre Postes :").grid(row=4, column=0, sticky="e")
    entry_nb_postes = tk.Entry(frame_form)
    entry_nb_postes.grid(row=4, column=1)
    entry_nb_postes.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Besoin de Lecteur :").grid(row=5, column=0, sticky="e")
    entry_besoin_lecteur = tk.Entry(frame_form)
    entry_besoin_lecteur.grid(row=5, column=1)
    entry_besoin_lecteur.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Email :").grid(row=6, column=0, sticky="e")
    entry_email = tk.Entry(frame_form)
    entry_email.grid(row=6, column=1)
    entry_email.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Téléphone :").grid(row=7, column=0, sticky="e")
    entry_telephone = tk.Entry(frame_form)
    entry_telephone.grid(row=7, column=1)
    entry_telephone.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    tk.Label(frame_form, text="Date d'Ouverture Envisagée :").grid(row=8, column=0, sticky="e")
    entry_date_ouverture = tk.Entry(frame_form)
    entry_date_ouverture.grid(row=8, column=1)
    entry_date_ouverture.bind("<KeyRelease>", lambda e: mettre_a_jour_visionneuse())

    # Frame pour la visionneuse de texte
    frame_visionneuse = tk.Frame(frame_main)
    frame_visionneuse.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Zone de texte pour la visionneuse
    texte_visionneuse_widget = tk.Text(frame_visionneuse, height=30, width=50, state=tk.DISABLED)
    texte_visionneuse_widget.pack(fill=tk.BOTH, expand=True)

    # Bouton pour sauvegarder les informations du prospect
    btn_sauvegarder = tk.Button(frame_form, text="Sauvegarder Fiche Prospect", command=lambda: ajouter_fiche_prospect(
        entry_nom_contact.get(), entry_ville.get(), entry_logiciel_metier.get(), entry_nb_praticiens.get(),
        entry_nb_postes.get(), entry_besoin_lecteur.get(), entry_email.get(), entry_telephone.get(), entry_date_ouverture.get()
    ))
    btn_sauvegarder.grid(row=9, columnspan=2, pady=10)

