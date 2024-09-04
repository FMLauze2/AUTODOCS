import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from docx import Document
from datetime import datetime, timedelta
import os
from docx2pdf import convert  # Pour générer le PDF directement à partir du fichier Word

# Chemin vers le modèle de contrat
chemin_modele_contrat = os.path.join(os.path.dirname(__file__), 'CONTRAT SERVICE.docx')

def create_contrat_services_tab(tab_contrat_services):
    # Créer un canvas pour le scroll
    canvas = tk.Canvas(tab_contrat_services)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Ajouter un scrollbar vertical lié au canvas
    scrollbar = tk.Scrollbar(tab_contrat_services, orient=tk.VERTICAL, command=canvas.yview)
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

    # Widgets du formulaire
    label_nom_cabinet = tk.Label(scrollable_frame, text="Nom du cabinet :")
    label_nom_cabinet.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    entry_nom_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_nom_cabinet.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    label_adresse_cabinet = tk.Label(scrollable_frame, text="Adresse du cabinet :")
    label_adresse_cabinet.grid(row=1, column=0, padx=10, pady=5, sticky="e")
    entry_adresse_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_adresse_cabinet.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    label_cp_cabinet = tk.Label(scrollable_frame, text="Code postal du cabinet :")
    label_cp_cabinet.grid(row=2, column=0, padx=10, pady=5, sticky="e")
    entry_cp_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_cp_cabinet.grid(row=2, column=1, padx=10, pady=5, sticky="w")

    label_ville_cabinet = tk.Label(scrollable_frame, text="Ville du cabinet :")
    label_ville_cabinet.grid(row=3, column=0, padx=10, pady=5, sticky="e")
    entry_ville_cabinet = tk.Entry(scrollable_frame, width=50)
    entry_ville_cabinet.grid(row=3, column=1, padx=10, pady=5, sticky="w")

    label_nombre_medecins = tk.Label(scrollable_frame, text="Nombre de médecins :")
    label_nombre_medecins.grid(row=4, column=0, padx=10, pady=5, sticky="e")
    entry_nombre_medecins = tk.Entry(scrollable_frame, width=50)
    entry_nombre_medecins.grid(row=4, column=1, padx=10, pady=5, sticky="w")

    label_prix = tk.Label(scrollable_frame, text="Prix du contrat :")
    label_prix.grid(row=5, column=0, padx=10, pady=5, sticky="e")
    entry_prix = tk.Entry(scrollable_frame, width=50)
    entry_prix.grid(row=5, column=1, padx=10, pady=5, sticky="w")

    label_liste_praticiens = tk.Label(scrollable_frame, text="Liste des praticiens (Nom, Prénom) :")
    label_liste_praticiens.grid(row=6, column=0, padx=10, pady=5, sticky="ne")
    listbox_praticiens = tk.Listbox(scrollable_frame, height=5, width=50)
    listbox_praticiens.grid(row=6, column=1, padx=10, pady=5, sticky="w")

    btn_ajouter_praticien = tk.Button(scrollable_frame, text="Ajouter Praticien", command=lambda: ajouter_praticien(listbox_praticiens))
    btn_ajouter_praticien.grid(row=7, column=1, padx=10, pady=5, sticky="w")

    label_date_realisation = tk.Label(scrollable_frame, text="Date de réalisation (début du contrat) :")
    label_date_realisation.grid(row=8, column=0, padx=10, pady=5, sticky="e")
    entry_date_realisation = tk.Entry(scrollable_frame, width=50)
    entry_date_realisation.insert(0, datetime.now().strftime("%Y-%m-%d"))
    entry_date_realisation.grid(row=8, column=1, padx=10, pady=5, sticky="w")

    # Fonction pour ajouter un praticien
    def ajouter_praticien(listbox_praticiens):
        praticien = simpledialog.askstring("Praticien", "Entrez le nom et prénom du praticien :")
        if praticien:
            listbox_praticiens.insert(tk.END, praticien)

    # Fonction pour remplacer les balises tout en conservant la mise en forme
    def remplacer_balises_dans_run(paragraph, balises_remplacement):
        """Remplacer les balises dans les runs d'un paragraphe sans modifier la mise en forme."""
        for run in paragraph.runs:
            for balise, remplacement in balises_remplacement.items():
                if balise in run.text:
                    run.text = run.text.replace(balise, remplacement)

    # Fonction pour calculer la date de fin du contrat (1 an après la date de réalisation)
    def calculer_fin_contrat(date_realisation_str):
        date_realisation = datetime.strptime(date_realisation_str, "%Y-%m-%d")
        date_fin = date_realisation + timedelta(days=365)  # Ajouter 365 jours pour une année
        return date_fin.strftime("%Y-%m-%d")

    # Fonction pour générer le contrat Word
    def generer_contrat():
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

        # Calculer la date de fin du contrat
        date_fin = calculer_fin_contrat(date_realisation)

        # Vérifier si le fichier modèle existe
        if not os.path.exists(chemin_modele_contrat):
            messagebox.showerror("Erreur", f"Le fichier {chemin_modele_contrat} est introuvable.")
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
            "[DEBUT_CONTRAT]": date_realisation,
            "[FIN_CONTRAT]": date_fin,
            "[DATE_REALISATION]": date_realisation,
            "[PRATICIENS]": ', '.join(praticiens)
        }

        # Remplacer les balises dans chaque paragraphe tout en conservant la mise en forme
        for paragraphe in doc.paragraphs:
            remplacer_balises_dans_run(paragraphe, balises_remplacement)

        # Sauvegarder le fichier Word
        fichier_sortie = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Document Word", "*.docx")])
        if fichier_sortie:
            doc.save(fichier_sortie)
            messagebox.showinfo("Succès", "Le contrat de services a été généré avec succès.")
            os.startfile(fichier_sortie)

    # Fonction pour générer directement un fichier PDF
    def generer_pdf():
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

        # Calculer la date de fin du contrat
        date_fin = calculer_fin_contrat(date_realisation)

        # Utiliser un modèle Word pour la conversion PDF
        doc = Document(chemin_modele_contrat)

        # Dictionnaire de balises et de valeurs de remplacement
        balises_remplacement = {
            "[NOM_CABINET]": nom_cabinet,
            "[ADRESSE_CABINET]": adresse_cabinet,
            "[CP_CABINET]": cp_cabinet,
            "[VILLE_CABINET]": ville_cabinet,
            "[NOMBRE_MEDECINS]": nombre_medecins,
            "[PRIX]": prix,
            "[DEBUT_CONTRAT]": date_realisation,
            "[FIN_CONTRAT]": date_fin,
            "[DATE_REALISATION]": date_realisation,
            "[PRATICIENS]": ', '.join(praticiens)
        }

        # Remplacer les balises dans chaque paragraphe tout en conservant la mise en forme
        for paragraphe in doc.paragraphs:
            remplacer_balises_dans_run(paragraphe, balises_remplacement)

        # Sauvegarder le document Word avant conversion en PDF
        fichier_sortie = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Document Word", "*.docx")])
        if fichier_sortie:
            doc.save(fichier_sortie)

            # Convertir le fichier Word en PDF
            fichier_pdf = fichier_sortie.replace('.docx', '.pdf')
            convert(fichier_sortie, fichier_pdf)
            messagebox.showinfo("Succès", "Le contrat de services a été généré en PDF.")
            os.startfile(fichier_pdf)

    # Bouton pour générer le contrat Word
    btn_generer_contrat = tk.Button(scrollable_frame, text="Générer et Ouvrir le Contrat Word", command=generer_contrat)
    btn_generer_contrat.grid(row=11, column=1, padx=10, pady=10)

    # Bouton pour générer le contrat PDF
    btn_generer_pdf = tk.Button(scrollable_frame, text="Générer et Ouvrir le Contrat PDF", command=generer_pdf)
    btn_generer_pdf.grid(row=12, column=1, padx=10, pady=10)
