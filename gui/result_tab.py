# gui/result_tab.py

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from matplotlib.ticker import MultipleLocator
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt


class ResultTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_file = None
        self.save_folder = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        self.setLayout(layout)

        self.info_label = QLabel("Analyse générale des résultats")
        layout.addWidget(self.info_label)

        self.btn_select_save = QPushButton("Sélectionner dossier de sauvegarde")
        self.btn_select_save.clicked.connect(self.select_save_folder)
        layout.addWidget(self.btn_select_save)

        self.save_folder_label = QLabel("Aucun dossier sélectionné")
        layout.addWidget(self.save_folder_label)

        self.chart_type_combo = QComboBox()
        # Ici nous proposons deux types de graphique
        self.chart_type_combo.addItem("Graphique linéaire par station")
        self.chart_type_combo.addItem("Graphique en barres par station")
        layout.addWidget(self.chart_type_combo)

        self.btn_generate = QPushButton("Générer graphiques")
        self.btn_generate.clicked.connect(self.generate_charts)
        layout.addWidget(self.btn_generate)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

    def set_excel_file(self, excel_file):
        self.excel_file = excel_file

    def select_save_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier de sauvegarde")
        if folder:
            self.save_folder = folder
            self.save_folder_label.setText(folder)

    def generate_charts(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Attention", "Aucun fichier Excel chargé.")
            return
        if not self.save_folder:
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner un dossier de sauvegarde.")
            return

        chart_type = self.chart_type_combo.currentText()

        # Chemin des logos (dossier "logo" à la racine)
        logo_folder = os.path.join(os.getcwd(), "logo")
        logo_altiplage_path = os.path.join(logo_folder, "logo_Altiplage.png")
        lamanche_logo_path = os.path.join(logo_folder, "lamanche.jpg")
        cnam_logo_path = os.path.join(logo_folder, "cnam.png")

        # Lire l'Excel pour obtenir les sites (feuilles sauf "Résumé")
        xls = pd.ExcelFile(self.excel_file)
        sheet_names = [s for s in xls.sheet_names if s != "Résumé"]

        saved_files = []

        for site in sheet_names:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=site)
                # Conversion des dates avec format explicite
                if "Date / Heure" in df.columns:
                    df["Date / Heure"] = pd.to_datetime(
                        df["Date / Heure"], format="%d/%m/%Y %Hh%Mm%S", errors="coerce", dayfirst=True)
                if "Résultat" in df.columns:
                    df["Résultat"] = pd.to_numeric(df["Résultat"], errors="coerce")

                # Création d'une nouvelle figure pour chaque site
                fig, ax = plt.subplots(figsize=(10, 6))

                if chart_type == "Graphique linéaire par station":
                    if "Date / Heure" in df.columns and "Résultat" in df.columns:
                        df = df.sort_values("Date / Heure")
                        ax.plot(df["Date / Heure"], df["Résultat"], marker="o", color='blue', linestyle='-')
                        ax.set_xlabel("Date")
                        ax.set_ylabel("Résultat (cm)")
                        ax.tick_params(axis='x', rotation=45)
                    ax.set_title(f"Evolution sédimentaire: - Site : {site}")

                elif chart_type == "Graphique en barres par station":
                    if "Date / Heure" in df.columns and "Résultat" in df.columns:
                        df = df.sort_values("Date / Heure")
                        dates_str = df["Date / Heure"].dt.strftime("%d/%m/%Y")
                        ax.bar(dates_str, df["Résultat"], color='blue')
                        ax.set_xlabel("Date")
                        ax.set_ylabel("Résultat (cm)")
                        ax.tick_params(axis='x', rotation=45)
                    ax.set_title(f"Evolution sédimentaire: - Site : {site}")
                else:
                    QMessageBox.warning(self, "Attention", "Type de graphique non reconnu.")
                    continue

                # Définition des limites de l'axe y : min - 10 et max + 10, avec un pas de 2 cm
                if not df["Résultat"].empty:
                    min_val = df["Résultat"].min()
                    max_val = df["Résultat"].max()
                else:
                    min_val, max_val = 0, 0
                ax.set_ylim(min_val - 10, max_val + 10)
                ax.yaxis.set_major_locator(MultipleLocator(2))
                ax.yaxis.grid(True, linestyle='--', alpha=0.7)

                fig.subplots_adjust(top=0.80)
                fig.text(0.5, 0.99, f"Evolution sédimentaire: - Site : {site}",
                         ha='center', va='center', fontsize=16, fontweight='bold', color='blue')

                # Ajout des logos
                ax_logo_left = fig.add_axes([0.02, 0.98, 0.10, 0.20], zorder=10)
                try:
                    logo_altiplage = mpimg.imread(logo_altiplage_path)
                    ax_logo_left.imshow(logo_altiplage)
                except Exception as e:
                    print(f"Erreur lors du chargement du logo Altiplage : {e}")
                ax_logo_left.axis('off')

                ax_logo_top_right1 = fig.add_axes([0.78, 0.98, 0.10, 0.20], zorder=10)
                try:
                    logo_lamanche = mpimg.imread(lamanche_logo_path)
                    ax_logo_top_right1.imshow(logo_lamanche)
                except Exception as e:
                    print(f"Erreur lors du chargement du logo lamanche : {e}")
                ax_logo_top_right1.axis('off')

                ax_logo_top_right2 = fig.add_axes([0.90, 0.98, 0.10, 0.20], zorder=10)
                try:
                    logo_cnam = mpimg.imread(cnam_logo_path)
                    ax_logo_top_right2.imshow(logo_cnam)
                except Exception as e:
                    print(f"Erreur lors du chargement du logo cnam : {e}")
                ax_logo_top_right2.axis('off')

                if chart_type == "Graphique linéaire par station":
                    output_file = f"{site}_line.png"
                else:
                    output_file = f"{site}_bar.png"
                output_path = os.path.join(self.save_folder, output_file)
                plt.savefig(output_path, bbox_inches='tight')
                plt.close(fig)
                saved_files.append(output_path)
            except Exception as e:
                print(f"Erreur lors de la génération du graphique pour le site {site}: {e}")
                continue

        if saved_files:
            QMessageBox.information(self, "Succès",
                                    f"Les graphiques ont été générés et sauvegardés dans :\n{self.save_folder}")
        else:
            QMessageBox.warning(self, "Erreur", "Aucun graphique n'a pu être généré.")


# Pour utiliser mpimg
import matplotlib.image as mpimg

if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    from gui.app import MainWindow

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
