# gui/result_tab.py

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox,
    QFileDialog, QMessageBox
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

        # Choix du dossier de sauvegarde
        self.btn_select_save = QPushButton("Sélectionner dossier de sauvegarde")
        self.btn_select_save.clicked.connect(self.select_save_folder)
        layout.addWidget(self.btn_select_save)

        self.save_folder_label = QLabel("Aucun dossier sélectionné")
        layout.addWidget(self.save_folder_label)

        # Type de graphique
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItem("Graphique linéaire par station")
        self.chart_type_combo.addItem("Graphique en barres par station")
        layout.addWidget(self.chart_type_combo)

        # Bouton de génération
        self.btn_generate = QPushButton("Générer graphiques")
        self.btn_generate.clicked.connect(self.generate_charts)
        layout.addWidget(self.btn_generate)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

    def set_excel_file(self, excel_file: str):
        self.excel_file = excel_file

    def select_save_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Sélectionner le dossier de sauvegarde"
        )
        if folder:
            self.save_folder = folder
            self.save_folder_label.setText(folder)

    def generate_charts(self):
        # Vérifications initiales
        if not self.excel_file:
            QMessageBox.warning(self, "Attention", "Aucun fichier Excel chargé.")
            return
        if not self.save_folder:
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner un dossier de sauvegarde.")
            return

        chart_type = self.chart_type_combo.currentText()

        # 1) Lecture de la feuille "Résumé" pour récupérer Z_CC49
        try:
            summary = pd.read_excel(
                self.excel_file,
                sheet_name="Résumé",
                usecols=["Station", "Z_CC49"]
            )
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire la feuille 'Résumé' :\n{e}")
            return

        summary["Station"] = summary["Station"].astype(str).str.strip().str.upper()

        # 2) Préparation des chemins de logos
        logo_folder = os.path.join(os.getcwd(), "logo")
        logo_paths = {
            "altiplage": os.path.join(logo_folder, "Altipl5.png"),
            "lamanche":  os.path.join(logo_folder, "lamanche.jpg"),
            "cnam":      os.path.join(logo_folder, "cnam.png")
        }

        # 3) Parcours des feuilles (hors "Résumé")
        xls = pd.ExcelFile(self.excel_file)
        sheets = [s for s in xls.sheet_names if s != "Résumé"]
        saved_files = []

        for site in sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=site)

                # Conversion de la date
                if "Date / Heure" in df.columns:
                    df["Date / Heure"] = pd.to_datetime(
                        df["Date / Heure"],
                        format="%d/%m/%Y %Hh%Mm%S",
                        dayfirst=True,
                        errors="coerce"
                    )

                # Conversion du résultat en numérique (cm)
                if "Résultat" in df.columns:
                    df["Résultat"] = pd.to_numeric(df["Résultat"], errors="coerce")

                # Récupération de Z_CC49
                code = site.strip().upper()
                row = summary.loc[summary["Station"] == code, "Z_CC49"]
                if not row.empty and pd.notna(row.iloc[0]):
                    z_ref = float(row.iloc[0])
                    df["Hauteur sable (m)"] = z_ref - df["Résultat"] / 100.0
                    plot_col = "Hauteur sable (m)"
                    y_label = "Hauteur sable (m)"
                    has_ref = True
                else:
                    plot_col = "Résultat"
                    y_label = "Résultat (cm)"
                    has_ref = False
                    z_ref = None

                # Tri par date si possible
                if "Date / Heure" in df.columns:
                    try:
                        df = df.sort_values("Date / Heure")
                    except Exception:
                        pass

                # --- Création du graphique ---
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.set_title(f"Évolution sédimentaire – Site : {site}",
                             pad=15, fontweight="bold")

                if "Date / Heure" in df.columns and plot_col in df.columns:
                    # Tracé des points
                    if "linéaire" in chart_type.lower():
                        ax.scatter(df["Date / Heure"], df[plot_col],
                                   s=50, zorder=3)
                        # droite de tendance entre premier et dernier
                        if len(df) >= 2:
                            x0, y0 = df["Date / Heure"].iloc[0], df[plot_col].iloc[0]
                            x1, y1 = df["Date / Heure"].iloc[-1], df[plot_col].iloc[-1]
                            ax.plot([x0, x1], [y0, y1],
                                    linestyle="--", linewidth=1.5,
                                    color="orange", label="Tendance extrêmes",
                                    zorder=2)
                        ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.6), frameon=False)

                    else:
                        dates = df["Date / Heure"].dt.strftime("%d/%m/%Y")
                        ax.bar(dates, df[plot_col],
                               alpha=0.7, zorder=2)

                    ax.set_xlabel("Date")
                    ax.set_ylabel(y_label)
                    ax.tick_params(axis="x", rotation=45)

                    # grille légère
                    ax.grid(which="major", linestyle=":", alpha=0.5)

                    # axe Y de 0 à Z_CC49+1 si dispo
                    if has_ref:
                        ax.set_ylim(0, z_ref + 1)
                    else:
                        ax.set_ylim(bottom=0)

                    # ligne de référence et légende à droite
                    if has_ref:
                        ax.axhline(
                            y=z_ref,
                            linestyle="--",
                            label="Hauteur poteau réf. CC49",
                            color="C1"
                        )
                        ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False)

                # --- Ajout des logos ---
                for pos, path in zip(
                    [(0.02, 0.92, 0.10, 0.08),
                     (0.78, 0.92, 0.10, 0.08),
                     (0.90, 0.92, 0.10, 0.08)],
                    logo_paths.values()
                ):
                    try:
                        img = mpimg.imread(path)
                        ax_img = fig.add_axes(pos, zorder=10)
                        ax_img.imshow(img)
                        ax_img.axis("off")
                    except Exception as e:
                        print(f"Erreur logo {os.path.basename(path)} : {e}")

                # --- Sauvegarde ---
                suffix = "line" if "linéaire" in chart_type.lower() else "bar"
                out_name = f"{site}_{suffix}.png"
                out_path = os.path.join(self.save_folder, out_name)
                plt.savefig(out_path, bbox_inches="tight")
                plt.close(fig)
                saved_files.append(out_path)

            except Exception as e:
                print(f"[{site}] erreur graphique : {e}")
                continue

        # --- Message final ---
        if saved_files:
            QMessageBox.information(
                self, "Succès",
                f"{len(saved_files)} graphiques générés dans :\n{self.save_folder}"
            )
        else:
            QMessageBox.warning(self, "Erreur", "Aucun graphique n'a pu être généré.")
