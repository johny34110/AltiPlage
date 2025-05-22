import os
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox,
    QFileDialog, QMessageBox, QGroupBox, QCheckBox, QHBoxLayout
)

from functions.result_utilis import (
    load_summary, load_station_data, load_ram_info
)


class ResultTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_file = None
        self.save_folder = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Analyse générale des résultats"))

        self.btn_select_save = QPushButton("Sélectionner dossier de sauvegarde")
        self.btn_select_save.clicked.connect(self.select_save_folder)
        layout.addWidget(self.btn_select_save)
        self.save_folder_label = QLabel("Aucun dossier sélectionné")
        layout.addWidget(self.save_folder_label)

        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems([
            "Graphique linéaire par station",
            "Graphique en barres par station"
        ])
        layout.addWidget(self.chart_type_combo)

        grp_opts = QGroupBox("Options traces additionnelles")
        hbox = QHBoxLayout()
        self.cb_phma_ngf = QCheckBox("PHMA (m)")
        self.cb_pmve_ngf = QCheckBox("PMVE (m)")
        self.cb_pmme_ngf = QCheckBox("PMME (m)")
        self.cb_nm_ngf   = QCheckBox("NM (m)")
        self.cb_avg      = QCheckBox("Afficher moyenne")
        for cb in (self.cb_phma_ngf, self.cb_pmve_ngf,
                   self.cb_pmme_ngf, self.cb_nm_ngf, self.cb_avg):
            hbox.addWidget(cb)
        grp_opts.setLayout(hbox)
        layout.addWidget(grp_opts)

        self.btn_generate = QPushButton("Générer graphiques")
        self.btn_generate.clicked.connect(self.generate_charts)
        layout.addWidget(self.btn_generate)
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

    def set_excel_file(self, excel_file: str):
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

        try:
            summary = load_summary(self.excel_file)
            ram_info = load_ram_info()
            xls = pd.ExcelFile(self.excel_file)
            sheets = [s for s in xls.sheet_names if s != "Résumé"]
        except Exception as e:
            QMessageBox.critical(self, "Erreur lecture", str(e))
            return

        saved = []
        colors = {"PHMA (m)": "C2", "PMVE (m)": "C3", "PMME (m)": "C4", "NM (m)": "C5"}
        for site in sheets:
            try:
                df = load_station_data(self.excel_file, site)
                code = site.strip().upper()

                info = ram_info.get(code, {})
                z_ref = info.get("Z_CC49")
                site_name = site.strip()
                if z_ref is None:
                    row = summary[summary["Station"] == code]
                    if not row.empty:
                        z_ref = float(row["Z_CC49"].iloc[0])

                if z_ref is not None and "Résultat" in df.columns:
                    df["Hauteur sable (m)"] = z_ref - df["Résultat"] / 100.0
                    plot_col = "Hauteur sable (m)"
                    ylabel = "Hauteur sable (m)"
                else:
                    plot_col = "Résultat"
                    ylabel = "Résultat (cm)"

                if "Date / Heure" in df.columns:
                    df = df.sort_values("Date / Heure")
                mask = pd.notna(df.get("Date / Heure")) & pd.notna(df.get(plot_col))
                df_plot = df.loc[mask]
                dates = df_plot.get("Date / Heure")
                y = df_plot.get(plot_col)

                fig, ax = plt.subplots(figsize=(10, 6))
                ax.set_title(f"Observation hauteur sédimentaire - Station {site_name}", fontweight="bold")

                # Nuage de points plus petit
                if "linéaire" in self.chart_type_combo.currentText().lower():
                    ax.scatter(dates, y, s=20, color="C1", zorder=3, label=plot_col)
                else:
                    dates_str = dates.dt.strftime("%d/%m/%Y")
                    ax.bar(dates_str, y, alpha=0.7, color="C1", zorder=2, label=plot_col)

                # Poteau réf en noir
                if z_ref is not None:
                    ax.axhline(y=z_ref, linestyle='--', color="black", label="Poteau")

                # Seuils NGF en couleurs distinctes
                mapping = {
                    self.cb_phma_ngf: ("PHMA (m NGF)", "PHMA (m)"),
                    self.cb_pmve_ngf: ("PMVE (m NGF)", "PMVE (m)"),
                    self.cb_pmme_ngf: ("PMME (m NGF)", "PMME (m)"),
                    self.cb_nm_ngf:   ("NM (m NGF)",   "NM (m)")
                }
                for cb, (key_ngf, label_simple) in mapping.items():
                    val = info.get(key_ngf)
                    if cb.isChecked() and pd.notna(val):
                        ax.axhline(y=val, linestyle=':', color=colors[label_simple], label=label_simple)

                # Moyenne
                if self.cb_avg.isChecked() and len(df_plot) > 0:
                    m = y.mean()
                    ax.axhline(y=m, linestyle='-.', color="C6", label=f"Moyenne {plot_col}")

                # Mise en forme et légende avec titre
                ax.set_xlabel("Date")
                ax.set_ylabel(ylabel)
                ax.grid(which="major", linestyle=":", alpha=0.5)
                legend = ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False)
                legend.set_title("Système de référence IGN69")
                ax.tick_params(axis="x", rotation=45)

                # Logos
                logo_folder = os.path.join(os.getcwd(), "logo")
                for pos, fname in zip(
                    [(0.02, 0.92, 0.10, 0.08), (0.78, 0.92, 0.10, 0.08), (0.90, 0.92, 0.10, 0.08)],
                    ["Altipl4.png", "lamanche.jpg", "cnam.png"]
                ):
                    path = os.path.join(logo_folder, fname)
                    if os.path.exists(path):
                        img = mpimg.imread(path)
                        ax_img = fig.add_axes(pos, zorder=10)
                        ax_img.imshow(img)
                        ax_img.axis("off")

                # Sauvegarde
                suffix = "line" if "linéaire" in self.chart_type_combo.currentText().lower() else "bar"
                out_path = os.path.join(self.save_folder, f"{site}_{suffix}.png")
                fig.savefig(out_path, bbox_inches="tight")
                plt.close(fig)
                saved.append(out_path)

            except Exception as e:
                print(f"[{site}] Erreur : {e}")
                continue

        if saved:
            QMessageBox.information(
                self, "Succès",
                f"{len(saved)} graphiques générés dans :\n{self.save_folder}"
            )
        else:
            QMessageBox.warning(self, "Erreur", "Aucun graphique n'a pu être généré.")
