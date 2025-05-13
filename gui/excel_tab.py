# gui/excel_tab.py

import os
import sqlite3
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QGroupBox, QFormLayout
)
from functions.excel_utils import create_or_update_excel


class ExcelTab(QWidget):
    """
    Onglet pour générer ou mettre à jour le fichier Excel de résultats.
    Sauvegarde et restaure automatiquement les derniers dossiers d'entrée et de sortie
    via une petite base SQLite dans Database/settings.db
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.input_folder = None
        self.output_folder = None
        self.excel_file = None

        # Préparation de la base SQLite de settings
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        db_dir = os.path.join(base_dir, 'Database')
        os.makedirs(db_dir, exist_ok=True)
        self.db_path = os.path.join(db_dir, 'settings.db')
        self._init_db()

        # Construction de l'UI
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        group = QGroupBox("Paramètres")
        form_layout = QFormLayout()
        group.setLayout(form_layout)

        self.label_input = QLabel("Aucun dossier sélectionné")
        form_layout.addRow("Dossier d'entrée (sites) :", self.label_input)
        self.btn_select_input = QPushButton("Sélectionner dossier d'entrée")
        self.btn_select_input.clicked.connect(self.select_input_folder)
        form_layout.addRow(self.btn_select_input)

        self.label_output = QLabel("Aucun dossier sélectionné")
        form_layout.addRow("Dossier de sortie (Excel) :", self.label_output)
        self.btn_select_output = QPushButton("Sélectionner dossier de sortie")
        self.btn_select_output.clicked.connect(self.select_output_folder)
        form_layout.addRow(self.btn_select_output)

        main_layout.addWidget(group)

        self.btn_generate = QPushButton("Générer / Mettre à jour Excel")
        self.btn_generate.clicked.connect(self.generate_excel)
        main_layout.addWidget(self.btn_generate)

        self.status_label = QLabel("")
        main_layout.addWidget(self.status_label)

        # Chargement des réglages
        self._load_settings()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        conn.commit()
        conn.close()

    def _save_setting(self, key, value):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("REPLACE INTO settings(key, value) VALUES(?, ?)" , (key, value))
        conn.commit()
        conn.close()

    def _get_setting(self, key):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT value FROM settings WHERE key=?", (key,))
        row = c.fetchone()
        conn.close()
        return row[0] if row else None

    def _load_settings(self):
        inp = self._get_setting('input_folder')
        out = self._get_setting('output_folder')
        if inp and os.path.isdir(inp):
            self.input_folder = inp
            self.label_input.setText(inp)
        if out and os.path.isdir(out):
            self.output_folder = out
            self.label_output.setText(out)

    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier d'entrée")
        if folder:
            self.input_folder = folder
            self.label_input.setText(folder)
            self._save_setting('input_folder', folder)

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier de sortie")
        if folder:
            self.output_folder = folder
            self.label_output.setText(folder)
            self._save_setting('output_folder', folder)

    def generate_excel(self):
        if not self.input_folder:
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner le dossier d'entrée.")
            return
        if not self.output_folder:
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner le dossier de sortie.")
            return
        try:
            excel_file, msg = create_or_update_excel(self.input_folder, self.output_folder)
            self.excel_file = excel_file
            self.status_label.setText(msg)
            QMessageBox.information(self, "Succès", f"{msg}\nExcel généré : {excel_file}")
            main_window = self.window()
            if main_window and hasattr(main_window, 'update_excel_file_and_folder'):
                main_window.update_excel_file_and_folder(excel_file, self.input_folder)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la génération de l'Excel :\n{e}")
