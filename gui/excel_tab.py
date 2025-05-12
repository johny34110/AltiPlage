# gui/excel_tab.py
import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QGroupBox, QFormLayout
from functions.excel_utils import create_or_update_excel

class ExcelTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.input_folder = None
        self.output_folder = None
        self.excel_file = None

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

    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier d'entrée")
        if folder:
            self.input_folder = folder
            self.label_input.setText(folder)

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier de sortie")
        if folder:
            self.output_folder = folder
            self.label_output.setText(folder)

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
            if main_window and hasattr(main_window, "update_excel_file_and_folder"):
                main_window.update_excel_file_and_folder(excel_file, self.input_folder)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la génération de l'Excel :\n{str(e)}")
