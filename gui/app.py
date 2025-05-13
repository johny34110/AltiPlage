# gui/app.py

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from gui.excel_tab import ExcelTab
from gui.measure_tab import MeasureTab
from gui.result_tab import ResultTab


def run_app():
    """Démarre l'application principale avec les onglets Excel, Mesure et Résultat."""
    app = QApplication(sys.argv)
    main = QMainWindow()
    main.setWindowTitle("Altiplage")
    tabs = QTabWidget()

    excel_tab = ExcelTab()
    measure_tab = MeasureTab()
    result_tab = ResultTab()

    # Connecter les callbacks
    def update_excel(excel_file, input_folder):
        measure_tab.set_excel_file_and_folder(excel_file, input_folder)
        result_tab.set_excel_file(excel_file)

    main.update_excel_file_and_folder = update_excel

    tabs.addTab(excel_tab, "Excel")
    tabs.addTab(measure_tab, "Mesure")
    tabs.addTab(result_tab, "Résultats")

    main.setCentralWidget(tabs)
    main.show()
    sys.exit(app.exec_())
