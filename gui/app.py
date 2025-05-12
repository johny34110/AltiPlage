# gui/app.py
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from gui.excel_tab import ExcelTab
from gui.measure_tab import MeasureTab
from gui.result_tab import ResultTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestion des Photos - Excel, Mesure et Résultats")
        self.resize(1200, 800)
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.excel_tab = ExcelTab(self)
        self.measure_tab = MeasureTab(self)
        self.result_tab = ResultTab(self)
        self.tabs.addTab(self.excel_tab, "Excel")
        self.tabs.addTab(self.measure_tab, "Mesure Hauteur")
        self.tabs.addTab(self.result_tab, "Résultats")

    def update_excel_file_and_folder(self, excel_file, input_folder):
        print(f"Fichier Excel reçu par MainWindow : {excel_file}")
        print(f"Dossier d'entrée reçu par MainWindow : {input_folder}")
        self.measure_tab.set_excel_file_and_folder(excel_file, input_folder)
        self.result_tab.set_excel_file(excel_file)


def run_app():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    run_app()
