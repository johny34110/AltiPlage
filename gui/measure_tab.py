# gui/measure_tab.py

import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox,
    QGraphicsView, QGraphicsScene, QGraphicsRectItem,
    QInputDialog, QFileDialog, QLabel
)
from PyQt5.QtCore import Qt, QRectF, QUrl
from PyQt5.QtGui import QPen, QPixmap, QTransform, QDesktopServices
from functions.excel_utils import update_excel_result
from functions.measure_utils import calculate_height


class ImageViewer(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.pixmap_item = None
        self.current_rect_item = None
        self.selections = []  # Liste des zones sélectionnées (QRectF)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)

    def setImage(self, image_path):
        try:
            pixmap = QPixmap(image_path)
            if pixmap.isNull():
                print("QPixmap is null for", image_path)
                return False
            max_width = 800
            if pixmap.width() > max_width:
                pixmap = pixmap.scaledToWidth(max_width, Qt.SmoothTransformation)
            self.scene.clear()
            self.pixmap_item = self.scene.addPixmap(pixmap)
            self.setSceneRect(QRectF(pixmap.rect()))
            self.fitInView(self.pixmap_item, Qt.KeepAspectRatio)
            self.selections = []
            return True
        except Exception as e:
            print("Erreur dans setImage:", e)
            return False

    def wheelEvent(self, event):
        factor = 1.25 if event.angleDelta().y() > 0 else 0.8
        self.scale(factor, factor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and len(self.selections) < 2:
            self.origin = self.mapToScene(event.pos())
            pen = QPen(Qt.red, 2) if len(self.selections) == 0 else QPen(Qt.blue, 2)
            self.current_rect_item = self.scene.addRect(QRectF(self.origin, self.origin), pen)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.current_rect_item:
            pos = self.mapToScene(event.pos())
            self.current_rect_item.setRect(QRectF(self.origin, pos).normalized())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.current_rect_item:
            rect = self.current_rect_item.rect()
            if rect.width() > 2 and rect.height() > 2:
                self.selections.append(rect)
            else:
                self.scene.removeItem(self.current_rect_item)
            self.current_rect_item = None
        super().mouseReleaseEvent(event)

    def clearSelections(self):
        for item in list(self.scene.items()):
            if isinstance(item, QGraphicsRectItem):
                self.scene.removeItem(item)
        self.selections = []

    def rotateImage(self, angle):
        if self.pixmap_item:
            transform = QTransform().rotate(angle)
            new_pixmap = self.pixmap_item.pixmap().transformed(transform, Qt.SmoothTransformation)
            self.pixmap_item.setPixmap(new_pixmap)
            self.clearSelections()


class MeasureTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_file = None
        self.input_folder = None
        self.missing_photos = []
        self.current_photo = None
        self.current_photo_path = None
        self.calculated_value = None

        layout = QVBoxLayout(self)
        self.setLayout(layout)

        self.instruction_label = QLabel("Mode: Sélectionnez la zone de la RÈGLE (rouge)")
        layout.addWidget(self.instruction_label)

        btn_layout = QHBoxLayout()
        self.btn_list = QPushButton("Lister photos sans résultat")
        self.btn_list.clicked.connect(self.list_missing)
        btn_layout.addWidget(self.btn_list)

        self.btn_load = QPushButton("Charger photo suivante")
        self.btn_load.clicked.connect(self.load_next_photo)
        btn_layout.addWidget(self.btn_load)

        self.btn_open = QPushButton("Ouvrir chemin photo")
        self.btn_open.clicked.connect(self.open_current_photo)
        btn_layout.addWidget(self.btn_open)

        layout.addLayout(btn_layout)

        self.image_viewer = ImageViewer()
        layout.addWidget(self.image_viewer)

        ctrl = QHBoxLayout()
        self.btn_clear = QPushButton("Réinitialiser sélections")
        self.btn_clear.clicked.connect(self.image_viewer.clearSelections)
        ctrl.addWidget(self.btn_clear)

        self.btn_rotate = QPushButton("Pivoter 90°")
        self.btn_rotate.clicked.connect(lambda: self.image_viewer.rotateImage(90))
        ctrl.addWidget(self.btn_rotate)

        self.btn_calc = QPushButton("Calculer hauteur")
        self.btn_calc.clicked.connect(self.calculate_current_height)
        ctrl.addWidget(self.btn_calc)

        self.btn_save = QPushButton("Sauvegarder résultat")
        self.btn_save.clicked.connect(self.save_current_result)
        ctrl.addWidget(self.btn_save)

        layout.addLayout(ctrl)

    def set_excel_file_and_folder(self, excel_file, input_folder):
        self.excel_file = excel_file
        self.input_folder = input_folder

    def list_missing(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Attention", "Aucun fichier Excel chargé.")
            return

        xls = pd.ExcelFile(self.excel_file)
        missing = []
        for sheet in xls.sheet_names:
            if sheet == 'Résumé':
                continue
            df = pd.read_excel(self.excel_file, sheet_name=sheet)
            if 'Nom de la photo' not in df.columns or 'Résultat' not in df.columns:
                continue
            for idx, val in df['Résultat'].items():
                if pd.isna(val) or str(val).strip() == '':
                    photo = df.at[idx, 'Nom de la photo']
                    missing.append((sheet, idx+2, photo))
        self.missing_photos = missing

        if not missing:
            QMessageBox.information(self, "Info", "Aucune photo sans résultat.")
        else:
            QMessageBox.information(self, "Info", f"{len(missing)} photo(s) à mesurer.")

    def load_next_photo(self):
        while self.missing_photos:
            sheet, row, photo = self.missing_photos.pop(0)
            path = os.path.join(self.input_folder, sheet, photo)
            if os.path.exists(path):
                self.current_photo = (sheet, row, photo)
                self.current_photo_path = path
                if not self.image_viewer.setImage(path):
                    QMessageBox.warning(self, "Erreur", "Impossible de charger l'image.")
                self.image_viewer.clearSelections()
                self.calculated_value = None
                self.instruction_label.setText("Mode: Sélectionnez la zone de la RÈGLE (rouge)")
                return
        QMessageBox.information(self, "Terminé", "Toutes les photos ont été mesurées.")

    def calculate_current_height(self):
        if len(self.image_viewer.selections) < 2:
            QMessageBox.warning(self, "Attention", "Sélectionnez la règle puis le piquet.")
            return
        ruler, piquet = self.image_viewer.selections
        ruler_cm, ok = QInputDialog.getDouble(self, "Hauteur règle", "Hauteur règle (cm):", decimals=2)
        if not ok or ruler_cm <= 0:
            QMessageBox.warning(self, "Erreur", "Hauteur doit être positive.")
            return
        res = calculate_height(ruler, piquet, ruler_cm, fov_deg=0, image_width_px=0)
        if res is None:
            QMessageBox.warning(self, "Erreur", "Calcul impossible.")
            return
        val, ok = QInputDialog.getDouble(self, "Modifier hauteur", "Hauteur (cm):", res, 0, decimals=2)
        if ok:
            self.calculated_value = val
            QMessageBox.information(self, "Résultat", f"Hauteur finale: {val:.2f} cm")

    def save_current_result(self):
        if self.calculated_value is None or not self.current_photo:
            QMessageBox.warning(self, "Attention", "Aucune mesure à sauvegarder.")
            return
        sheet, row, photo = self.current_photo
        update_excel_result(self.excel_file, sheet, row, self.calculated_value)
        QMessageBox.information(self, "Sauvegardé", f"Mesure sauvegardée pour {photo}.")
        self.image_viewer.clearSelections()
        self.calculated_value = None
        self.load_next_photo()

    def open_current_photo(self):
        if self.current_photo_path and os.path.exists(self.current_photo_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.current_photo_path))
        else:
            QMessageBox.warning(self, "Erreur", "Chemin photo indisponible.")
