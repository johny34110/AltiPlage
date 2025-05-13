# gui/measure_tab.py

import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox,
    QGraphicsView, QGraphicsScene, QGraphicsRectItem,
    QDoubleSpinBox, QFileDialog, QLabel
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
        self.selections = []
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)

    def setImage(self, image_path):
        pixmap = QPixmap(image_path)
        if pixmap.isNull():
            print("QPixmap is null for", image_path)
            return False
        if pixmap.width() > 800:
            pixmap = pixmap.scaledToWidth(800, Qt.SmoothTransformation)
        self.scene.clear()
        self.pixmap_item = self.scene.addPixmap(pixmap)
        self.setSceneRect(QRectF(pixmap.rect()))
        self.fitInView(self.pixmap_item, Qt.KeepAspectRatio)
        self.selections = []
        return True

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

        # Rule height and remaining photos
        top_info = QHBoxLayout()
        layout.addLayout(top_info)
        top_info.addWidget(QLabel("Hauteur règle (cm) :"))
        self.rule_height_spin = QDoubleSpinBox()
        self.rule_height_spin.setRange(0.1, 1000.0)
        self.rule_height_spin.setValue(12.0)
        self.rule_height_spin.setDecimals(2)
        top_info.addWidget(self.rule_height_spin)
        top_info.addStretch()
        top_info.addWidget(QLabel("Photos restantes :"))
        self.remaining_label = QLabel("0")
        top_info.addWidget(self.remaining_label)

        # Photo controls
        btn_layout = QHBoxLayout()
        layout.addLayout(btn_layout)
        self.btn_list = QPushButton("Lister sans résultat")
        self.btn_list.clicked.connect(self.list_missing)
        btn_layout.addWidget(self.btn_list)
        self.btn_load = QPushButton("Charger suivante")
        self.btn_load.clicked.connect(self.load_next_photo)
        btn_layout.addWidget(self.btn_load)
        self.btn_open = QPushButton("Ouvrir photo")
        self.btn_open.clicked.connect(self.open_current_photo)
        btn_layout.addWidget(self.btn_open)

        # Image viewer
        self.image_viewer = ImageViewer()
        layout.addWidget(self.image_viewer)

        # Measured height input
        measure_layout = QHBoxLayout()
        layout.addLayout(measure_layout)
        measure_layout.addWidget(QLabel("Hauteur sable (cm) :"))
        self.measure_spin = QDoubleSpinBox()
        self.measure_spin.setRange(0, 10000.0)
        self.measure_spin.setDecimals(2)
        measure_layout.addWidget(self.measure_spin)

        # Selection and action controls
        ctrl = QHBoxLayout()
        layout.addLayout(ctrl)
        self.btn_clear = QPushButton("Réinit sélections")
        self.btn_clear.clicked.connect(self.image_viewer.clearSelections)
        ctrl.addWidget(self.btn_clear)
        self.btn_rotate = QPushButton("Pivoter 90°")
        self.btn_rotate.clicked.connect(lambda: self.image_viewer.rotateImage(90))
        ctrl.addWidget(self.btn_rotate)
        self.btn_calc = QPushButton("Calculer hauteur")
        self.btn_calc.clicked.connect(self.calculate_current_height)
        ctrl.addWidget(self.btn_calc)
        self.btn_save = QPushButton("Sauvegarder")
        self.btn_save.clicked.connect(self.save_current_result)
        ctrl.addWidget(self.btn_save)

        # Instruction label
        self.instruction_label = QLabel("Mode : tracez la règle (rouge)")
        layout.addWidget(self.instruction_label)

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
                    missing.append((sheet, idx+2, df.at[idx, 'Nom de la photo']))
        self.missing_photos = missing
        self.remaining_label.setText(str(len(missing)))
        QMessageBox.information(self, "Info", f"{len(missing)} photo(s) sans résultat.")

    def load_next_photo(self):
        if not self.missing_photos:
            QMessageBox.information(self, "Info", "Aucune photo à charger.")
            return
        sheet, row, photo = self.missing_photos.pop(0)
        self.remaining_label.setText(str(len(self.missing_photos)))
        path = os.path.join(self.input_folder, sheet, photo)
        if os.path.exists(path):
            self.current_photo = (sheet, row, photo)
            self.current_photo_path = path
            if not self.image_viewer.setImage(path):
                QMessageBox.warning(self, "Erreur", "Impossible de charger l'image.")
            self.image_viewer.clearSelections()
            self.calculated_value = None
            self.measure_spin.setValue(0)
            self.instruction_label.setText("Mode : tracez la règle (rouge)")
        else:
            QMessageBox.warning(self, "Erreur", f"Fichier absent : {path}")

    def calculate_current_height(self):
        if len(self.image_viewer.selections) < 2:
            QMessageBox.warning(self, "Attention", "Sélectionnez la règle puis le piquet.")
            return
        ruler_rect, piquet_rect = self.image_viewer.selections
        ruler_cm = self.rule_height_spin.value()
        result_cm = calculate_height(ruler_rect, piquet_rect, ruler_cm, fov_deg=0, image_width_px=0)
        if result_cm is None:
            QMessageBox.warning(self, "Erreur", "Calcul impossible.")
            return
        self.calculated_value = result_cm
        # Affiche la valeur calculée dans le spinbox pour modification manuelle
        self.measure_spin.setValue(round(result_cm, 2))
        self.instruction_label.setText("Ajustez la valeur puis cliquez sur Sauvegarder")

    def save_current_result(self):
        if self.calculated_value is None or not self.current_photo:
            QMessageBox.warning(self, "Attention", "Rien à sauvegarder.")
            return
        sheet, row, photo = self.current_photo
        to_save = self.measure_spin.value()
        update_excel_result(self.excel_file, sheet, row, to_save)
        QMessageBox.information(self, "Sauvegardé", f"Mesure enregistrée pour {photo}.")
        self.load_next_photo()

    def open_current_photo(self):
        if self.current_photo_path and os.path.exists(self.current_photo_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.current_photo_path))
        else:
            QMessageBox.warning(self, "Erreur", "Chemin indisponible.")
