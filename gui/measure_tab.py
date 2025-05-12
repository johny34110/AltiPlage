# gui/measure_tab.py
import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QMessageBox,
    QGraphicsView, QGraphicsScene, QGraphicsRectItem, QInputDialog, QDialog,
    QFormLayout, QLineEdit, QLabel, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QRectF, QUrl
from PyQt5.QtGui import QPen, QPixmap, QTransform, QDesktopServices
from functions.excel_utils import list_photos_without_result, update_excel_result
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
        if event.angleDelta().y() > 0:
            factor = 1.25
        else:
            factor = 0.8
        self.scale(factor, factor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if len(self.selections) >= 2:
                return
            self.origin = self.mapToScene(event.pos())
            pen = QPen(Qt.red, 2, Qt.SolidLine) if len(self.selections) == 0 else QPen(Qt.blue, 2, Qt.SolidLine)
            self.current_rect_item = self.scene.addRect(QRectF(self.origin, self.origin), pen)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.current_rect_item:
            current_pos = self.mapToScene(event.pos())
            rect = QRectF(self.origin, current_pos).normalized()
            self.current_rect_item.setRect(rect)
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
        for item in self.scene.items():
            if isinstance(item, QGraphicsRectItem):
                self.scene.removeItem(item)
        self.selections = []

    def rotateImage(self, angle):
        if self.pixmap_item:
            try:
                transform = QTransform()
                transform.rotate(angle)
                new_pixmap = self.pixmap_item.pixmap().transformed(transform, Qt.SmoothTransformation)
                self.pixmap_item.setPixmap(new_pixmap)
                self.clearSelections()
            except Exception as e:
                print("Erreur lors de la rotation :", e)


class MeasureTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_file = None
        self.input_folder = None
        self.missing_photos = []  # Liste de tuples (sheet_name, row_idx, photo_name)
        self.current_photo = None
        self.current_photo_path = None
        self.calculated_value = None

        main_layout = QVBoxLayout(self)
        self.setLayout(main_layout)

        self.instruction_label = QLabel("Mode: Sélectionnez la zone de la RÈGLE (rouge)")
        main_layout.addWidget(self.instruction_label)

        top_layout = QHBoxLayout()
        self.btn_list = QPushButton("Lister photos sans résultat")
        self.btn_list.clicked.connect(self.list_missing)
        top_layout.addWidget(self.btn_list)

        self.btn_load_next = QPushButton("Charger photo suivante")
        self.btn_load_next.clicked.connect(self.load_next_photo)
        top_layout.addWidget(self.btn_load_next)

        self.btn_open_path = QPushButton("Ouvrir chemin photo")
        self.btn_open_path.clicked.connect(self.open_current_photo)
        top_layout.addWidget(self.btn_open_path)
        main_layout.addLayout(top_layout)

        self.image_viewer = ImageViewer()
        main_layout.addWidget(self.image_viewer)

        ctrl_layout = QHBoxLayout()
        self.btn_clear = QPushButton("Réinitialiser sélections")
        self.btn_clear.clicked.connect(self.image_viewer.clearSelections)
        ctrl_layout.addWidget(self.btn_clear)

        self.btn_rotate = QPushButton("Pivoter 90°")
        self.btn_rotate.clicked.connect(lambda: self.image_viewer.rotateImage(90))
        ctrl_layout.addWidget(self.btn_rotate)

        self.btn_calculate = QPushButton("Calculer hauteur")
        self.btn_calculate.clicked.connect(self.calculate_current_height)
        ctrl_layout.addWidget(self.btn_calculate)

        self.btn_save = QPushButton("Sauvegarder résultat")
        self.btn_save.clicked.connect(self.save_current_result)
        ctrl_layout.addWidget(self.btn_save)
        main_layout.addLayout(ctrl_layout)

    def set_excel_file_and_folder(self, excel_file, input_folder):
        print(f"Réception fichier Excel dans MeasureTab : {excel_file}")
        self.excel_file = excel_file
        self.input_folder = input_folder

    def list_missing(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Attention", "Aucun fichier Excel chargé.")
            return
        self.missing_photos = list_photos_without_result(self.excel_file)
        if not self.missing_photos:
            QMessageBox.information(self, "Info", "Aucune photo sans résultat.")
        else:
            QMessageBox.information(self, "Info", f"{len(self.missing_photos)} photo(s) à mesurer.")

    def load_next_photo(self):
        while self.missing_photos:
            self.current_photo = self.missing_photos.pop(0)
            sheet_name, row_idx, photo_name = self.current_photo
            photo_path = os.path.normpath(os.path.join(self.input_folder, sheet_name, photo_name))
            if os.path.exists(photo_path):
                self.current_photo_path = photo_path
                if not self.image_viewer.setImage(photo_path):
                    QMessageBox.warning(self, "Erreur", "Impossible de charger l'image.")
                self.image_viewer.clearSelections()
                self.calculated_value = None
                self.instruction_label.setText("Mode: Sélectionnez la zone de la RÈGLE (rouge)")
                return
            else:
                print(f"Image non trouvée: {photo_path}")
                continue
        QMessageBox.information(self, "Terminé", "Toutes les photos ont été mesurées.")

    def calculate_current_height(self):
        if len(self.image_viewer.selections) < 2:
            QMessageBox.warning(self, "Attention", "Sélectionnez d'abord la zone de la RÈGLE, puis celle du PIQUET.")
            return
        if len(self.image_viewer.selections) == 1:
            self.instruction_label.setText("Mode: Sélectionnez maintenant la zone du PIQUET (bleu)")
            QMessageBox.information(self, "Info", "Tracez la zone du PIQUET (en bleu) et réessayez.")
            return
        ruler_rect = self.image_viewer.selections[0]
        piquet_rect = self.image_viewer.selections[1]
        ruler_height_cm, ok = QInputDialog.getDouble(
            self, "Hauteur de la règle",
            "Entrez la hauteur réelle de la règle (en cm) :", decimals=2)
        if not ok or ruler_height_cm <= 0:
            QMessageBox.warning(self, "Erreur", "La hauteur de la règle doit être positive.")
            return
        result = calculate_height(ruler_rect, piquet_rect, ruler_height_cm, fov_deg=0, image_width_px=0)
        if result is None:
            QMessageBox.warning(self, "Erreur", "Erreur lors du calcul de la hauteur.")
        else:
            modified_result, ok = QInputDialog.getDouble(
                self, "Modifier la hauteur",
                "Hauteur calculée (modifiez si besoin) :", result, min=0, decimals=2)
            if ok:
                self.calculated_value = modified_result
                QMessageBox.information(self, "Résultat", f"Hauteur finale : {self.calculated_value:.2f} cm")

    def save_current_result(self):
        if self.calculated_value is None:
            QMessageBox.warning(self, "Attention", "Aucune mesure calculée.")
            return
        if not self.current_photo:
            QMessageBox.warning(self, "Attention", "Aucune photo chargée.")
            return
        sheet_name, row_idx, photo_name = self.current_photo
        update_excel_result(self.excel_file, sheet_name, row_idx, self.calculated_value)
        QMessageBox.information(self, "Sauvegardé", f"Résultat sauvegardé pour {photo_name}.")
        self.image_viewer.clearSelections()
        self.calculated_value = None
        self.load_next_photo()

    def open_current_photo(self):
        if self.current_photo_path and os.path.exists(self.current_photo_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.current_photo_path))
        else:
            QMessageBox.warning(self, "Erreur", "Chemin de la photo non disponible.")


from PyQt5.QtWidgets import QInputDialog
