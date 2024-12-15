import sys
import os
import tempfile
from urllib.parse import unquote
from PyQt5 import QtWidgets, QtCore, QtGui
import fitz  # PyMuPDF
from pikepdf import Pdf, Array

class PagePreviewWidget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.doc = None
        self.page_num = None
        self.original_qimage = None
        self.zoom_factor = 1.0  # Default zoom (100%)

        self.scroll_area = QtWidgets.QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #f0f0f0;")

        self.image_label = QtWidgets.QLabel("No page selected")
        self.image_label.setAlignment(QtCore.Qt.AlignCenter)
        self.image_label.setFrameShape(QtWidgets.QFrame.Box)
        self.image_label.setLineWidth(1)
        self.image_label.setScaledContents(False)
        self.image_label.setBackgroundRole(QtGui.QPalette.Base)

        self.scroll_area.setWidget(self.image_label)

        self.top_cut_spin = QtWidgets.QDoubleSpinBox()
        self.top_cut_spin.setRange(0, 99.9)
        self.top_cut_spin.setSingleStep(1)
        self.top_cut_spin.valueChanged.connect(self.update_preview)

        self.bottom_cut_spin = QtWidgets.QDoubleSpinBox()
        self.bottom_cut_spin.setRange(0, 99.9)
        self.bottom_cut_spin.setSingleStep(1)
        self.bottom_cut_spin.valueChanged.connect(self.update_preview)

        self.zoom_spin = QtWidgets.QDoubleSpinBox()
        self.zoom_spin.setRange(0.1, 5.0)
        self.zoom_spin.setSingleStep(0.1)
        self.zoom_spin.setValue(1.0)
        self.zoom_spin.valueChanged.connect(self.reload_page_with_zoom)

        controls_layout = QtWidgets.QHBoxLayout()
        controls_layout.addWidget(QtWidgets.QLabel("Top cut (%):"))
        controls_layout.addWidget(self.top_cut_spin)
        controls_layout.addWidget(QtWidgets.QLabel("Bottom cut (%):"))
        controls_layout.addWidget(self.bottom_cut_spin)
        controls_layout.addWidget(QtWidgets.QLabel("Zoom:"))
        controls_layout.addWidget(self.zoom_spin)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.scroll_area, stretch=1)
        layout.addLayout(controls_layout)
        self.setLayout(layout)

    def load_page(self, doc, page_num, top_cut=0, bottom_cut=0):
        self.doc = doc
        self.page_num = page_num
        self.top_cut_spin.setValue(top_cut)
        self.bottom_cut_spin.setValue(bottom_cut)
        self.render_page()

    def reload_page_with_zoom(self):
        if self.doc is not None and self.page_num is not None:
            self.zoom_factor = self.zoom_spin.value()
            self.render_page()

    def render_page(self):
        page = self.doc.load_page(self.page_num)
        matrix = fitz.Matrix(self.zoom_factor, self.zoom_factor)
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        img_data = pix.tobytes("png")
        qimg = QtGui.QImage.fromData(img_data, "PNG")
        self.original_qimage = qimg.copy()
        self.update_preview()

    def update_preview(self):
        if self.original_qimage is None:
            return

        t = self.top_cut_spin.value()
        b = self.bottom_cut_spin.value()

        orig_img = self.original_qimage
        height = orig_img.height()
        width = orig_img.width()

        top_px = int((t / 100.0) * height)
        bottom_px = int((b / 100.0) * height)

        # Create a copy for display
        display_img = QtGui.QImage(orig_img)
        painter = QtGui.QPainter(display_img)
        painter.setCompositionMode(QtGui.QPainter.CompositionMode_SourceOver)
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(QtGui.QColor("#ffffcc"))

        if top_px > 0:
            painter.drawRect(0, 0, width, top_px)
        if bottom_px > 0:
            painter.drawRect(0, height - bottom_px, width, bottom_px)

        painter.end()

        self.image_label.setPixmap(QtGui.QPixmap.fromImage(display_img))
        self.image_label.resize(display_img.width(), display_img.height())

    def get_cuts(self):
        return (self.top_cut_spin.value(), self.bottom_cut_spin.value())


class PDFTrimApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Trimmer")

        self.pdf_path = None
        self.doc = None
        self.selected_pages = []
        self.page_trim = {}  # page_num: (top_cut, bottom_cut)
        self.current_page = None

        main_layout = QtWidgets.QVBoxLayout(self)

        path_layout = QtWidgets.QHBoxLayout()
        self.pdf_path_edit = QtWidgets.QLineEdit(self)
        load_btn = QtWidgets.QPushButton("Load")
        load_btn.clicked.connect(self.load_pdf_and_pages)
        path_layout.addWidget(QtWidgets.QLabel("PDF Path:"))
        path_layout.addWidget(self.pdf_path_edit)
        path_layout.addWidget(load_btn)

        range_layout = QtWidgets.QHBoxLayout()
        self.start_page_edit = QtWidgets.QSpinBox()
        self.start_page_edit.setMinimum(1)
        # Remove maximum limit
        self.start_page_edit.setMaximum(999999)  # Just a large number
        self.end_page_edit = QtWidgets.QSpinBox()
        self.end_page_edit.setMinimum(1)
        # Remove maximum limit
        self.end_page_edit.setMaximum(999999)  # Just a large number
        range_layout.addWidget(QtWidgets.QLabel("Start Page:"))
        range_layout.addWidget(self.start_page_edit)
        range_layout.addWidget(QtWidgets.QLabel("End Page:"))
        range_layout.addWidget(self.end_page_edit)

        splitter = QtWidgets.QSplitter()
        self.pages_list = QtWidgets.QListWidget()
        self.pages_list.itemSelectionChanged.connect(self.page_selected)
        self.preview_widget = PagePreviewWidget()

        splitter.addWidget(self.pages_list)
        splitter.addWidget(self.preview_widget)
        splitter.setStretchFactor(1, 1)

        bottom_layout = QtWidgets.QHBoxLayout()
        self.export_btn = QtWidgets.QPushButton("Export & Copy to Clipboard")
        self.export_btn.clicked.connect(self.export_pdf)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.export_btn)

        main_layout.addLayout(path_layout)
        main_layout.addLayout(range_layout)
        main_layout.addWidget(splitter)
        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

    def load_pdf_and_pages(self):
        path = self.pdf_path_edit.text().strip()
        # Handle file:/// prefix
        if path.startswith("file:///"):
            path = path[8:]  # Remove 'file:///'
            path = unquote(path)  # Decode URL-encoded characters

        if not path or not os.path.exists(path):
            QtWidgets.QMessageBox.warning(self, "Error", "File does not exist.")
            return

        try:
            self.doc = fitz.open(path)
            self.pdf_path = path
            self.load_pages_range()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to open PDF: {e}")
            self.doc = None

    def load_pages_range(self):
        if self.doc is None:
            QtWidgets.QMessageBox.warning(self, "Error", "No PDF loaded.")
            return

        total_pages = len(self.doc)
        start = self.start_page_edit.value() - 1
        end = self.end_page_edit.value() - 1
        if start < 0 or end < 0 or start > end or end >= total_pages:
            QtWidgets.QMessageBox.warning(self, "Error", f"Invalid page range. PDF has {total_pages} pages.")
            return

        # Clear previous trims
        self.page_trim.clear()
        self.selected_pages = list(range(start, end + 1))
        for p in self.selected_pages:
            self.page_trim[p] = (0, 0)

        self.current_page = None
        self.refresh_pages_list()

    def refresh_pages_list(self):
        self.pages_list.clear()
        for p in self.selected_pages:
            self.pages_list.addItem(f"Page {p+1}")

        if not self.selected_pages:
            self.preview_widget.image_label.setText("No page selected")

    def page_selected(self):
        items = self.pages_list.selectedItems()
        if not items or self.doc is None:
            return

        # Before switching to new page, save current page's cuts
        if self.current_page is not None:
            cur_t, cur_b = self.preview_widget.get_cuts()
            self.page_trim[self.current_page] = (cur_t, cur_b)

        item_text = items[0].text()  # "Page X"
        page_num = int(item_text.split()[1]) - 1
        top_cut, bottom_cut = self.page_trim.get(page_num, (0, 0))
        self.preview_widget.load_page(self.doc, page_num, top_cut, bottom_cut)
        self.current_page = page_num

    def export_pdf(self):
        if not self.selected_pages:
            QtWidgets.QMessageBox.warning(self, "Error", "No pages selected.")
            return

        if self.current_page is not None:
            t, b = self.preview_widget.get_cuts()
            self.page_trim[self.current_page] = (t, b)

        try:
            with Pdf.open(self.pdf_path) as pdf:
                new_pdf = Pdf.new()
                for p in self.selected_pages:
                    page = pdf.pages[p]
                    mediabox = page.mediabox
                    # Convert Decimal to float
                    x0, y0, x1, y1 = map(float, mediabox)

                    width = x1 - x0
                    height = y1 - y0
                    t, b = self.page_trim[p]
                    t = float(t)
                    b = float(b)

                    top_cut_abs = (t/100.0)*height
                    bottom_cut_abs = (b/100.0)*height
                    new_y0 = y0 + bottom_cut_abs
                    new_y1 = y1 - top_cut_abs
                    if new_y1 <= new_y0:
                        new_y1 = new_y0 + 1

                    page.CropBox = [x0, new_y0, x1, new_y1]
                    new_pdf.pages.append(page)

                # Save to a temporary file
                temp_dir = tempfile.gettempdir()
                temp_pdf_path = os.path.join(temp_dir, "trimmed_output.pdf")
                new_pdf.save(temp_pdf_path)

                self.copy_file_to_clipboard(temp_pdf_path)
                QtWidgets.QMessageBox.information(self, "Success", "New PDF copied to clipboard!")

        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to export PDF: {e}")

    def copy_file_to_clipboard(self, file_path):
        mime_data = QtCore.QMimeData()
        url = QtCore.QUrl.fromLocalFile(file_path)
        mime_data.setUrls([url])
        QtWidgets.QApplication.clipboard().setMimeData(mime_data)

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = PDFTrimApp()
    window.resize(1200, 800)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
