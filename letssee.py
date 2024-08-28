import sys
import os
import fitz  # PyMuPDF
import subprocess
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog,
                             QGraphicsDropShadowEffect, QProgressBar, QHBoxLayout, QFrame)
from PyQt5.QtGui import QFont, QColor, QIcon, QPainter, QPainterPath, QPen, QLinearGradient
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve, QRectF, QRect, QSize

class GlowingProgressBar(QProgressBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTextVisible(False)
        self.setRange(0, 100)
        self.setValue(0)
        self.setFixedHeight(6)
        self.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #1A1A2E;
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background-color: transparent;
            }
        """)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Draw background
        bgPath = QPainterPath()
        bgPath.addRoundedRect(QRectF(self.rect()), 3, 3)
        painter.fillPath(bgPath, QColor("#1A1A2E"))
        
        # Draw glowing bar
        progress = self.value() / 100
        bar_width = int(self.width() * progress)
        glow_rect = QRect(0, 0, bar_width, self.height())
        
        gradient = QLinearGradient(0, 0, bar_width, 0)
        gradient.setColorAt(0, QColor(0, 180, 216, 150))  # Start color (semi-transparent)
        gradient.setColorAt(1, QColor(0, 180, 216, 255))  # End color (fully opaque)
        
        painter.setBrush(gradient)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(glow_rect, 3, 3)
        
        # Draw glow effect
        glow = QColor(0, 180, 216, 50)
        for i in range(5):
            painter.setBrush(Qt.NoBrush)
            painter.setPen(QPen(glow, i))
            painter.drawRoundedRect(glow_rect.adjusted(-i, -i, i, i), 3 + i, 3 + i)

class AestheticButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                background-color: #00B4D8;
                color: white;
                border: none;
                padding: 10px 20px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #48CAE4;
            }
            QPushButton:pressed {
                background-color: #0096C7;
                padding: 11px 19px;
            }
        """)
        self.setCursor(Qt.PointingHandCursor)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setXOffset(0)
        shadow.setYOffset(5)
        shadow.setColor(QColor(0, 0, 0, 80))
        self.setGraphicsEffect(shadow)

class PDFProcessorThread(QThread):
    progress_update = pyqtSignal(int)
    status_update = pyqtSignal(str)
    finished = pyqtSignal(str)

    def __init__(self, input_pdf_path, output_folder):
        super().__init__()
        self.input_pdf_path = input_pdf_path
        self.output_folder = output_folder

    def run(self):
        try:
            pdf_document = fitz.open(self.input_pdf_path)
            total_pages = len(pdf_document)

            for page_number in range(total_pages):
                page = pdf_document[page_number]
                page_width = page.rect.width

                # Search for "Question number" on the page
                text_instances = page.search_for("Question number")

                # For each instance of "Question number", draw double lines across the page
                for rect in text_instances:
                    offset = 10
                    line_gap = 2

                    line1_start = fitz.Point(0, rect.y0 - offset)
                    line1_end = fitz.Point(page_width, rect.y0 - offset)
                    line2_start = fitz.Point(0, rect.y0 - offset - line_gap)
                    line2_end = fitz.Point(page_width, rect.y0 - offset - line_gap)

                    # Draw the lines with a gradient color
                    page.draw_line(line1_start, line1_end, color=(0, 0, 0), width=1.5)
                    page.draw_line(line2_start, line2_end, color=(0, 0, 0), width=1.5)

                # Remove ticks and crosses
                images = page.get_images(full=True)
                for img in images:
                    xref = img[0]
                    img_dict = pdf_document.extract_image(xref)
                    if img_dict["width"] == 16 and img_dict["height"] == 16:
                        page.delete_image(xref)

                progress = int((page_number + 1) / total_pages * 100)
                self.progress_update.emit(progress)

            # Save the modified PDF
            output_pdf_path = os.path.join(self.output_folder, f"{os.path.splitext(os.path.basename(self.input_pdf_path))[0]}_removed_ticks_and_crosses.pdf")
            pdf_document.save(output_pdf_path)
            pdf_document.close()

            self.status_update.emit("✨ Tick & Cross removed from the file.")
            self.finished.emit(output_pdf_path)
        except Exception as e:
            self.status_update.emit(f"Error: {str(e)}")

class PDFModifierApp(QWidget):
    def __init__(self):
        super().__init__()
        self.output_folder = os.path.expanduser("~/Downloads")
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PDF Modifier')
        self.setGeometry(100, 100, 600, 400)
        self.setWindowIcon(QIcon('pdf_icon.png'))

        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #1A1A2E, stop:1 #16213E);
                color: #E0E1DD;
                font-family: 'Segoe UI', Arial;
            }
            QLabel {
                color: #90E0EF;
                font-weight: bold;
            }
        """)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(40, 40, 40, 40)

        self.title_label = QLabel('🌊 PDF Tick & Cross Remover')
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("font-size: 24pt; color: #CAF0F8; font-weight: bold; margin-bottom: 20px;")
        main_layout.addWidget(self.title_label)

        content_frame = QFrame()
        content_frame.setStyleSheet("""
            QFrame {
                background-color: rgba(255, 255, 255, 0.1);
                border-radius: 15px;
                padding: 20px;
            }
        """)
        content_layout = QVBoxLayout(content_frame)
        content_layout.setSpacing(15)

        # PDF Selection
        pdf_layout = QHBoxLayout()
        self.pdf_label = QLabel('Select PDF:')
        self.pdf_label.setStyleSheet("font-size: 14pt;")
        pdf_layout.addWidget(self.pdf_label)

        self.pdf_button = AestheticButton('Choose File')
        self.pdf_button.clicked.connect(self.select_pdf)
        pdf_layout.addWidget(self.pdf_button)
        content_layout.addLayout(pdf_layout)

        # Output Folder Selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel('Output folder:')
        self.output_label.setStyleSheet("font-size: 14pt;")
        output_layout.addWidget(self.output_label)

        self.output_button = AestheticButton('Select Folder')
        self.output_button.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_button)
        content_layout.addLayout(output_layout)

        # Output Path Display
        self.output_path_label = QLabel(self.output_folder)
        self.output_path_label.setStyleSheet("color: #E0E1DD; font-size: 12pt; font-weight: normal;")
        self.output_path_label.setWordWrap(True)
        content_layout.addWidget(self.output_path_label)

        # Progress Bar
        self.progress_bar = GlowingProgressBar()
        content_layout.addWidget(self.progress_bar)

        # Status Label
        self.status_label = QLabel('')
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("color: #90E0EF; font-size: 12pt; font-weight: normal;")
        content_layout.addWidget(self.status_label)

        main_layout.addWidget(content_frame)
        self.setLayout(main_layout)

    def select_pdf(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", self.output_folder, "PDF Files (*.pdf);;All Files (*)", options=options)

        if file_path:
            self.modify_pdf(file_path)

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", self.output_folder)
        if folder:
            self.output_folder = folder
            self.output_path_label.setText(self.output_folder)

    def modify_pdf(self, input_pdf_path):
        self.pdf_processor_thread = PDFProcessorThread(input_pdf_path, self.output_folder)
        self.pdf_processor_thread.progress_update.connect(self.update_progress)
        self.pdf_processor_thread.status_update.connect(self.update_status)
        self.pdf_processor_thread.finished.connect(self.conversion_finished)
        self.pdf_processor_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_status(self, message):
        self.status_label.setText(message)

    def conversion_finished(self, output_pdf_path):
        self.status_label.setText(f"✨ Tick & Cross removed from the file and saved to the selected folder.")
        self.convert_to_bw(output_pdf_path)

    def check_printer_installed(self):
        command = "powershell -Command \"& {if (Get-Printer -Name 'Microsoft Print to PDF' -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }}\""
        result = subprocess.run(command, shell=True)
        return result.returncode == 0

    def install_printer(self):
        command = ("powershell -Command \"& {Add-Printer -Name 'Microsoft Print to PDF' -DriverName 'Microsoft Print To PDF' "
                   "-PortName 'PORTPROMPT:' -ErrorAction SilentlyContinue}\"")
        subprocess.run(command, shell=True)

    def convert_to_bw(self, input_pdf):
        if not self.check_printer_installed():
            self.status_label.setText('Microsoft Print to PDF is not installed. Attempting to install...')
            self.install_printer()
            if not self.check_printer_installed():
                self.status_label.setText('Failed to install Microsoft Print to PDF. Please install it manually.')
                return

        output_pdf = os.path.join(self.output_folder, f"{os.path.splitext(os.path.basename(input_pdf))[0]}_BW.pdf")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        powershell_command = (
            f"powershell -ExecutionPolicy Bypass -File \"{script_dir}\\PrintToPDF.ps1\" "
            f"-inputPdf \"{input_pdf}\" -outputPdf \"{output_pdf}\""
        )
        result = subprocess.run(powershell_command, shell=True)
        
        if result.returncode == 0:
            self.status_label.setText(f'PDF has been successfully opened in Edge.')
        else:
            self.status_label.setText('Failed to convert the PDF to black and white.')

    def open_pdf(self, pdf_path):
        if sys.platform.startswith('darwin'):  # macOS
            os.system(f'open "{pdf_path}"')
        elif sys.platform.startswith('win'):   # Windows
            os.system(f'start "" "{pdf_path}"')
        else:  # Linux and other Unix-like
            os.system(f'xdg-open "{pdf_path}"')

def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    window = PDFModifierApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()