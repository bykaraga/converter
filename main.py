import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from converter import FileConverter

class FileConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dosya Dönüştürücü")
        self.setMinimumSize(600, 400)
        
        # Ana widget ve layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Başlık etiketi
        title_label = QLabel("Dosya Dönüştürücü")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; margin: 20px;")
        layout.addWidget(title_label)
        
        # Dosya seçme butonu
        self.select_file_btn = QPushButton("Dosya Seç")
        self.select_file_btn.clicked.connect(self.select_file)
        layout.addWidget(self.select_file_btn)
        
        # Sürükle-bırak alanı
        self.drop_label = QLabel("Veya dosyaları buraya sürükleyin")
        self.drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_label.setStyleSheet("border: 2px dashed #aaa; padding: 20px;")
        layout.addWidget(self.drop_label)
        
        # Dönüştürme butonları
        self.excel_to_pdf_btn = QPushButton("Excel → PDF")
        self.word_to_pdf_btn = QPushButton("Word → PDF")
        self.pdf_to_word_btn = QPushButton("PDF → Word")
        self.image_to_pdf_btn = QPushButton("Görsel → PDF")
        
        self.excel_to_pdf_btn.clicked.connect(lambda: self.convert_file("excel_to_pdf"))
        self.word_to_pdf_btn.clicked.connect(lambda: self.convert_file("word_to_pdf"))
        self.pdf_to_word_btn.clicked.connect(lambda: self.convert_file("pdf_to_word"))
        self.image_to_pdf_btn.clicked.connect(lambda: self.convert_file("image_to_pdf"))
        
        layout.addWidget(self.excel_to_pdf_btn)
        layout.addWidget(self.word_to_pdf_btn)
        layout.addWidget(self.pdf_to_word_btn)
        layout.addWidget(self.image_to_pdf_btn)
        
        # Sürükle-bırak özelliğini etkinleştir
        self.setAcceptDrops(True)
        
        # Seçili dosya yolu
        self.selected_file = None
        
    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", "", "Tüm Dosyalar (*.*)")
        if file_name:
            self.selected_file = file_name
            self.drop_label.setText(f"Seçilen dosya: {file_name}")
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    
    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.selected_file = files[0]
            self.drop_label.setText(f"Seçilen dosya: {self.selected_file}")
    
    def convert_file(self, conversion_type):
        if not self.selected_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce bir dosya seçin!")
            return
            
        output_path, _ = QFileDialog.getSaveFileName(self, "Kaydet", "", "PDF Dosyası (*.pdf);;Word Dosyası (*.docx)")
        if not output_path:
            return
            
        converter = FileConverter()
        success = False
        
        if conversion_type == "excel_to_pdf":
            success = converter.excel_to_pdf(self.selected_file, output_path)
        elif conversion_type == "word_to_pdf":
            success = converter.word_to_pdf(self.selected_file, output_path)
        elif conversion_type == "pdf_to_word":
            success = converter.pdf_to_word(self.selected_file, output_path)
        elif conversion_type == "image_to_pdf":
            success = converter.image_to_pdf(self.selected_file, output_path)
            
        if success:
            QMessageBox.information(self, "Başarılı", "Dönüştürme işlemi tamamlandı!")
        else:
            QMessageBox.critical(self, "Hata", "Dönüştürme işlemi başarısız oldu!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec()) 