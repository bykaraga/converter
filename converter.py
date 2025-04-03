import pandas as pd
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pdfplumber import open as pdf_open
from pdf2docx import Converter
from PIL import Image
import os

class FileConverter:
    @staticmethod
    def excel_to_pdf(input_path, output_path):
        try:
            # Excel dosyasını oku
            df = pd.read_excel(input_path)
            
            # PDF oluştur
            c = canvas.Canvas(output_path, pagesize=letter)
            width, height = letter
            
            # Başlık
            c.setFont("Helvetica-Bold", 16)
            c.drawString(50, height - 50, "Excel Verileri")
            
            # Verileri yazdır
            y = height - 100
            c.setFont("Helvetica", 12)
            for index, row in df.iterrows():
                for col in df.columns:
                    c.drawString(50, y, f"{col}: {row[col]}")
                    y -= 20
                y -= 20
                if y < 50:
                    c.showPage()
                    y = height - 50
            
            c.save()
            return True
        except Exception as e:
            print(f"Excel to PDF dönüşümünde hata: {str(e)}")
            return False

    @staticmethod
    def word_to_pdf(input_path, output_path):
        try:
            # Word dosyasını oku
            doc = Document(input_path)
            
            # PDF oluştur
            c = canvas.Canvas(output_path, pagesize=letter)
            width, height = letter
            
            y = height - 50
            c.setFont("Helvetica", 12)
            
            for para in doc.paragraphs:
                c.drawString(50, y, para.text)
                y -= 20
                if y < 50:
                    c.showPage()
                    y = height - 50
            
            c.save()
            return True
        except Exception as e:
            print(f"Word to PDF dönüşümünde hata: {str(e)}")
            return False

    @staticmethod
    def pdf_to_word(input_path, output_path):
        try:
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()
            return True
        except Exception as e:
            print(f"PDF to Word dönüşümünde hata: {str(e)}")
            return False

    @staticmethod
    def image_to_pdf(input_path, output_path):
        try:
            image = Image.open(input_path)
            image.save(output_path, "PDF", resolution=100.0)
            return True
        except Exception as e:
            print(f"Görsel to PDF dönüşümünde hata: {str(e)}")
            return False 