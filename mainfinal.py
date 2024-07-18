import os
from pdf2image import convert_from_path
import pytesseract
import re
import csv

# Pfade zu Poppler und Tesseract
poppler_path = r'C:\Users\Ikram.Hjira\Downloads\Release-24.02.0-0\poppler-24.02.0\Library\bin'
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Ikram.Hjira\AppData\Local\Tesseract-OCR\tesseract.exe'
tessdata_dir_config = r'--tessdata-dir "C:\Users\Ikram.Hjira\AppData\Local\Tesseract-OCR\tessdata"'

# Muster für Bestell- und Auftragsnummern
bestellnummer_pattern = r'\b((?:58|59)\d{8})\b'
auftragsnummer_pattern = r'\b((?:3|4|7)\d{7})\b'

# Funktion zur Suche nach Bestell- und Auftragsnummern
def find_numbers(text):
    bestellnummer = re.findall(bestellnummer_pattern, text)
    auftragsnummer = re.findall(auftragsnummer_pattern, text)
    return bestellnummer, auftragsnummer

# Funktion zur Verarbeitung einer PDF-Datei und Schreiben in CSV
def process_pdf_and_write_to_csv(pdf_path, csv_writer):
    try:
        images = convert_from_path(pdf_path, poppler_path=poppler_path)
        text = ''
        for image in images:
            text += pytesseract.image_to_string(image, lang='deu', config=tessdata_dir_config)
        
        bestellnummer, auftragsnummer = find_numbers(text)

        # Überprüfen und Schreiben der Ergebnisse in die CSV-Datei
        if bestellnummer or auftragsnummer:
            print(f'Bestell- oder Auftragsnummer gefunden für: {pdf_path}')
            csv_writer.writerow([', '.join(bestellnummer), ', '.join(auftragsnummer)])
        else:
            print(f'Keine Bestell- oder Auftragsnummer gefunden für: {pdf_path}')

    except Exception as e:
        print(f'Error processing {pdf_path}: {e}')

# Funktion zum Durchsuchen eines Verzeichnisses nach PDF-Dateien und Schreiben in CSV
def search_directory_and_write_to_csv(directory, output_csv):
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(['Bestellnummer', 'Auftragsnummer'])

        found_any = False
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(root, file)
                    process_pdf_and_write_to_csv(pdf_path, csv_writer)
                    found_any = True

        if not found_any:
            print(f'Keine PDF-Dateien im Verzeichnis {directory} gefunden.')

# Hauptverzeichnis zum Suchen
main_directory = r'U:\project\Neuer Ordner'
output_csv = r'U:\project\Neuer Ordner\results.csv'  # Hier den gewünschten Dateinamen für die CSV-Datei angeben

# Suche im Hauptverzeichnis und Schreiben in CSV
search_directory_and_write_to_csv(main_directory, output_csv)
