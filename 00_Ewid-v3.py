#import PyPDF2
##import camelot
import openpyxl
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from pdfminer.high_level import extract_text, extract_pages
from pdfminer.layout import LAParams, LTTextBoxHorizontal, LTTextLine, LTTextBox
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from tabulate import tabulate

# Otwórz plik PDF
nazwa_pliku = input('Podaj nazwę pliku PDF w katalogu (np. Wypisy.pdf): ')
if len(nazwa_pliku) < 1 : nazwa_pliku = 'Wypisy.pdf'

# Odczytaj plik PDF przez PyPDF2
#pdf_reader = PyPDF2.Pdf{"path":"/tauri/F/DANE/01_Python/04_Ewid/00_Ewid-v2.py"}Reader(nazwa_pliku) 

# Utwórz nowy skoroszyt Excel
skoroszyt = openpyxl.Workbook()

# Utwórz arkusze Excel
sheet_zestaw = skoroszyt.active
sheet_zestaw.title = "WYPISY"

# Dodaj nagłówki do arkuszy.
sheet_zestaw.append(['L.p.', 'Data',  'Województwo', 'Powiat', 'Jednostka ewidencyjna', 'Obręb', 
                     'NUMER DZIAŁKI', 'PODMIOT EWIDENCYJNY', 'POW.DZIAŁKI', 'DOKUMENT'])

# Zmienna do wyśrodkuj i pogrub
pogrub = Font(bold=True)
srodek = Alignment(horizontal='center')

fp = open(nazwa_pliku, 'rb')
parser = PDFParser(fp)
document = PDFDocument(parser)

if not document.is_extractable:
    raise PDFTextExtractionNotAllowed

rsrcmgr = PDFResourceManager()

laparams = LAParams()
device = PDFPageAggregator(rsrcmgr, laparams=laparams)
interpreter = PDFPageInterpreter(rsrcmgr, device)

# Definicja do szukania nazw w pliku
def szukaj_nazw(miejsce_szukania, szukana_nazwa):
    szukaj_match = re.search(rf'.*?{szukana_nazwa}\s(.*?)\n', miejsce_szukania)
    if szukaj_match:
        wyszukane = szukaj_match.group(1)
    else:
        wyszukane = "Nie znaleziono"
    return wyszukane

# Pętla do iterowania po stronach, tabelach i wierszach
for page in PDFPage.create_pages(document):
    interpreter.process_page(page)
    layout = device.get_result()
    

    # Analiza układu strony i identyfikacja tabel
    tabela = []
    wiersz = []
    poprzedni_y = 0
    for element in layout:
        if isinstance(element, LTTextBox):
            if abs(element.y0 - poprzedni_y) > 15:  # Nowy wiersz, jeśli różnica w Y jest większa niż 10
                if wiersz:
                    tabela.append(wiersz)
                wiersz = []
            wiersz.append(element.get_text().replace('\n', ' ').strip().replace('   ', '  ').replace('  ', ' '))
            poprzedni_y = element.y0
    if wiersz:
        tabela.append(wiersz)
    for lista_1 in tabela:
        print(lista_1)
        #print(tabela)
    #print(tabulate(tabela, headers="firstrow", tablefmt="grid"))    

    # Wyświetlenie tabeli
    #for wiersz in tabela:
        #print(wiersz)
        #print(type(wiersz))
        # Wyświetlenie tabeli
        

'''
lp = 0

# Odczytaj plik przez PDFminer
for page_layout in extract_pages(nazwa_pliku):
    for element in page_layout:
        if isinstance(element, LTTextLine):
            print(elemen.get_text())
            #print(type(element))
            
            page_text = ""
            for text_line in element:
                if isinstance(text_line, LTTextLine):
                    page_text += text_line.get_text()
                    
                    # Wyszukaj datę
                    data = szukaj_nazw(page_text, "z dnia")

                    # Wyszukaj województwo
                    woj = szukaj_nazw(page_text, "Województwo")

                    # Wyszukaj powiat
                    powiat = szukaj_nazw(page_text, "Powiat")

                    # Wyszukaj jednostka ewidencyjna
                    je = szukaj_nazw(page_text, "ewidencyjna")

                    # Wyszukaj obręb
                    obreb = szukaj_nazw(page_text, "bręb")

                    # Dodaj dane do arkusza Excel
                    sheet_zestaw.append([lp, data, woj, powiat, je, obreb, 'NUMER DZIAŁKI', 'PODMIOT EWIDENCYJNY', 'POW.DZIAŁKI', 'DOKUMENT'])

# Wysrodkuj dane w komórkach w arkuszu
licznik = 0
for row in sheet_zestaw.iter_rows():
    licznik = licznik + 1
    if licznik == 1 :
        for cell in row:
            cell.alignment = srodek
            cell.font = pogrub
    else:        
        for cell in row:
            cell.alignment = srodek
            
# Dopasuj szerokość kolumn do zawartości
for column_cells in sheet_zestaw.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    sheet_zestaw.column_dimensions[column_cells[0].column_letter].width = length

# Przetwórz dane z nazwy pliku wejściwego PDF, aby je wykorzystać do pliku wyjściowego XLSX
nazwa_pliku_bez_rozszerzenia = nazwa_pliku.replace(".pdf", "")
nazwa_pliku_xlsx = "WYPISY_" + nazwa_pliku_bez_rozszerzenia + ".xlsx"

# Zapisz skoroszyt Excel
skoroszyt.save(nazwa_pliku_xlsx)

print(f"Dane zapisano do pliku: {nazwa_pliku_xlsx}")
'''
fp.close()
