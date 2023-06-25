# Input: Reporter-Abfrage "Gesamtliste Stimmbildung (Vorlage für Übersichtsplan)"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
import xlsxwriter
import datetime
import re

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d")
felder = ["Datum", "Zeit", "Vorname", "Name", "Alter", "Chor", "Stimme", "Instrumente", "Zeit aktiv", "Schule", "Lieder", "Stibi", "Mot.", "Int.", "Sti.", "h.S.", "Ver.", "Aktuelles Lied"]
breite = [10, 8, 15, 15, 4, 12, 10, 12, 10, 16, 30, 18, 5, 5, 5, 5, 5, 25]

print("Bitte Reporter-Abfrage 'Planung Vorsingen (Gesamtliste)'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

if not data.startswith("lfd. Nr.\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

with io.StringIO(data) as infile:
    daten = list(csv.DictReader(infile, delimiter="\t"))

spatzen = []
neuer_spatz = False

for eintrag in daten:
    if eintrag["lfd. Nr."]:
        if neuer_spatz:
            spatz["Lieder"] = spatz["Lieder"].strip()
            spatz["Instrumente"] = spatz["Instrumente"].strip()
            spatzen.append(spatz)
        neuer_spatz = True
        spatz = eintrag.copy()
        spatz["Instrumente"] = ""
        spatz["Lieder"] = ""
        del spatz["Schule/Lied"]
        del spatz["Stimme/Instr."]
        del spatz["Wert"]
    if eintrag["Stimme/Instr."] in ("Alt", "Sopran 1", "Sopran 2"):
        spatz["Stimme"] = eintrag["Stimme/Instr."]
    elif eintrag["Stimme/Instr."]:
        spatz["Instrumente"] += eintrag["Stimme/Instr."] + "\n"
    if eintrag["Schule/Lied"] == "Schule":
        spatz["Schule"] = eintrag["Wert"]
    elif eintrag["Schule/Lied"]:
        if m := re.search(r"\d+", eintrag["Schule/Lied"]):
            spatz["Lieder"] += m.group(0) + ": "
        spatz["Lieder"] += eintrag["Wert"] + "\n"

spatz["Lieder"] = spatz["Lieder"].strip()
spatz["Instrumente"] = spatz["Instrumente"].strip()
spatzen.append(spatz) # Letzten Spatzen auch noch in die Liste

spatzen.sort(key=lambda l:l["Chor"], reverse=True)

with open("Vorsingen.csv", "w", newline="", encoding="cp1252") as outfile:
    writer = csv.DictWriter(outfile, fieldnames=felder, extrasaction="ignore", delimiter=";")
    writer.writeheader()
    for spatz in spatzen:
        writer.writerow(spatz)

with open("Vorsingen.csv", newline="", encoding="cp1252") as infile:
    reader = csv.reader(infile, delimiter=";")
    data = list(reader)

anzahl_zeilen = len(data)
anzahl_spalten = len(data[0])

with xlsxwriter.Workbook(f"Vorsingen_{heute}.xlsx") as outxlsx:
    excel = outxlsx.add_worksheet("Plan Vorsingen")
    excel.set_paper(8) # DIN A3
    excel.set_landscape()
    excel.set_margins(0.3, 0.3, 0.6, 0.6) # Seitenränder in Zoll
    excel.fit_to_pages(1,0) # An Seitenbreite anpassen
    excel.set_header("&CPlan fürs Vorsingen – Ulmer Spatzen Chor – Stand: &D")
    excel.set_default_row(45) # Zeilenhöhe: Drei Textzeilen
    for spalte, eintrag in enumerate(breite):
        excel.set_column(spalte, spalte, eintrag)
    excel.repeat_rows(0)    
    standard = outxlsx.add_format({"text_wrap": True, "valign": "top", "align": "left"})
    spalten = [{"header": item, "format": standard} for item in data[0]]
    table_style = {"data": data[1:], 
                   "style": "Table Style Medium 15",
                   "name": "Vorsingen",
                   "columns": spalten}
    datum = outxlsx.add_format({'num_format': 'dd.mm.yy', "valign": "top", "align": "left"})
    uhrzeit = outxlsx.add_format({'num_format': 'hh:mm', "valign": "top", "align": "left"})
    table_style["columns"][0]["format"] = datum
    table_style["columns"][1]["format"] = uhrzeit
    excel.add_table(0, 0, anzahl_zeilen-1, anzahl_spalten-1, table_style)
    
print(f"Fertig! Die Datei Vorsingen_{heute}.xlsx wurde im aktuellen Ordner abgelegt.")
