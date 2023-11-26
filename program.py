from datetime import datetime, timedelta
import xlwings as xw
import random
import tkinter as tk
from tkinter import filedialog
import os

# Excel-Zeit in Minuten umwandeln
def excel_zeit_in_minuten(excel_zeit):
    """Excel-Zeit in Minuten umwandeln."""
    gesamtminuten = round(excel_zeit * 24 * 60)
    return gesamtminuten

# Überprüfen, ob der Wert ein gültiges Datum im angegebenen Format ist
def ist_gueltiges_datum(zeitwert, datumsformat="%Y-%m-%d %H:%M:%S"):
    """Überprüfen, ob der Wert ein gültiges Datum im spezifizierten Format ist."""
    try:
        return datetime.strptime(zeitwert, datumsformat)
    except (ValueError, TypeError):
        return None

# Zufällige Zeit in einem Bereich generieren
def zufaellige_zeit_im_bereich(start_stunde, ende_stunde, inkrement, datum):
    """Eine zufällige Zeit zwischen Start- und Endstunde in Inkrementen von 'inkrement' Minuten generieren."""
    gesamtminuten = random.randrange(start_stunde * 60, ende_stunde * 60, inkrement)
    stunden = gesamtminuten // 60
    minuten = gesamtminuten % 60
    return datetime(datum.year, datum.month, datum.day, stunden, minuten)

# Excel-Datei verarbeiten
def verarbeite_excel(dateipfad):
    # Deutsche Monatsnamen und Wochentage für die Filterung definieren
    deutsche_monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", 
                       "Juli", "August", "September", "Oktober", "November", "Dezember"]

    # Workbook mit xlwings öffnen
    with xw.App(visible=False) as app:  
        arbeitsbuch = xw.Book(dateipfad)

        # Durch alle Blätter im Arbeitsbuch iterieren
        for blatt in arbeitsbuch.sheets:
            blatt_name = blatt.name

            # Blätter überspringen, die nach deutschen Monaten benannt sind
            if blatt_name in deutsche_monate:
                continue

            # Durch jede Zeile im Blatt iterieren
            for zeile in blatt.range('A1:P100').rows:  # Bereich nach Bedarf anpassen
                try:
                    geparstes_datum = ist_gueltiges_datum(str(zeile[0].value))
                    if geparstes_datum is None:
                        continue

                    if zeile[11].value is None or float(zeile[11].value) <= 0:
                        continue

                    # Erforderliche Arbeitsdauer in Minuten berechnen
                    erforderliche_minuten = excel_zeit_in_minuten(zeile[11].value)

                    # Zufällige Startzeit zwischen 14:00 und 18:00 generieren
                    start_zeit = zufaellige_zeit_im_bereich(14, 17, 5, geparstes_datum)

                    # Endzeit berechnen
                    ende_zeit = start_zeit + timedelta(minutes=erforderliche_minuten)

                    # Zeiten den Zellen zuweisen
                    zeile[3].value = str(start_zeit.time())
                    zeile[4].value = str(ende_zeit.time())

                except Exception as e:
                    print(f"Fehler bei der Verarbeitung der Zeile: {e}")
                    pass

        # Arbeitsbuch speichern
        arbeitsbuch.save()

# GUI für die Excel-Verarbeitung
def gui_excel_verarbeitung():
    fenster = tk.Tk()
    fenster.title("Excel Verarbeiter")

    def oeffne_dateidialog():
        dateipfad = filedialog.askopenfilename(title="Datei auswählen",
                                               filetypes=[("Excel-Dateien", "*.xlsx")])
        dateipfad_eingabe.delete(0, tk.END)
        dateipfad_eingabe.insert(0, dateipfad)

    def starte_verarbeitung():
        dateipfad = dateipfad_eingabe.get()
        if os.path.exists(dateipfad):
            verarbeite_excel(dateipfad)
            status_label.config(text="Verarbeitung abgeschlossen")
        else:
            status_label.config(text="Ungültiger Dateipfad")

    dateipfad_eingabe = tk.Entry(fenster, width=50)
    dateipfad_eingabe.pack(pady=10)

    durchsuchen_button = tk.Button(fenster, text="Durchsuchen", command=oeffne_dateidialog)
    durchsuchen_button.pack(pady=5)

    start_button = tk.Button(fenster, text="Verarbeitung starten", command=starte_verarbeitung)
    start_button.pack(pady=5)

    status_label = tk.Label(fenster, text="")
    status_label.pack(pady=10)

    fenster.mainloop()

# GUI-Funktion aufrufen
gui_excel_verarbeitung()
