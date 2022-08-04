import datetime
from pathlib import Path
import pandas as pd

import PySimpleGUI as sg
from docxtpl import DocxTemplate
import win32com.client as win32  # pip install pywin32
import os, sys  # Standard Python Libraries

document_path = Path(__file__).parent / "Arbeitsbericht.docx"
doc = DocxTemplate(document_path)

#today = datetime.datetime.today()
#today_in_one_week = today + datetime.timedelta(days=7)


EXCEL_FILE = 'Arbeitsbericht.xlsx'
df = pd.read_excel(EXCEL_FILE)


layout = [

    [sg.Text("Dein Name:", font=("Helvetica", 15)), sg.Input(key="NAME"), ],
    [sg.Text("Dein Nachname:", font=("Helvetica", 15)), sg.Input(key="NACHNAME")],
    [sg.Text("Kalenderwoche:", font=("Helvetica", 15)), sg.Input(key="KW")],
    [sg.Text("Gib die Stunden für Montag ein:", font=("Helvetica", 15)), sg.Input(key="USER_STUNDEN_MONTAG")],
    [sg.Text("Was würde am Montag gemacht?", font=("Helvetica", 15)), sg.Multiline(key="USER_BESCHREIBUNG_MONTAG", size=(50, 3))],    
    [sg.Text("Gib die Stunden für Dienstag ein:", font=("Helvetica", 15)), sg.Input(key="USER_STUNDEN_DIENSTAG")],
    [sg.Text("Was würde am Dienstag gemacht?", font=("Helvetica", 15)), sg.Multiline(key="USER_BESCHREIBUNG_DIENSTAG", size=(50, 3))],
    [sg.Text("Gib die Stunden für Mittwoch ein:", font=("Helvetica", 15)), sg.Input(key="USER_STUNDEN_MITTWOCH")],
    [sg.Text("Was würde am Mittwoch gemacht?", font=("Helvetica", 15)), sg.Multiline(key="USER_BESCHREIBUNG_MITTWOCH", size=(50, 3))],
    [sg.Text("Gib die Stunden für Donnerstag ein:", font=("Helvetica", 15)), sg.Input(key="USER_STUNDEN_DONNERSTAG")],
    [sg.Text("Was würde am Donnerstag gemacht?", font=("Helvetica", 15)), sg.Multiline(key="USER_BESCHREIBUNG_DONNERSTAG", size=(50, 3))],
    [sg.Text("Gib die Stunden für Freitag ein:", font=("Helvetica", 15)), sg.Input(key="USER_STUNDEN_FREITAG")],
    [sg.Text("Was würde am Freitag gemacht?", font=("Helvetica", 15)), sg.Multiline(key="USER_BESCHREIBUNG_FREITAG", size=(50, 3))],
    

    #[sg.Text("Client name:"), sg.Input(key="CLIENT", do_not_clear=False)],
    #[sg.Text("Vendor name:"), sg.Input(key="VENDOR", do_not_clear=False)],
    #[sg.Text("Amount:"), sg.Input(key="AMOUNT", do_not_clear=False)],
    #[sg.Text("Description1:"), sg.Input(key="LINE1", do_not_clear=False)],
    #[sg.Text("Description2:"), sg.Input(key="LINE2", do_not_clear=False)],
    
    [sg.Button("Speichern(Word)", font=("Helvetica", 15)), sg.Button("Speichern(Excel)", font=("Helvetica", 15)), sg.Button("Speichern(PDF)", font=("Helvetica", 15))],
    
    [sg.Exit("Löschen", font=("Helvetica", 15)), sg.Exit("Beenden", font=("Helvetica", 15))],

    ]

window = sg.Window("Wochenberichte Generator APP", layout, element_justification="right")

#window = sg.Window("Wochenberichte Generator APP", layout, element_justification="right", size=(700, 550))

#window = sg.Window("Wochenberichte Generator APP", layout, element_justification="right")


def clear_input():
    for key in values:
        window[key]('')
    return None

def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", r".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Beenden":
        break

    if event == 'Löschen':
        clear_input()
        
    if event == "Speichern(Word)":
        # Add calculated fields to our dict
        #values["NONREFUNDABLE"] = round(float(values["AMOUNT"]) * 0.2, 2)
        #values["TODAY"] = today.strftime("%Y-%m-%d")
        #values["TODAY_IN_ONE_WEEK"] = today_in_one_week.strftime("%Y-%m-%d")

        # Render the template, save new word document & inform user
        doc.render(values)
        output_path = Path(__file__).parent / f"Arbeitsbericht_KW_{values['KW']}_{values['NAME']}.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

    if event == 'Speichern(Excel)':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        output_path_ex = Path(__file__).parent / f"Arbeitsbericht_KW_{values['KW']}_{values['NAME']}.xlsx"
        df.to_excel(output_path_ex, index=False)
        sg.popup("File saved", f"File has been saved here: {output_path_ex}")

    if event == 'Speichern(PDF)':
        output_path = Path(__file__).parent / f"Arbeitsbericht_KW_{values['KW']}_{values['NAME']}.docx"
        path_to_word_document = os.path.join(os.getcwd(), output_path)
        convert_to_pdf(path_to_word_document)
        sg.popup("File saved", f"File has been saved here: {output_path}")

window.close()