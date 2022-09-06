#--------------------------------------------------------------------------------------------------------------------------------------------

import os

import pdf

from pathlib import Path
from turtle import home
import pandas as pd
import win32com.client as win32  # pip install pywin32

from docx2pdf import convert

import PySimpleGUI as sg
from docxtpl import DocxTemplate

import PySimpleGUI as sg
import csv, os

#--------------------------------------------------------------------------------------------------------------------------------------------

working_directory = os.getcwd()


cwd = os.getcwd()  # Get the current working directory (cwd)
files = os.listdir(cwd)  # Get all the files in that directory
working_directory = os.getcwd()

home = Path.home()
document_path = Path(cwd, "Arbeitsbericht.docx")
doc = DocxTemplate(document_path)

#--------------------------------------------------------------------------------------------------------------------------------------------

def clear_input():
    for key in values:
        window[key]('')
    return None

def save_as_word():
    ''' Save new file or save existing file with another name '''
    file_name: str = sg.popup_get_file('Save As', save_as=True, no_window=True,
                                           file_types=(('Word-Dokument', f"Arbeitsbericht_KW_{values['KW']}_{values['NAME']}.docx"), ('ALL Files', '*.*'),),
                                           modal=True, default_path=f"Arbeitsbericht_KW_{values['KW']}_{values['NAME']}.docx")
    return file_name

#--------------------------------------------------------------------------------------------------------------------------------------------

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
 
    [sg.Button("Convert DOCX to PDF", font=("Helvetica", 15))],

    [sg.Button("Speichern", font=("Helvetica", 15)), sg.Exit("Löschen", font=("Helvetica", 15)), sg.Exit("Beenden", font=("Helvetica", 15))],

    
    ]


window = sg.Window("Wochenberichte Generator APP", layout, element_justification="right", finalize=True)

#--------------------------------------------------------------------------------------------------------------------------------------------

while True:
    event, values = window.read()

    if event in (None, 'Beenden'):
        break

    if event == 'Löschen':
        clear_input()

    elif event =='Convert DOCX to PDF':
        pdf.pdf_create()
        

    elif event =='Speichern':
        doc.render(values)
        FILE_NAME = save_as_word()
        doc.save(FILE_NAME)

#--------------------------------------------------------------------------------------------------------------------------------------------

window.close()