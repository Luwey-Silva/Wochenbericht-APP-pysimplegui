
import PySimpleGUI as sg
import os
from docx2pdf import convert

working_directory = os.getcwd()


def pdf_create():
    layout_pdf = [  
            [sg.Text("Choose a WORD(.docx) file:")],
            [sg.InputText(key="-FILE_PATH-"), 
            sg.FileBrowse(initial_folder=working_directory, file_types=[("Word Dokument", "*.docx")])],
            [sg.Button('Submit'), sg.Exit()]
    ]

    pdf_window = sg.Window("PDF Converter", layout_pdf, modal=True)

    while True:
        event, values = pdf_window.read()
        if event in (None, 'Exit'):
            break
        elif event == "Submit":
            pdf_address = values["-FILE_PATH-"]
            convert(pdf_address)
            sg.popup("File saved", f"File has been saved here: {pdf_address}")

    pdf_window.close()
