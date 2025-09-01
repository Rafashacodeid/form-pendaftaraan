import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkGreen4')

EXCEL_FILE = 'Pendaftaran.xlsx'

# Bikin file Excel kosong kalau belum ada
try:
    df = pd.read_excel(EXCEL_FILE)
except FileNotFoundError:
    df = pd.DataFrame(columns=['Nama','Tlp','Alamat','Tgl Lahir','Jekel','Belajar','Menonton','Musik'])
    df.to_excel(EXCEL_FILE, index=False)

layout = [
    [sg.Text('Masukan Data Kamu: ')],
    [sg.Text('Nama', size=(15,1)), sg.InputText(key='Nama')],
    [sg.Text('No Telp', size=(15,1)), sg.InputText(key='Tlp')],
    [sg.Text('Alamat', size=(15,1)), sg.Multiline(key='Alamat')],
    [sg.Text('Tgl Lahir', size=(15,1)), sg.InputText(key='Tgl Lahir'),
     sg.CalendarButton('Kalender', target='Tgl Lahir', format=('%d-%M-%Y'))],
    [sg.Text('Jenis Kelamin', size=(15,1)), sg.Combo(['pria','wanita'], key='Jekel')],
    [sg.Text('Hobi', size=(15,1)), sg.Checkbox('Belajar', key='Belajar'),
     sg.Checkbox('Menonton', key='Menonton'),
     sg.Checkbox('Musik', key='Musik')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Form Pendaftaran', layout)

def clear_input(values):
    for key in values:
        window[key]('')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input(values)
    if event == 'Submit':
        new_df = pd.DataFrame([values])
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Berhasil Disimpan')
        clear_input(values)

window.close()
