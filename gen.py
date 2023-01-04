from docxtpl import DocxTemplate
import PySimpleGUI as sg
from datetime import datetime

doc = DocxTemplate('template.docx')

layout = [
    [sg.Text("Project Name"), sg.Input(key="project_name")],
    [sg.Text("Address"), sg.Input(key="address")],
    [sg.Text("Speed Limit"), sg.Input(key="speed_limit")],
    [sg.Text("Tax Map"), sg.Input(key="tax_map")],
    [sg.Text("Parcel"), sg.Input(key="parcel")],
    [sg.Text("SDP"), sg.Input(key="sdp")],
    [sg.Text("Author"), sg.Input(key="author")],
    [sg.Text("P.E. Full Name"), sg.Input(key="pe_full_name")],
    [sg.Text("P.E. License Number"), sg.Input(key="pe_license_no")],
    [sg.Text("P.E. License Expiry Date"), sg.Input(key="pe_expiration")],
    [sg.Text("Date"), sg.Input(key="date", size=(20,1)), sg.CalendarButton("Select", close_when_date_chosen=True, target="date")],
    [sg.Text("Start time"), sg.Input(key="start_time")],
    [sg.Text("End time"), sg.Input(key="end_time")],
    [sg.Text("Weather"), sg.Input(key="weather")],
    [sg.Text("Temperature (F)"), sg.Input(key="temp")],
    [sg.Button("Generate Study"), sg.Exit()]
]
window = sg.Window("Free Flow Study Generator", layout, element_justification='right')



while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Generate Study':
        print(event, values)

        tempDate = datetime.strptime(values['date'], '%Y-%m-%d %H:%M:%S')
        values['date'] = tempDate.strftime('%Y-%m-%d')
    

        doc.render(values)
        doc.save('output.docx')
        sg.popup("Free flow study generated successfully.")


