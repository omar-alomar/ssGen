from docxtpl import DocxTemplate
import PySimpleGUI as sg
from datetime import datetime
import usaddress as addr

doc = DocxTemplate('template.docx')
sg.theme('LightGrey1')
layout = [
    [sg.Text("Project Name"), sg.Input(key="project_name")],
    [sg.Text("Full Address"), sg.Input(key="address")],
    [sg.Text("Tax Map"), sg.Input(key="tax_map", size=(5,1)),
     sg.Text("Parcel"), sg.Input(key="parcel", size=(5,1)), 
     sg.Text("SDP"), sg.Input(key="sdp", size=(7,1))],
    [sg.Text("Author"), sg.Input(key="author")],
    [sg.Text("P.E. Full Name"), sg.Input(key="pe_full_name")],
    [sg.Text("P.E. License Number"), sg.Input(key="pe_license_no", size=(9,1)),
     sg.Text("Expiry Date"), sg.Input(key="pe_expiration", size=(9,1)), sg.CalendarButton("Select", close_when_date_chosen=True, target="pe_expiration", no_titlebar=False)],
     [sg.Text("")],
    [sg.Text("Start time"), sg.Input(key="start_time", size=(8,1)),
     sg.Text("End time"), sg.Input(key="end_time", size=(8,1)),
     sg.Text("Date"), sg.Input(key="date", size=(9,1)), sg.CalendarButton("Select", close_when_date_chosen=True, target="date", no_titlebar=False)],
    [sg.Text("Speed Limit"), sg.Input(key="speed_limit", size=(5,1)),
     sg.Text("Weather"), sg.Input(key="weather", size=(20,1)),
     sg.Text("Temperature (F)"), sg.Input(key="temp", size=(5,1))],
     [sg.Text("")],
    [sg.Text("Car Direction 1"),
     sg.Combo(['North bound','Northeast bound','East bound','Southeast bound','South bound','Southwest bound','West bound','Northwest bound'], key="direction1"),
     sg.Text("Car Direction 2"),
     sg.Combo(['North bound','Northeast bound','East bound','Southeast bound','South bound','Southwest bound','West bound','Northwest bound'], key="direction2"),
     sg.Text("")],
    [sg.Multiline(key='dir1_speeds', size=(30,10)),
     sg.Multiline(key='dir2_speeds', size=(30,10))],
    [sg.Text("Please enter speeds as comma separated values with no spaces                  ")],
    [sg.Text("")],
    [sg.Button("Generate Study"), sg.Exit()]
]
window = sg.Window("Free Flow Study Generator", layout, element_justification='right')



while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Generate Study':
        # print(event, values)

        tempDate = datetime.strptime(values['date'], '%Y-%m-%d %H:%M:%S')
        values['date'] = tempDate.strftime('%Y/%m/%d')
        values['date_long'] = tempDate.strftime('%A %B %d, %Y')

        tempDate = datetime.strptime(values['pe_expiration'], '%Y-%m-%d %H:%M:%S')
        values['pe_expiration'] = tempDate.strftime('%Y-%m-%d')

        values['month_date'] = tempDate.strftime('%B %Y')

        parsedAddressTemp = addr.tag(values['address'])
        parsedAddress = parsedAddressTemp[0]
        values['road_name'] = parsedAddress['StreetName'] + ' ' + parsedAddress['StreetNamePostType']

        values['weather'] = values['weather'].lower()

        dir1Speeds = values['dir1_speeds'].split(',')        
        dir1Speeds = [int(i) for i in dir1Speeds]
        values['dir1_speeds'] = dir1Speeds
        
        dir2Speeds = values['dir2_speeds'].split(',')
        dir2Speeds = [int(i) for i in dir2Speeds]
        values['dir2_speeds'] = dir2Speeds

        dir1Vehs = len(dir1Speeds)
        dir2Vehs = len(dir2Speeds)
        totalVehs = dir1Vehs + dir2Vehs
        values['total_vehs'] = str(totalVehs)

        dir1SpeedsSorted = sorted(dir1Speeds)
        dir2SpeedsSorted = sorted(dir2Speeds)

        # creating dicts to store stuff
        dir1SpeedsDict = {}
        dir1Percents = {}
        dir1CumPercents = {}
        for i in range(len(dir1SpeedsSorted)):
            if dir1SpeedsSorted[i] != dir1SpeedsSorted[i-1]:
                dir1SpeedsDict[dir1SpeedsSorted[i]] = 0
                dir1Percents[dir1SpeedsSorted[i]] = 0
                dir1CumPercents[dir1SpeedsSorted[i]] = 0

        for i in dir1SpeedsSorted:
            count = 0
            for j in dir1SpeedsSorted:
                if i == j:
                    count += 1
            dir1SpeedsDict[i] = count

        cumPercent = 0.0
        for speed in dir1Percents.items():
            dir1Percents[speed[0]] = dir1SpeedsDict[speed[0]] / dir1Vehs
            cumPercent += dir1Percents[speed[0]]
            dir1CumPercents[speed[0]] = cumPercent

        dir2SpeedsDict = {}
        dir2Percents = {}
        dir2CumPercents = {}
        for i in range(len(dir2SpeedsSorted)):
            if dir2SpeedsSorted[i] != dir2SpeedsSorted[i-1]:
                dir2SpeedsDict[dir2SpeedsSorted[i]] = 0
                dir2Percents[dir2SpeedsSorted[i]] = 0
                dir2CumPercents[dir2SpeedsSorted[i]] = 0

        for i in dir2SpeedsSorted:
            count = 0
            for j in dir2SpeedsSorted:
                if i == j:
                    count += 1
            dir2SpeedsDict[i] = count

        cumPercent = 0.0
        for speed in dir2Percents.items():
            dir2Percents[speed[0]] = dir2SpeedsDict[speed[0]] / dir2Vehs
            cumPercent += dir2Percents[speed[0]]
            dir2CumPercents[speed[0]] = cumPercent
        


        print(dir2Percents)
        print(dir2CumPercents)



        doc.render(values)
        doc.save('output.docx')
        sg.popup("Free flow study generated successfully.")


