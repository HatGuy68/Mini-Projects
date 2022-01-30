import openpyxl

path = "./sample.xlsx"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

row = sheet_obj.max_row
column = sheet_obj.max_column

lastSerialNumber = row
running = True

while running:
    running = False
    PatientID = sheet_obj['B'+str(lastSerialNumber)].value + 1
    firstName = input('Enter patient first name: ')
    lastName = input('Enter patient last name: ')
    condition = input('Enter patient health problem: ')
    medication = str(input('Enter all medication: ').split(' '))
    activities = str(input('Enter all activities: ').split(' '))
    
    sheet_obj['A'+str(lastSerialNumber+1)] = lastSerialNumber
    sheet_obj['B'+str(lastSerialNumber+1)] = PatientID
    sheet_obj['C'+str(lastSerialNumber+1)] = firstName
    sheet_obj['D'+str(lastSerialNumber+1)] = lastName
    sheet_obj['E'+str(lastSerialNumber+1)] = condition
    sheet_obj['F'+str(lastSerialNumber+1)] = medication
    sheet_obj['G'+str(lastSerialNumber+1)] = activities
    
    wb_obj.save(path)


# Patient Name :
# Max
# Mustermann
# Gesundheitsproblem: Herzinsuffizienz              Health problem: heart failure
# Medikamente:                                      Medicines
# Blutverdünner Thomapyrin                          blood thinner thomapyrin
# Aspirin
# Tätigkeiten                                       activities
# Blutdruck                                         blood-pressure
# Pulsschlag                                        Pulse-rate
# Blutzucker spiegel                                blood-sugar-levels
# Blut abgenommen                                   Blood-drawn
