import xlrd#imports package for reading excel files
file_location = "D:/Simon/School/Year 12/Geography/Fieldwork/python/Excel Files/Fox Hill track test.xlsx"
workbook = xlrd.open_workbook(file_location)#opens excel file
sheet = workbook.sheet_by_index(0)#selects sheet in excel file

long = []
for row in range(2, sheet.nrows):#creates variable "row" and runs code from row 3 to the last row
    long.append(sheet.cell_value(row, 7))#adds the read value to the end of the array "long"

lat = []
for row in range(2, sheet.nrows):
    lat.append(sheet.cell_value(row, 6))

alt = []
for row in range(2, sheet.nrows):
    alt.append(sheet.cell_value(row, 8))

#print (long)
#print (lat)

kml = "<?xml version='1.0' encoding='UTF-8'?><kml xmlns='http://www.opengis.net/kml/2.2'><Document><Style id='linestyleGreen'><LineStyle><color>7f00ff00</color><width>4</width><gx:labelVisibility>1</gx:labelVisibility></LineStyle></Style>"
kml = kml + "<Placemark><name>Fox Hill Test</name><description>Fox Hill Track</description><styleUrl>#linestyleGreen</styleUrl><LineString><coordinates>"

file = open("test.kml","w")
for n in range(0, len(long)):
    kml = kml + "\n" + str(long[n]) + "," + str(lat[n]) + "," + str(alt[n])
kml = kml + "</coordinates></LineString></Placemark></Document></kml>"
    
file.write(kml)
file.close()
