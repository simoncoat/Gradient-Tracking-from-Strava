import xlrd                                         #import xlrd package
import math                                         #import math package
thresh1 = 0.1                                       #lower threshold for the slope
thresh2 = 0.3                                       #upper threshold for the slope
file_name = "Fox Hill track test.xlsx"              #file name of the excel file

file_location = "Excel Files/" + file_name          #location of excel file
workbook = xlrd.open_workbook(file_location)        #importing excel file
sheet = workbook.sheet_by_index(0)                  #select sheet in excel file

long = []
lat = []
slope = []
dlong = []
dlat = []
x = []
y = []
alt = []

for row in range(2, sheet.nrows):                   #from row 3 to the end of the file
    long.append(sheet.cell_value(row, 7))           #add the values from the cells in colum H into the list called long

for row in range(2, sheet.nrows):
    lat.append(sheet.cell_value(row, 6))            #add the values from the cells in colum G into the list called long
#print(long)



for i in range(0, len(long)-1):                     #from poiston 0 to 1 minus the length of list 'long'
    dlong.append(long[i+1]-long[i])                 #finding the difference in longitude

for i in range(0, len(lat)-1):
    dlat.append(lat[i+1]-lat[i])

for i in range(0, len(dlong)):
    x.append(dlong[i]*60*1852*math.cos(lat[i]))     #converting longitude difference to metres

for i in range(0, len(dlat)):
    y.append(dlat[i]*60*1852)                       #converting the latitude difference to metres



for row in range(2, sheet.nrows):
    alt.append(sheet.cell_value(row, 8))            #obtaining altitude from spreasheet



for i in range(0, len(long)-1):
	distance = math.sqrt((x[i]**2)+(y[i]**2))   #finding the distance between the GPS points
	dalt = alt[i+1]-alt[i]                      #finding the difference in altitude
	temp = math.fabs(dalt/distance)             #obtain the absolute value for slope
	slope.append(temp)                          #storing values for slope
	

#print (slope)
#print ("max="+str(max(slope)))

kml = "<?xml version='1.0' encoding='UTF-8'?>\n<kml xmlns='http://www.opengis.net/kml/2.2'>\n<Document><Style id='linestyleGreen'><LineStyle>\n<color>7f00ff00</color>\n<width>1</width>\n<gx:labelVisibility>1</gx:labelVisibility>\n</LineStyle>\n</Style><Style id='linestyleRed'>\n<LineStyle>\n<color>7f0000ff</color>\n<width>4</width>\n<gx:labelVisibility>1</gx:labelVisibility>\n</LineStyle>\n</Style><Style id='linestyleYellow'>\n<LineStyle>\n<color>7f00ffff</color>\n<width>4</width>\n<gx:labelVisibility>1</gx:labelVisibility>\n</LineStyle>\n</Style>"
boolAbove = False
    

#kml = kml + "</coordinates></LineString></Placemark>"
    

for i in range(0, len(slope)):
    if boolAbove:
        if slope[i] > thresh1 and slope[i] < thresh2:
            kml = kml + "\n" + str(long[i]) + "," + str(lat[i]) + "," + str(alt[i])
        else:
            kml = kml + "</coordinates>\n</LineString>\n</Placemark>"
            boolAbove = False
    else:
        if slope[i] > thresh1 and slope[i] < thresh2:
            boolAbove = True
            kml = kml + "\n<Placemark>\n<name>Fox Hill Test</name>\n<description>Fox Hill Track</description>\n<styleUrl>#linestyleYellow</styleUrl>\n<LineString>\n<coordinates>"
            kml = kml + "\n" + str(long[i]) + "," + str(lat[i]) + "," + str(alt[i])
            
for i in range(0, len(slope)):
    if boolAbove:
        if slope[i] > thresh2:
            kml = kml + "\n" + str(long[i]) + "," + str(lat[i]) + "," + str(alt[i])
        else:
            kml = kml + "</coordinates>\n</LineString>\n</Placemark>"
            boolAbove = False
    else:
        if slope[i] > thresh2:
            boolAbove = True
            kml = kml + "\n<Placemark>\n<name>Fox Hill Test</name>\n<description>Fox Hill Track</description>\n<styleUrl>#linestyleRed</styleUrl>\n<LineString>\n<coordinates>"
            kml = kml + "\n" + str(long[i]) + "," + str(lat[i]) + "," + str(alt[i])

kml = kml + "\n<Placemark>\n<name>Fox Hill Test</name>\n<description>Fox Hill Track</description>\n<styleUrl>#linestyleGreen</styleUrl>\n<LineString>\n<coordinates>"
for n in range(0, len(lat)):
    kml = kml + "\n" + str(long[n]) + "," + str(lat[n]) + "," + str(alt[n])
kml = kml + "</coordinates>\n</LineString>\n</Placemark>"

kml = kml + "</Document></kml>"
file = open("test2.kml","w")
file.write(kml)
file.close()
