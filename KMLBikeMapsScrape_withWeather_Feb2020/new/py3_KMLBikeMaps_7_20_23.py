##This script will be used for the city of Ottawa to webscrape data from www.BikeMaps.org for certain parts of the city or whatever location 
# they need information from.
# Conner Leverett, BikeMaps Intern, March 4th, 2016
# Updated Colin Ferster, BikeMaps Post Doc, May 14, 2019

# Some of this code was taken from here: http://chrisagocs.blogspot.ca/2013/01/how-to-parse-kml-file-and-find-centroid.html

try:
    # Import Modules
    import json
    import urllib
    from urllib.request import urlopen
    from xml.dom.minidom import parseString
    import os
    import sys

    sys.path.append(os.path.abspath(os.path.dirname(__file__)) + "/Modules")
    import xlsxwriter
    import time
except Exception as ex:
    print(ex)
    time.sleep(5)


def extractDataFromKML(passedInString):
    # Open the file
    kmlFile = open(passedInString, 'r')

    # This turns the kml file into a string
    k = (kmlFile.read())

    # This turns the string into a parsable object
    dom = parseString(k)

    # This finds the tag 'coordinates' which hold the polygon lines
    coordinates = dom.getElementsByTagName('coordinates')

    # Turns the data into a list of coordinates
    listOfCoordinates = (coordinates[0].firstChild.data.split(','))

    newListOfCoordinates = []

    # Create holder for list of float coordinates
    listOfFloatCoordinates = []

    # Turns the coordinates from a string to a float
    for coordinate in listOfCoordinates:
        foundSpace = False
        for x in range(len(coordinate)):
            if coordinate[x] == " ":
                newListOfCoordinates.append(coordinate[x + 1:])
                foundSpace = True
                break
        if foundSpace == False:
            newListOfCoordinates.append(coordinate)

    newListOfCoordinates.pop()

    for num in newListOfCoordinates:
        listOfFloatCoordinates.append(float(num))

    # Create Holder for final list of lat/long
    finalListOfCoordinates = []

    # Taking the lon/lat and entering them into a final list of coordinates which would be represented as [[-123.5,48.5],[-123.6,48.2]...]
    # Note that longitude comes before latitude
    for num in range(0, len(listOfFloatCoordinates), 2):
        littleList = [listOfFloatCoordinates[num], listOfFloatCoordinates[num + 1]]
        finalListOfCoordinates.append(littleList)

    return finalListOfCoordinates


####This code was found here: http://geospatialpython.com/2011/01/point-in-polygon.html and was not written by Conner Leverett
def point_in_poly(x, y, poly):
    # Determine if a point is inside a given polygon or not
    # Polygon is a list of (x,y) pairs. This function
    # returns True or False.  The algorithm is called
    # the "Ray Casting Method".

    n = len(poly)
    inside = False

    p1x, p1y = poly[0]
    for i in range(n + 1):
        p2x, p2y = poly[i % n]
        if y > min(p1y, p2y):
            if y <= max(p1y, p2y):
                if x <= max(p1x, p2x):
                    if p1y != p2y:
                        xints = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                    if p1x == p2x or x <= xints:
                        inside = not inside
        p1x, p1y = p2x, p2y

    return inside


def scrapeAndStore(polygon, bbx):
    # List of websites which need to be scraped
    listOfURLs = ["https://bikemaps.org/incidents-only.json", "https://bikemaps.org/hazards.json",
                  "https://bikemaps.org/thefts.json"]

    # Create lists to hold the data for each city
    IncidentsList = []
    HazardsList = []
    TheftsList = []

    # Lists of lists
    overallList = [IncidentsList, HazardsList, TheftsList]

    # Loop through all of the websites which need data extracted from
    for url in listOfURLs:
        # Open the website
        response = urlopen(url)
        # Grab the data from the website
        data = json.loads(response.read())
        # print(data)
        # Loop through points and determine if they fall within any city limits, if they do, append that point's info to the City list
        if url == "https://bikemaps.org/incidents.json?bbox={}".format(bbx):
            for dt in data:
                point = dt['incident']
                if point_in_poly(point['geometry']['coordinates'][0], point['geometry']['coordinates'][1],
                                 polygon) == True:
                    IncidentsList.append(point)
        else:
            for point in (data['features']):
                # Check to see if it's within area limits
                if point_in_poly(point['geometry']['coordinates'][0], point['geometry']['coordinates'][1],
                                 polygon) == True:
                    if url == "https://bikemaps.org/hazards.json":
                        HazardsList.append(point)
                        continue
                    if url == "https://bikemaps.org/hazards.json":
                        HazardsList.append(point)
                        continue
                    if url == "https://bikemaps.org/incidents-only.json":
                        IncidentsList.append(point)
                        continue

    return overallList


def createCSVSheets(csvFile, bold):
    # Add the incident sheet
    incident = csvFile.add_worksheet('Incident')
    # Add titles to the incident sheet
    incident.write(0, 0, 'i_type', bold)
    incident.write(0, 1, 'incident_with', bold)
    incident.write(0, 2, 'date', bold)
    incident.write(0, 3, 'p_type', bold)
    incident.write(0, 4, 'personal_involvement', bold)
    incident.write(0, 5, 'details', bold)
    incident.write(0, 6, 'impact', bold)
    incident.write(0, 7, 'injury', bold)
    incident.write(0, 8, 'trip_purpose', bold)
    incident.write(0, 9, 'regular_cyclist', bold)
    incident.write(0, 10, 'helmet', bold)
    incident.write(0, 11, 'road_conditions', bold)
    incident.write(0, 12, 'sightlines', bold)
    incident.write(0, 13, 'cars_on_roadside', bold)
    incident.write(0, 14, 'bike_lights', bold)
    incident.write(0, 15, 'terrain', bold)
    incident.write(0, 16, 'aggressive', bold)
    incident.write(0, 17, 'intersection', bold)
    incident.write(0, 18, 'witness_vehicle', bold)
    incident.write(0, 19, 'bicycle_type', bold)
    incident.write(0, 20, 'ebike', bold)
    incident.write(0, 21, 'ebike_class', bold)
    incident.write(0, 22, 'ebike_speed', bold)
    incident.write(0, 23, 'direction', bold)
    incident.write(0, 24, 'turning', bold)
    incident.write(0, 25, 'age', bold)
    incident.write(0, 26, 'gender', bold)
    incident.write(0, 27, 'birthmonth', bold)

    # incident.write(0, 25, 'weather_summary', bold)
    # incident.write(0, 26, 'weather_sunrise_time', bold)
    # incident.write(0, 27, 'weather_sunset_time', bold)
    # incident.write(0, 28, 'weather_dawn', bold)
    # incident.write(0, 29, 'weather_dusk', bold)
    # incident.write(0, 30, 'weather_precip_intensity', bold)
    # incident.write(0, 31, 'weather_precip_probability', bold)
    # incident.write(0, 32, 'weather_precip_type', bold)
    # incident.write(0, 33, 'weather_temperature', bold)
    # incident.write(0, 34, 'weather_black_ice_risk', bold)
    # incident.write(0, 35, 'weather_wind_speed', bold)
    # incident.write(0, 36, 'weather_wind_bearing', bold)
    # incident.write(0, 37, 'weather_visibility_km', bold)
    incident.write(0, 28, 'pk', bold)
    incident.write(0, 29, 'longitude', bold)
    incident.write(0, 30, 'latitude', bold)

    # Add the hazards sheet
    hazards = csvFile.add_worksheet('Hazards')
    # Add title to the hazards sheet
    hazards.write(0, 0, 'i_type', bold)
    hazards.write(0, 1, 'date', bold)
    hazards.write(0, 2, 'p_type', bold)
    hazards.write(0, 3, 'details', bold)
    hazards.write(0, 4, 'age', bold)
    hazards.write(0, 5, 'birthmonth', bold)
    hazards.write(0, 7, 'pk', bold)
    hazards.write(0, 8, 'longitude', bold)
    hazards.write(0, 9, 'latitude', bold)
    hazards.write(0, 10, 'gender', bold)

    # Add the theft sheet
    theft = csvFile.add_worksheet('Theft')
    # Add titles to thefts
    theft.write(0, 0, 'i_type', bold)
    theft.write(0, 1, 'date', bold)
    theft.write(0, 2, 'p_type', bold)
    theft.write(0, 3, 'details', bold)
    theft.write(0, 4, 'how_locked', bold)
    theft.write(0, 5, 'lock', bold)
    theft.write(0, 6, 'locked_to', bold)
    theft.write(0, 7, 'lighting', bold)
    theft.write(0, 8, 'traffic', bold)
    theft.write(0, 9, 'police_report', bold)
    theft.write(0, 10, 'police_report_num', bold)
    theft.write(0, 11, 'insurance_claim', bold)
    theft.write(0, 12, 'insurance_claim_num', bold)
    theft.write(0, 13, 'regular_cyclist', bold)
    theft.write(0, 14, 'pk', bold)
    theft.write(0, 15, 'longitude', bold)
    theft.write(0, 16, 'latitude', bold)

    listOfSheets = [incident, hazards, theft]
    return listOfSheets


def writeToIncident(incident, point, row):
    incident.write(row, 0, point['properties']['i_type'])
    incident.write(row, 1, point['properties']['incident_with'])
    incident.write(row, 2, point['properties']['date'])
    incident.write(row, 3, point['properties']['p_type'])
    incident.write(row, 4, point['properties']['personal_involvement'])
    incident.write(row, 5, point['properties']['details'])
    incident.write(row, 6, point['properties']['impact'])
    incident.write(row, 7, point['properties']['injury'])
    incident.write(row, 8, point['properties']['trip_purpose'])
    incident.write(row, 9, point['properties']['regular_cyclist'])
    incident.write(row, 10, point['properties']['helmet'])
    incident.write(row, 11, point['properties']['road_conditions'])
    incident.write(row, 12, point['properties']['sightlines'])
    incident.write(row, 13, point['properties']['cars_on_roadside'])
    incident.write(row, 14, point['properties']['bike_lights'])
    incident.write(row, 15, point['properties']['terrain'])
    incident.write(row, 16, point['properties']['aggressive'])
    incident.write(row, 17, point['properties']['intersection'])
    incident.write(row, 18, point['properties']['witness_vehicle'])
    incident.write(row, 19, point['properties']['bicycle_type'])
    incident.write(row, 20, point['properties']['ebike'])
    incident.write(row, 21, point['properties']['ebike_class'])
    incident.write(row, 22, point['properties']['ebike_speed'])
    incident.write(row, 23, point['properties']['direction'])
    incident.write(row, 24, point['properties']['turning'])
    incident.write(row, 25, point['properties']['age'])
    if len(point['properties']['gender']) > 0:
        incident.write(row, 26, point['properties']['gender'][0])
    else:
        incident.write(row, 26, '')
    incident.write(row, 27, point['properties']['birthmonth'])
    # incident.write(row, 25, point['properties']['weather_summary'])
    # incident.write(row, 26, point['properties']['weather_sunrise_time'])
    # incident.write(row, 27, point['properties']['weather_sunset_time'])
    # incident.write(row, 28, point['properties']['weather_dawn'])
    # incident.write(row, 29, point['properties']['weather_dusk'])
    # incident.write(row, 30, point['properties']['weather_precip_intensity'])
    # incident.write(row, 31, point['properties']['weather_precip_probability'])
    # incident.write(row, 32, point['properties']['weather_precip_type'])
    # incident.write(row, 33, point['properties']['weather_temperature'])
    # incident.write(row, 34, point['properties']['weather_black_ice_risk'])
    # incident.write(row, 35, point['properties']['weather_wind_speed'])
    # incident.write(row, 36, point['properties']['weather_wind_bearing'])
    # incident.write(row, 37, point['properties']['weather_visibility_km'])
    incident.write(row, 28, point['properties']['pk'])
    incident.write(row, 29, point['geometry']['coordinates'][0])
    incident.write(row, 30, point['geometry']['coordinates'][1])


def writeToHazards(hazards, point, row):
    hazards.write(row, 0, point['properties']['i_type'])
    hazards.write(row, 1, point['properties']['date'])
    hazards.write(row, 2, point['properties']['p_type'])
    hazards.write(row, 3, point['properties']['details'])
    hazards.write(row, 4, point['properties']['age'])
    hazards.write(row, 5, point['properties']['birthmonth'])
    hazards.write(row, 7, point['properties']['pk'])
    hazards.write(row, 8, point['geometry']['coordinates'][0])
    hazards.write(row, 9, point['geometry']['coordinates'][1])
    if len(point['properties']['gender']) > 0:
        hazards.write(row, 10, point['properties']['gender'][0])
    else:
        hazards.write(row, 10, '')


def writeToThefts(theft, point, row):
    theft.write(row, 0, point['properties']['i_type'])
    theft.write(row, 1, point['properties']['date'])
    theft.write(row, 2, point['properties']['p_type'])
    theft.write(row, 3, point['properties']['details'])
    theft.write(row, 4, point['properties']['how_locked'])
    theft.write(row, 5, point['properties']['lock'])
    theft.write(row, 6, point['properties']['locked_to'])
    theft.write(row, 6, point['properties']['lighting'])
    theft.write(row, 8, point['properties']['traffic'])
    theft.write(row, 9, point['properties']['police_report'])
    theft.write(row, 10, point['properties']['police_report_num'])
    theft.write(row, 11, point['properties']['insurance_claim'])
    theft.write(row, 12, point['properties']['insurance_claim_num'])
    theft.write(row, 13, point['properties']['regular_cyclist'])
    theft.write(row, 14, point['properties']['pk'])
    theft.write(row, 15, point['geometry']['coordinates'][0])
    theft.write(row, 16, point['geometry']['coordinates'][1])


def writeToCSV(city, listOfCSVSheets):
    # Keep track of spot in list
    spotInList = 0
    # This is looping through the list
    for list in city:
        sheetToWrite = None
        incidentRow = 1
        hazardsRow = 1
        theftRow = 1

        # Incidents
        if spotInList == 0:
            for object in listOfCSVSheets:
                if object.name == 'Incident':
                    sheetToWrite = object
                    break
            for point in list:
                writeToIncident(sheetToWrite, point, incidentRow)
                incidentRow = incidentRow + 1
        # Hazards
        if spotInList == 1:
            for object in listOfCSVSheets:
                if object.name == 'Hazards':
                    sheetToWrite = object
                    break
            for point in list:
                writeToHazards(sheetToWrite, point, hazardsRow)
                hazardsRow = hazardsRow + 1
        # Thefts
        if spotInList == 2:
            for object in listOfCSVSheets:
                if object.name == 'Theft':
                    sheetToWrite = object
                    break
            for point in list:
                writeToThefts(sheetToWrite, point, theftRow)
                theftRow = theftRow + 1

        spotInList = spotInList + 1


def createAndWriteCSV(nameOfCSV, city):
    # Creates CSV File
    try:
        # The workbook path is weird because it needs to be an absolute path to have permission to enter the folder
        workbook = xlsxwriter.Workbook(os.path.abspath(os.path.dirname(__file__)) + os.sep + nameOfCSV,
                                       {'strings_to_urls': False})
    except Exception as ex:
        print(ex)
        time.sleep(5)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Call to createCSV Sheets
    listOfCSVSheets = createCSVSheets(workbook, bold)

    # Write to CSV
    writeToCSV(city, listOfCSVSheets)

    # Close workbook
    workbook.close()


world = False
test = False
# This should show an error if anything goes wrong.
if test:
    data = extractDataFromKML('Santa_Barbara.kml')
elif world:
    data = [[-179, -89],
            [179, -89],
            [179, 89],
            [-179, 89]]
else:
    data = extractDataFromKML(sys.argv[1])

bbx_val = [min(data, key=lambda sublist: sublist[0])[0], min(data, key=lambda sublist: sublist[1])[1],
           max(data, key=lambda sublist: sublist[0])[0], max(data, key=lambda sublist: sublist[1])[1]]
bbx_str = (",".join(repr(e) for e in bbx_val))

city = scrapeAndStore(data, bbx_str)
if test:
    createAndWriteCSV('Santa_Barbara.kml'[('Santa_Barbara.kml'.rfind(os.sep)) + 1:-4] + '.xlsx', city)
    print(bbx_str)
elif world:
    createAndWriteCSV('world.kml'[('world.kml'.rfind(os.sep)) + 1:-4] + '.xlsx', city)
else:
    createAndWriteCSV(sys.argv[1][(sys.argv[1].rfind(os.sep)) + 1:-4] + '.xlsx', city)

# except Exception as ex:
#     print(ex)
#     time.sleep(5)
