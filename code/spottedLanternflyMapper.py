# SpottedLanternflyMapper.py
#
#
# Author: Javed Hossain 04.26.2024
#
# Purpose: Generate yearly maps of the number of twitter posts about Spotted Lanternly (SLF) spread or sighting in the contiguous 48 states in US mainland. 
#
#
# Procedure Summary:  
#           --Take an excel file containing twitter posts about SLF along with date and location  
#           --Classify and count posts that are about SLF spread or sighting in each state
#           --Update the SLF count column in State Boundary feature class in ArcGIS Pro template project
#           --Export a single or multiple maps (depending on user preference)
#           --Export a CSV file containing classification results and locations extracted from the posts
#
#
# Main Steps:
# Step 1: Read excel file and load it into a Pandas dataframe
#
# Step 2: For each year, from start year till end year, classify posts about SLF spread/sighting created in that year                  
#            
# Step 3: Export classification results as a csv file if the user has asked for it                  
#
# Step 4: Count the number of such posts for each state and each year
#         
# Step 5: Create SLFMap objects that contain SLF post count for all states in a particular year
#
# Step 6: Based on the SLFMap objects use update cursor to update SLF count in the State Boundary polygon feature class
#
# Step 7: Export maps for each year in the specified year range
#
#
# Software Requirements: ArcGIS Pro arcpy package and spacy (with en_core_web_sm) must be installed.
#
#
# Usage: Input_File Header_Row Text_Column Location_Column Date_Column Map_Type Start_Year End_Year Output_Directory Export_Classification_Results
#
#
# Example input: "C:/GIS540_FinalProject/data/Test.xlsx" 5 "Full Text" "City Code" "Date" "Multiple" "2017" "2020" "C:\GIS540_FinalProject\output" true

import sys
import pandas
import re
import spacy
import arcpy
import gc
import os

#------------------------------# User Defined Class #------------------------------#
class SLFMap:
    """ 
    A class to store SLF sighting/spread type post count for each state.

    ...

    Attributes
    ----------
    title : str
        title of the map
    slfCount : dictionary
        dictionary with two letter state abbr. as keys and SLF count (int) as values

    Methods
    -------
    merge(anotherSlfMap):
        merges the passed SLFMap object with the current object
    """

    def __init__(self):
        """ Initialize this SLFMap's title and slfcount """
        self.title = "Untitled"
        # This dictionary stores the number of SLF sighting/spread type posts for each US state/territory
        self.slfCount = {'AK': 0, 'AL': 0, 'AR': 0, 'AS': 0, 'AZ': 0,
                         'CA': 0, 'CO': 0, 'CT': 0, 'DC': 0, 'DE': 0,
                         'FL': 0, 'GA': 0, 'MN': 0, 'HI': 0, 'IA': 0,
                         'ID': 0, 'IL': 0, 'IN': 0, 'KS': 0, 'KY': 0,
                         'LA': 0, 'MA': 0, 'MD': 0, 'ME': 0, 'MI': 0,
                         'MO': 0, 'MP': 0, 'MS': 0, 'MT': 0, 'NC': 0,
                         'ND': 0, 'NE': 0, 'NH': 0, 'NJ': 0, 'NM': 0,
                         'NV': 0, 'NY': 0, 'OH': 0, 'OK': 0, 'OR': 0,
                         'PA': 0, 'PR': 0, 'RI': 0, 'SC': 0, 'SD': 0,
                         'TN': 0, 'TX': 0, 'UT': 0, 'VA': 0, 'VI': 0,
                         'VT': 0, 'WA': 0, 'WI': 0, 'WV': 0, 'WY': 0}

    def merge(self, anotherSlfMap):
        """ Merge the passed SLFMap object with the current object (self) """
        for key in self.slfCount.keys():
            self.slfCount[key] = self.slfCount[key] + anotherSlfMap.slfCount[key]


#------------------------------# End of User Defined Class #------------------------------#

#------------------------------# User Defined Functions #------------------------------#

def contains(text, wordlist):
    """Return true if a word from the wordlist is found in text"""
    text = text.lower()
    for word in wordlist:
        if word in text:
            return True
    return False


def classify(text):
    """Return True for spread/sighting type posts, False for other types of posts"""
    # The words have spaces around them to prevent matching substrings
    wordlist1 = [" found ", " killed ", " spotted ", " attacked ", " attacking ", " caught ",
                 " saw ", " squished ", " stomped ", " discovered ", " quarantine ", " everywhere ",
                 " reported ", " seen ", " infested ", " stumbled ", " invade ", " observed "]

    wordlist2 = [" call ", " report ", " website ", " page ", " information "]

    text = text.lower()
    spread_sighting = False
    if " slf " in text and contains(text, wordlist1) and not contains(text, wordlist2):
        spread_sighting = True

    return spread_sighting


def cleanTextStep1(text):
    """Return text after performing some string replacements."""
    text = text.lower()
    # make references to SLF uniform
    text = text.replace("spotted lanternfly", "SLF")
    text = text.replace("spotted lanternflies", "SLF")
    text = text.replace("spotted lantern fly", "SLF")
    text = text.replace("spotted lantern flies", "SLF")
    text = text.replace("spottedlanternfly", "SLF")
    text = text.replace("spottedlanternflies", "SLF")
    text = text.replace("lanternfly", "SLF")
    text = text.replace("lanternflies", "SLF")
    text = text.replace("lantern fly", "SLF")
    text = text.replace("lantern flies", "SLF")
    # remove symbols
    text = text.replace("#", "")
    text = text.replace("@", "")
    # remove urls
    text = re.sub(r'http\S+', '', text)
    return text.strip()


def cleanTextStep2(spacy_doc, text):
    """Return text after removing everything except alphabets, digits and punctuations."""
    cleanText = ""
    for token in spacy_doc:
        if token.is_alpha or token.is_digit: # or token.is_punct:
            cleanText += " " + str(token)
            cleanText = cleanText.strip()
    return cleanText


def printMessage(message):
    """ Print message in both Python console and ArcGIS Pro """
    print(message)
    arcpy.AddMessage(message)


#------------------------------# End of User Defined Functions #------------------------------#

#------------------------------# Main Code Begins #------------------------------#

# Arguments:
# "C:/GIS540_FinalProject/data/Test.xlsx"
# "5"
# "Full Text"
# "City Code"
# "Date"
# "Multiple"
# "2017"
# "2020"
# "C:\GIS540_FinalProject"
# True

try:
    inputFilePath = sys.argv[1]
    headerRow = int(sys.argv[2])
    textColumn = sys.argv[3]
    locationColumn = sys.argv[4]
    dateColumn = sys.argv[5]
    outputType = sys.argv[6]
    startYear = sys.argv[7]
    endYear = sys.argv[8]
    outputDir = sys.argv[9]
    exportCSV = sys.argv[10]
except IndexError:
    arcpy.AddError("Arguments missing or invalid format")
    sys.exit(1)
except TypeError:
    arcpy.AddError("Invalid argument provided for header row.")
    sys.exit(1)

# If start and end years are same, than output a single map
if startYear == endYear:
    outputType = "Single"

# print progress
printMessage("Reading input file:")
printMessage(inputFilePath)

try:
    # Read input excel file
    # I am ignoring the Sheet0 part, I'm hoping you won't notice it.
    df = pandas.read_excel(inputFilePath, sheet_name="Sheet0")
except OSError:
    arcpy.AddError("Error reading input file. Please provide a path to an .xlsx file.")
    sys.exit(1)

# set header row
headers = df.iloc[headerRow]
# create a new dataframe that excludes the lines before the column header (values start from row 6)
new_df = pandas.DataFrame(df.values[headerRow+1:], columns=headers)

# create lists from the data columns
textList = new_df[textColumn].tolist()
locationList = new_df[locationColumn].tolist()
dateList = new_df[dateColumn].tolist()

# the pandas data frames are no longer needed, so delete them and call garbage collector
del df
del new_df
gc.collect()

# create file for writing results
if exportCSV == "true":
    # create text file writer for writing results
    try:
        resultWriter = open(outputDir + "\\Results.csv", "w", encoding="utf-8")
    except FileNotFoundError:
        arcpy.AddError("Invalid directory.")
        sys.exit(1)
    except PermissionError:
        arcpy.AddError("Results.csv file is locked by another program (possibly MS Excel).")
        sys.exit(1)
    # write the column names
    columns = "Index,Date,Classification,CityCode,ExtractedLocation,Text\n"
    resultWriter.write(columns)

# Load NLP model
try:
    model = spacy.load("en_core_web_sm")
except OSError:
    arcpy.AddError("Loading spacy model failed.")
    sys.exit(1)

# Create a list to keep track of last 100 posts, to avoid processing duplicate posts
last100posts = []

# These variables are for tracking progress
currentProgress = 0
previousProgress = 0

# print progress
printMessage("Processing text:")

# a dictionary of map objects with years as keys
mapDict = {}

# Count the number of rows that needs to be processed
numberOfRowsToProcess = 0
for i in range(len(locationList)):
    datetime = str(dateList[i]).split(" ")
    dateParts = datetime[0].split("-")
    postYear = dateParts[0]
    if startYear <= postYear <= endYear:
        numberOfRowsToProcess += 1

# in case of an invalid year range
if numberOfRowsToProcess == 0:
    arcpy.AddError("Error: Invalid year range")
    sys.exit(1)

numberOfProcessedRows = 0
# main work loop
for i in range(len(locationList)):
    # report progress
    try:
        currentProgress = int((numberOfProcessedRows / numberOfRowsToProcess) * 100)
    except ZeroDivisionError:
        arcpy.AddError("Please don't divide things with zero?")
        sys.exit(1)

    if currentProgress > previousProgress:
        if currentProgress % 10 == 0: # print progress percentage every 10%
            printMessage("Progress: " + str(currentProgress) + "%")
    previousProgress = currentProgress

    # the dates in the input file has YYYY-MM-DD HH:MM:SS.0 format
    # extracting month and year from datetime
    datetime = str(dateList[i]).split(" ")
    dateParts = datetime[0].split("-")
    postYear = dateParts[0]
    postMonth = dateParts[1]

    # process posts that were published within the specified year range 
    if startYear <= postYear <= endYear:
        # if a map object for current post year does not exist in map object dictionary, create one
        if postYear not in mapDict.keys():
            mapDict[postYear] = SLFMap()
            mapDict[postYear].title = postYear

        # copy the twitter post from the list to a variable
        text = textList[i]
        # Create spaCy NLP model for extracting geopolitical entities
        doc = model(text)       

        extractedLocation = ""
        # Loop through entities and extract locations (GPE) using spaCy
        for ent in doc.ents:
            if ent.label_ == 'GPE':
                extractedLocation += ent.text + ", "
        extractedLocation = extractedLocation.strip()
        extractedLocation = extractedLocation.rstrip(",")

        # Pandas thinks locations are floats for some reason, so casting it to a string
        postLocation = str(locationList[i])
        # Sometimes casting empty locations returns "nan", I don't want them
        if postLocation == 'nan':
            postLocation = ""

        # text cleanup step 1: perform string replacements
        text = cleanTextStep1(text)
        # Create spaCy NLP model for classification
        doc = model(text)
        # text cleanup step 2: only keep words and numbers, remove everything else
        text = cleanTextStep2(doc, text)

        # if the post is classified as spread/sighting and is not a duplicate post
        if classify(text) and text not in last100posts:
            classification = "Spread/Sighting"
            # add post at the end of the recent posts list
            last100posts.append(text)
            if len(last100posts) > 100:
                # remove the first post from the list
                last100posts.pop(0)

            # if location is available and is within the US
            if postLocation.startswith("USA."):
                locationParts = postLocation.split(".")
                postState = locationParts[1]
                
                # (if this state exists in the dictionary, it should, but what if it doesn't?)
                if postState in mapDict[postYear].slfCount.keys():
                    # Increment SLF sighting count for postState in current map object
                    mapDict[postYear].slfCount[postState] += 1

        else:  # the post is classified as other
            classification = "Other"

        if exportCSV == "true":
            # double quotes messes up the csv file, so replacing them with single quote
            text = textList[i].replace("\"", "'")
            # prepare current line for writing
            lineToWrite = (str(i+1) + "," + datetime[0] + "," + classification
                           + ",\"" + postLocation + "\",\"" + extractedLocation + "\"," + "\"" + text + "\"\n")
            # write line to file
            resultWriter.write(lineToWrite)

        # increment number of rows that have been processed
        numberOfProcessedRows += 1

del model
del last100posts

if exportCSV == "true":
    resultWriter.close()
    printMessage("Classification results exported to: ")
    printMessage(outputDir + "\\Results.csv")

# For single output type, merge all the SLFMap objects into one
if outputType == "Single":
    singleMap = SLFMap()
    for slfmap in mapDict.values():
        singleMap.merge(slfmap)
    mapDict.clear()
    if startYear == endYear:
        singleMapTitle = startYear
    else:
        singleMapTitle = startYear + " - " + endYear
    singleMap.title = singleMapTitle
    mapDict[singleMapTitle] = singleMap

# Open ArcGIS Pro project
try:
    # When working inside ArcGIS Pro,
    # use the project name "CURRENT"
    aprx = arcpy.mp.ArcGISProject("CURRENT")

except OSError:
    # Code was executed outside ArcGIS Pro
    # Use the project full path file name
    try:
        projectPath = "C:/Users/javed/OneDrive/Desktop/Spring_2024/GIS_540/FinalProject/SLF.aprx"
        aprx = arcpy.mp.ArcGISProject(projectPath)
    except OSError:
        arcpy.AddError("Couldn't find/open ArcGIS project at " + projectPath)
        sys.exit(1)

# Print message based on output type
if outputType == "Single":
    printMessage("Map exported to:")
else:
    printMessage("Maps exported to:")

# SLFMap dict contains multiple SLFMap objects with years as keys,
# each slfmap object contains a dictionary of state abbreviations as keys and SLF count as values.
# The word  "map" has become profoundly confusing at this point.
for m in mapDict.values():
    # Get a list of ArcGIS Pro Map objects.
    mapList = aprx.listMaps()
    # Get the first ArcGIS Pro Map object.
    myMap = mapList[0]
    # Get a list of Layer objects.
    layers = myMap.listLayers()
    # Get the first Layer object.
    myLayer = layers[0]

    # Get path to ArcGIS project and the State polygon feature class
    projectPath = aprx.filePath
    featureClass = projectPath.replace("SLF.aprx", "SLF.gdb\\States")

    # Update State boundary layer based on the current SLFMap object
    fields = ["STUSPS", "SLF_COUNT"]
    uc = arcpy.da.UpdateCursor(in_table=featureClass, field_names=fields) # Iterate through each row.
    for row in uc:
        if row[0] in m.slfCount.keys():
            row[1] = m.slfCount[row[0]]
        else:
            row[1] = 0
        uc.updateRow(row)
    del uc

    # Make sure the layer is visible
    myLayer.visible = True

    # Get layer's symbology object.
    sym = myLayer.symbology

    # Modify symbology renderer
    sym.updateRenderer('GraduatedColorsRenderer')
    sym.renderer.classificationField = 'SLF_COUNT'
    sym.renderer.classificationMethod = 'NaturalBreaks'
    sym.renderer.breakCount = 12
    sym.renderer.colorRamp = aprx.listColorRamps('Oranges (Continuous)')[0]
    # Update symbology renderer
    myLayer.symbology = sym

    # Get a list of Layout objects.
    layouts = aprx.listLayouts()
    # Get the first Layout object.
    myLayout = layouts[0]
    # Get the layout elements.
    elems = myLayout.listElements()
    # Get the map frame
    mapFrames = myLayout.listElements('MAPFRAME_ELEMENT')
    mapFrame = mapFrames[0]
    # loop through layout elements
    for e in elems:
        if 'Title' in e.name:
            # Update title
            e.textSize = 20
            e.text = "Spotted Lanternfly Spread/Sighting, US Mainland, " + m.title
            e.elementPositionX = mapFrame.elementPositionX + (mapFrame.elementWidth * 0.5) - (e.elementWidth * 0.5)

        if 'Legend' in e.name:
            # Position the legend at the center of the page below the mapframe
            e.elementPositionX = myLayout.pageWidth * 0.5

        if 'Map Frame' in e.name:
            # Set the scale to 1:20,000,000
            e.camera.scale = 20000000

    # Export maps as .png
    myLayout.exportToPNG(outputDir + "\\" + m.title + ".png", resolution=200)
    printMessage(outputDir + "\\" + m.title + ".png")

print("Done")

del aprx
del mapDict

#------------------------------# End of Code #------------------------------#
