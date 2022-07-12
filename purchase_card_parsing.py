import os
from pydoc import doc
from typing import final
from xml.dom.minidom import Document
import zipfile
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom
from glob import glob
import win32com.client as win32
from win32com.client import constants
# probably going to use pandas DF for the table data
import pandas



#               this part doesnt work

def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)




# removes XML tags
def cleanFinalString(FS):
    FSCleaned = (FS.replace("<w:t>", "")).replace("<w:t xml:space=\"preserve\">", "")
    return(FSCleaned)




# this is where code that works begins
# the first line uses a test docx that is just the og doc using online converter, adjust as neccesary 
# when or if python code can convert file

document = zipfile.ZipFile("test3.docx")
# parses body of the docx and pretty prints the xml code 
uglyXML = xml.dom.minidom.parseString( document.read("word/document.xml") ).toprettyxml(indent = "   ")
text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</')
prettyXML = text_re.sub('>\g<1></', uglyXML)

lengthOfString =  len(prettyXML)
startingIndex = prettyXML.find("<w:t", 0)
finishingIndex = startingIndex
tagOne = "<w:t>"
tagTwo = "<w:t xml:space=\"preserve\">"
finalString = ""
# loop searches the xml to find info and concatenates all of it into one string
while finishingIndex < lengthOfString:
    tagOneIndex = prettyXML.find(tagOne, finishingIndex)
    tagTwoIndex = prettyXML.find(tagTwo, finishingIndex)
    startingIndex = min(tagOneIndex, tagTwoIndex)
    
    finishingIndex = prettyXML.find("</w:t>", startingIndex)
    finalString = finalString + prettyXML[startingIndex:finishingIndex] + " "
    if finishingIndex < 0:
        finishingIndex = lengthOfString + 1


# removes extra tags
finalString = cleanFinalString(finalString)
finalList = []

# this is where parsing begins

# finding part number
partNumSIndex = finalString.find("PART NUMBER", 0)
partNumFIndex = finalString.find("DESCRIPTION", 0)
partNum = finalString[partNumSIndex + 11:partNumFIndex]
finalList.append(partNum)

# adjusting finalString so it doesnt include the part we just parsed
finalString = finalString[partNumFIndex:len(finalString)]

# finding description
DescriptionSIndex = finalString.find("DESCRIPTION", 0)
DescriptionFIndex = finalString.find("STANDARD COST", 0)
Desc = finalString[DescriptionSIndex + 11:DescriptionFIndex]
finalList.append(Desc)

finalString = finalString[DescriptionFIndex:len(finalString)]

# finding standard cost
SCSIndex = finalString.find("STANDARD COST", 0)
SCFIndex = finalString.find("ACCT.DIST.", 0)
SC = finalString[SCSIndex + 13:SCFIndex]
finalList.append(SC)

finalString = finalString[SCFIndex:len(finalString)]

# finding ACCT.DIST.
ADSIndex = finalString.find("ACCT.DIST.", 0)
ADFIndex = finalString.find("NOTE", 0)
AD = finalString[ADSIndex + 10:ADFIndex]
finalList.append(AD)

finalString = finalString[ADFIndex:len(finalString)]

# finding notes
NoteSIndex = finalString.find("NOTE", 0)
NoteFIndex = finalString.find("DATE", 0)
NOTE = finalString[NoteSIndex + 4:NoteFIndex]
finalList.append(NOTE)

finalString = finalString[ (finalString.find("PRICE ", 0)) + 6 : len(finalString)]

finalDict = dict()
finalList.append(finalDict)

# finding date
DateSIndex = finalString.find("REQUEST DATE ", 0)
finalString = finalString[13:len(finalString)]

DateFIndex = finalString.find(" ", 0)
DATE = finalString[DateSIndex:DateFIndex]

dateList = []
dateList.append(DATE)
finalDict["Date"] = dateList

finalString = finalString[DateFIndex:len(finalString)]

# finding vendor 
VenSIndex = finalString.find(" ", 0)
VenFIndex = finalString.find(" ", 1)
Ven = finalString[VenSIndex:VenFIndex]

vendorList = []
vendorList.append(Ven)
finalDict["Vendor"] = vendorList

finalString = finalString[VenFIndex:len(finalString)]

# finding order no. 
ONSIndex = finalString.find(" ", 0)
ONFIndex = finalString.find(" ", 1)
ON = finalString[ONSIndex:ONFIndex]

ONList = []
ONList.append(ON)
finalDict["Order Number"] = ONList

finalString = finalString[ONFIndex:len(finalString)]

# finding quantity 
QSIndex = finalString.find(" ", 0)
QFIndex = finalString.find(" ", 1)
Q = finalString[QSIndex:QFIndex]

QList = []
QList.append(Q)
finalDict["Quantity"] = QList

finalString = finalString[QFIndex:len(finalString)]

# finding price 
PriceSIndex = finalString.find(" ", 0)
PriceFIndex = finalString.find(" ", 1)
Price = finalString[PriceSIndex:PriceFIndex]

PriceList = []
PriceList.append(Price)
finalDict["Price"] = PriceList

finalString = finalString[PriceFIndex:len(finalString)]

# finding request date
RDSIndex = finalString.find(" ", 0)
RDFIndex = finalString.find(" ", 1)
RD = finalString[RDSIndex:RDFIndex]

RDList = []
RDList.append(RD)
finalDict["Request Date"] = RDList

finalString = finalString[RDFIndex:len(finalString)]

# finding location (?)
LocSIndex = finalString.find(" ", 0)
LocFIndex = min(finalString.find("S.C.", 1), finalString.find("/", 1) - 2 ) 
Loc = finalString[LocSIndex:LocFIndex]

LocList = []
LocList.append(Loc)
finalDict["Location"] = LocList

finalString = finalString[LocFIndex:len(finalString)]
finalString.strip()






# if the table has more than one line, this if block parses through it
if (finalString.find("S.C.", 0) > 10):
    finalString = finalString[0:finalString.find("S.C.", 0) + 4]
    

    # the code inside the while loop is basically the same code repeated
    while (finalString.find("S.C.", 0) > 10):

        finalString = finalString.strip()

        # finding date
        DateSIndex = 0
        DateFIndex = finalString.find(" ", 0)
        DATE = finalString[DateSIndex:DateFIndex]

        dateList.append(DATE)
        finalDict["Date"] = dateList

        finalString = finalString[DateFIndex:len(finalString)]

        # finding vendor 
        VenSIndex = finalString.find(" ", 0)
        VenFIndex = finalString.find(" ", 1)
        Ven = finalString[VenSIndex:VenFIndex]

        vendorList.append(Ven)
        finalDict["Vendor"] = vendorList

        finalString = finalString[VenFIndex:len(finalString)]

        # finding order no. 
        ONSIndex = finalString.find(" ", 0)
        ONFIndex = finalString.find(" ", 1)
        ON = finalString[ONSIndex:ONFIndex]

        ONList.append(ON)
        finalDict["Order Number"] = ONList

        finalString = finalString[ONFIndex:len(finalString)]

        # finding quantity 
        QSIndex = finalString.find(" ", 0)
        QFIndex = finalString.find(" ", 1)
        Q = finalString[QSIndex:QFIndex]

        QList.append(Q)
        finalDict["Quantity"] = QList

        finalString = finalString[QFIndex:len(finalString)]

        # finding price 
        PriceSIndex = finalString.find(" ", 0)
        PriceFIndex = finalString.find(" ", 1)
        Price = finalString[PriceSIndex:PriceFIndex]

        PriceList.append(Price)
        finalDict["Price"] = PriceList

        finalString = finalString[PriceFIndex:len(finalString)]

        # finding request date
        RDSIndex = finalString.find(" ", 0)
        RDFIndex = finalString.find(" ", 1)
        RD = finalString[RDSIndex:RDFIndex]

        RDList.append(RD)
        finalDict["Request Date"] = RDList

        finalString = finalString[RDFIndex:len(finalString)]

        # finding location (?)
        LocSIndex = finalString.find(" ", 0)
        LocFIndex = min(finalString.find("S.C.", 1), finalString.find("/", 1) - 2 )

        Loc = finalString[LocSIndex:LocFIndex]

        LocList.append(Loc)
        finalDict["Location"] = LocList

        finalString = finalString[LocFIndex:len(finalString)]

    
# printing info found above table
print("Part number: " + partNum)
print("Description: " + Desc)
print("Standard Cost: " + SC)
print("Acct. Dist.: "  + AD)
print("Notes: "  + NOTE)

# the final data type is a list with string elements for the part number, description, standard cost, acct. dist., notes, and...
# ... then the last element is a dictionary that represents the table on the card. The keys for the dictionary is the column...
# headings and the values are lists containing all the row values of that column
print(finalList)











