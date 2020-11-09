import win32com.client
import sys
from object_types import *

#path to req


def EA_connect():
    try:
        eaApp = win32com.client.Dispatch("EA.App")
    except:
        sys.stderr.write( "Unable to connect to EA\n" )
        return None

    # from here onwards, the full EA API is available as python objects
    # so we can

    mEaRep = eaApp.Repository

    if mEaRep.ConnectionString == '':
        sys.stderr.write( "EA has no Model loaded\n" )
        return None

    return eaApp

def eaCollection(C):
    if C.ObjectType == otCollection:
        for i in range(C.Count):
            yield C.GetAt(i)
    else:
        raise TypeError

    

def parseElement(currentElement, indent):
    print(indent + currentElement.Name + " " + currentElement.Stereotype)

    if currentElement.Elements.Count > 0:
        for element in eaCollection(currentElement.Elements):
            parseElement(element, indent+" ")

def parsePackage(currentPackage, indent):
    print(indent + currentPackage.Name + " Package")

    for element in eaCollection(currentPackage.Elements):
        parseElement(element, indent+" ")

    for package in eaCollection(currentPackage.Packages):
        parsePackage(package, indent+" ")


def parseItem(currentItem):
    if currentItem.ObjectType == otPackage:
        #iterate through packages
        parsePackage(currentItem, "")
    elif currentItem.ObjectType == otElement:
        parseElement(currentItem, "")
    else:
        print("Selection must be element or package.")



eaApp = EA_connect()
eaRep = eaApp.Repository

parseItem(eaRep.GetTreeSelectedObject())

"""
for package in eaRep.Models:
    parseItem(package)

print(otCollection)
"""


