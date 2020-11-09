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

def getElementByGUID(guid):
    return eaRep.GetElementByGuid(guid)

def removeConnections(element):
    """ Removes all connections to this element """
    while element.Connectors.Count:
        element.Connectors.DeleteAt(0, True)
        element.Update()
        element.Refresh()

    if element.Connectors.Count == 0:
        print(element.Name + " -> Success: Connectors deleted")
    else:
        print(element.Name + " -> Error: Connectors not deleted")

def addConnection(source, target):
    """ Connects a source element (Requirement) to a target element (Design) """
    connection = source.Connectors.AddNew("", "Dependency")
    connection.SupplierID = target.ElementID
    connection.Direction = "Source -> Destination"
    connection.Update()
    source.Connectors.Refresh()
    target.Connectors.Refresh()

    if connection.IsConnectorValid():
        print("Success: Connection created: " + source.Name + " -> " + target.Name)
    else:
        print("Error: Connection failed: " + source.Name + " -> " + target.Name)

def parseElement(currentElement, indent):
    print(indent + currentElement.Name + " " + currentElement.Stereotype)
    
    addConnection(currentElement, getElementByGUID("{479F7F71-0EAF-4285-A5F7-4437BCC41AA9}"))

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

guid = "{219C6461-CE79-4f57-98FF-BE652608F8F6}"

removeConnections(getElementByGUID(guid))

#parseItem(eaRep.GetTreeSelectedObject())

"""
for package in eaRep.Models:
    parseItem(package)

print(otCollection)
"""


