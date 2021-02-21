from typing import Type
import win32com.client
import sys
from object_types import *
import pandas as pd
import time

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

def addConnection(source, target, con_type):
    """ Connects a source element (Requirement) to a target element (Design) """
    connection = source.Connectors.AddNew("", con_type)
    connection.SupplierID = target.ElementID
    connection.Direction = "Source -> Destination"
    connection.Update()
    source.Connectors.Refresh()
    target.Connectors.Refresh()

    if connection.IsConnectorValid():
        print("Success: Connection created: " + source.Name + " -> " + target.Name)
    else:
        print("Error: Connection failed: " + source.Name + " -> " + target.Name)

def detelePorts(currentElement):
    if currentElement.Type == "Component":
        if currentElement.EmbeddedElements.Count > 0:
            for i in range(0, currentElement.EmbeddedElements.Count):
                #eaCollection(currentElement.EmbeddedElements):
                currentElement.EmbeddedElements.Delete(i)
    currentElement.EmbeddedElements.Refresh()

def parseElement(currentElement, indent):
    global components
    print(indent + currentElement.Name + " Stereotype: " + currentElement.Stereotype + " Type: " + currentElement.Type)
    #components.append([currentElement.Name, currentElement.Type, currentElement.ElementGUID])
    components["Name"].append(currentElement.Name)
    components["Type"].append(currentElement.Type)
    components["GUID"].append(currentElement.ElementGUID)
    
    """To delete ports, uncomment bellow"""
    #deletePorts(currentElement)
    

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


def main():
    eaApp = EA_connect()
    eaRep = eaApp.Repository

    #guid = "{219C6461-CE79-4f57-98FF-BE652608F8F6}"

    #removeConnections(getElementByGUID(guid))
    global components
    components["Name"] = []
    components["Type"] = []
    components["GUID"] = []

    parseItem(eaRep.GetTreeSelectedObject())

    df = pd.DataFrame(components)
    print(df)

    """
    for package in eaRep.Models:
        parseItem(package)

    print(otCollection)
    """

def add_port(component, pname, ptype, pstereotype):
    new_port = component.EmbeddedElements.AddNew(pname, ptype)
    new_port.Stereotype = pstereotype
    new_port.Update()
    component.EmbeddedElements.Refresh()
    return new_port

def add_interface(p, r, p_type, r_type):
    pass

def add_ports():
    component = getElementByGUID("{DD5CB9B7-DFA5-49f0-A993-8EAB5EE4D4EA}")
    #add_port(component, "test", "RequiredInterface", "Receiver")
    port = getElementByGUID("{82A5764D-5D38-4533-83AF-7A4954A8E712}")
    
    new_sender = add_port(component, "primeste_iar", "ProvidedInterface", "Sender")
    new_receiver = add_port(component, "trimite_iar", "RequiredInterface", "Receiver")
    addConnection(new_sender, new_receiver, "dependency")

    #print(f"Object: {new_port.ObjectType}, Name: {new_port.Name}, GUID: {new_port.ElementGUID}")
    #print(f"Object: {component.Name}")


components = {}
start = time.time()
eaApp = EA_connect()
eaRep = eaApp.Repository
thePackage = eaRep.GetTreeSelectedPackage()
add_ports()
end = time.time()
print("Duration: {}s".format(end-start))


