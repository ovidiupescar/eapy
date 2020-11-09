import win32com.client
import sys

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



eaApp = EA_connect()
eaRep = eaApp.Repository

for package in eaRep.Package:
    print(package.Name)



