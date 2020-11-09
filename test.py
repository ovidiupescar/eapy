import win32com.client
import sys





try:
    eaApp = win32com.client.Dispatch("EA.App")
except:
    sys.stderr.write( "Unable to connect to EA\n" )
    exit()

# from here onwards, the full EA API is available as python objects
# so we can

mEaRep = eaApp.Repository

if mEaRep.ConnectionString == '':
    sys.stderr.write( "EA has no Model loaded\n" )
    exit()


print("connecting...", mEaRep.ConnectionString)


for model in mEaRep.Models:
    print(model.Name)

print(mEaRep.GetTreeSelectedPackage().Name)

for package in mEaRep.GetTreeSelectedPackage().Packages:
    print(package.Name)