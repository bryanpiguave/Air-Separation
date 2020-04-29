import numpy as np
import win32com.client as win32
import os

def COMconnect(db_path):
    pro2 = "Nothing"
    pro2db = "Nothing"
    pro2 = win32.Dispatch("SimSciDbs.Database.102")

    pro2.Initialize()

    pro2.SetOption("showInternalObjects","1")
    pro2.SetOption("DoublePrecision","1")

    pro2.Import(os.path.splitext(db_path)[0]+'.inp')

    pro2db = pro2.OpenDatabase(db_path)

    #Get a security license (for better performance)
    pro2.GetSecuritySeat(2)

    return pro2, pro2db

def COMdisconnect(pro2, pro2db):
    #Release the security license
    pro2.ReleaseSecuritySeat()

    #Shut down the connection to the COM server
    pro2db = "Nothing"
    pro2 = "Nothing"

    return pro2, pro2db

db_path = os.path.join(".","Casebook_Air_Separation_Plant.prz")
db_path = r"C:\PythonEnhanced\Repositorio\Air-Separation\Casebook_Air_Separation_Plant.prz"
pro2, pro2db = COMconnect(db_path)
pro2check = pro2db.CheckData
pro2run = pro2.RunCalcs(db_path)
pro2db = pro2.OpenDatabase(db_path)


classcount = pro2.GetClassCount

#pro2, pro2db = COMdisconnect(pro2, pro2db)    


