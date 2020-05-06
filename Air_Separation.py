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

db_path = r"C:\PythonEnhanced\Repositorio\Air-Separation\Casebook_Air_Separation_Plant.prz"
pro2, pro2db = COMconnect(db_path)
pro2check = pro2db.CheckData



#Get all the classes from PROII
classcount = pro2.GetClassCount
Classes = list()
for i in range(classcount):
    Classes.append(pro2.GetClassNames(i))
    
#Get all the class from Unit group   
Unit_count = pro2.GetGroupClassCount("Unit")
Unit_classes=list()
for i in range(Unit_count):
    Unit_classes.append(pro2.GetGroupClassNames("Unit",i))

#Get all the groups of classes in PROII
Class_groups_count = pro2.GetGroupCount
Class_groups = list()
for i in range(Class_groups_count):
    Class_groups.append(pro2.GetGroupNames(i))




#Object names - Example: Stream
Stream_names_count=pro2db.GetObjectCount("Stream")
Stream_names_list =list()
for i in range(Stream_names_count):
    name =pro2db.GetObjectNames("Stream",i)
    Stream_names_list.append(name)

#Object properties - Example: Stream
Stream_object =pro2db.ActivateObject("Stream","1")

Stream_properties_count=Stream_object.GetAttributeCount
Stream_properties =list()
for i in range(Stream_properties_count):
    s_property =Stream_object.GetAttributeName(i)
    Stream_properties.append(s_property)




















#Correr la simulación
pro2run = pro2.RunCalcs(db_path)
pro2.GenerateReport(db_path)



#Cerrar la simulación
COMdisconnect(pro2, pro2db)




