"""Automates PRO/II air separation simulation: sets inputs, runs, and computes CAPEX/OPEX."""
import numpy as np
import win32com.client as win32
import os
from shutil import copyfile
import io as io
from cost_utils import (
	ASU_cost,
	Compressor_Cost,
	Dist_C,
	Exchanger_function,
	Expander_cost,
	Fp_tray,
	MSHE_COST,
	Tray_CP,
	Tray_cost,
	Towers_cost,
	pump_cost,
)

def read_original_prz( file_name ):
	f = io.open(file_name, buffering=-1, mode="rb")
	content = f.read()
	f.close()
	return content

def replace_content_file(file_name, content):
	f = io.open(file_name, buffering=-1, mode="wb")
	size = f.write(content)
	f.close()
	return size
	




def COMconnect(db_path):
    pro2 = "Nothing"
    pro2db = "Nothing"
    pro2 = win32.Dispatch("SimSciDbs.Database.102")

    init = pro2.Initialize()

    pro2.SetOption("showInternalObjects","0")
    pro2.SetOption("DoublePrecision","1")

    pro2.Import(os.path.splitext(db_path)[0]+'.inp')
    pro2db = pro2.OpenDatabase(db_path)

    #Get a security license (for better performance)
    pro2.GetSecuritySeat(2)

    return pro2, pro2db

#Disconnect COM interface
def COMdisconnect(pro2, pro2db):
    #Release the security license
    pro2_disc = pro2.ReleaseSecuritySeat()
    #Shut down the connection to the COM server
    pro2db = "Nothing"
    pro2 = "Nothing"

    return pro2, pro2db


db_path_origin= "D:\Bryan\ESPOL\Model_HR_3CDN.prz"
content = read_original_prz(db_path_origin)

db_path="Working_file.prz"
size = replace_content_file(db_path, content)
db_path_origin= "D:\Bryan\ESPOL\Model_HR_3CDN.prz"
content = read_original_prz(db_path_origin)


wdir =os.getcwd()

db_path = wdir + "\\"+"Working_file.prz"
size = replace_content_file(db_path, content)

pro2, pro2db = COMconnect(db_path)
pro2check = pro2db.CheckData



pro2, pro2db = COMconnect(db_path)
pro2check = pro2db.CheckData
pro2db = "Nothing"
pro2run = pro2.RunCalcs(db_path)
pro2db = pro2.OpenDatabase(db_path)



#### CHECK ####
if pro2run >= 4:
    print('first',pro2run)
    
    size = replace_content_file(db_path, content) 
    pro2run = pro2.RunCalcs(db_path)
    
#### INPUTS#####
# Simple Heat exchangers  

Temp_vector = np.array([303,303,313]) #((K) The temperature vector of this according to the ambient air conditions.
Exchangers_name = ["E6","E5","E7"]  #Exchangers in pump configuration

for i in range(len(Temp_vector)):
    Exchanger = pro2db.ActivateObject("Hx",Exchangers_name[i]) 
    Exchanger.PutAttribute(Temp_vector[i],"HotProdTempIn")
    Exchanger.Commit(True)
    print("Temp ",Exchanger.GetAttribute("HotProdTempIn"))
    
# Hot product liquid fraction
Liquid_fraction_vector = np.array([1, 0.9, 0.45])
Exchangers2_name =["E8", "E1", "E2"]  #Exchangers in Air Separation Process

for i in range(len(Liquid_fraction_vector)):
    Exchanger = pro2db.ActivateObject("Hx",Exchangers2_name[i])
    
    Exchanger.PutAttribute(Liquid_fraction_vector[i],"HotLiqFracIn")
    Exchanger.Commit(True)
    print("XL ",Exchanger.GetAttribute("HotLiqFracIn"))

HotDewIn_DTAD = 5.0  # K
DTAD = pro2db.ActivateObject("Hx","DTAD")
print("hot_dew ", DTAD.GetAttribute("HotDewIn"))
DTAD.PutAttribute(HotDewIn_DTAD,"HotDewIn")
DTAD.Commit(True)




"""
Duty_vector =np.array([2.9775,10.4166]) # x 10^6 Kcal/hr
Exchangers3_name = ["E3", "E4"]
for i in range(len(Duty_vector)):
    Exchanger = pro2db.ActivateObject("Hx",Exchangers3_name[i])
    print("Duty",Exchanger.GetAttribute("DutyIn"))    
    Exchanger.PutAttribute(Duty_vector[i],"DutyIn")
    Exchanger.Commit(True)
"""




#Destillation columns
Equip_names = ["HP","LP","ARG"]
Reflux_vector = np.array([1.527,0.972,0.416]) #Verificar dimensiones
HP_pressure = 587.685 #KPa  Rango de 580 a 600 KPa
LP_pressure = 118.550 #KPa  Rango de 100 a 150 KPa
ARG_pressure = 116.52 #KPa  Rango de 100 a 150 KPa
Top_tray_pressure_vector = np.array([HP_pressure, LP_pressure, ARG_pressure]) #KPa posiblemente

# Manipulated variable 
#Reflux Rate of Liquid
for i in range(len(Equip_names)):
    Reflux = Reflux_vector[i]
    Top_tray_pressure = Top_tray_pressure_vector[i]
    Column =pro2db.ActivateObject("ColumnIn",Equip_names[i])
    
    print("Reflux ",Column.GetAttribute("RefluxRateEst"))
    
    Column.PutAttribute(Reflux,"RefluxRateEst") 
    print("Top tray pressure ",Column.GetAttribute("TopTrayPressure"))
    Column.PutAttribute(Top_tray_pressure,"TopTrayPressure")
    Column.Commit(True)

# Condenser ARG Column
Duty_ARG = -3489.0 # 10^3 kcal/hr
ARG_Column = pro2db.ActivateObject("ColumnIn","ARG")
pro2db.CalculateUnitProps("ARG")
print("ARG Duty", ARG_Column.GetAttribute("HeaterDuties"))
ARG_Column.PutAttribute(Duty_ARG,"HeaterDuties")
ARG_Column.Commit(True)



#Correr la simulación
pro2check = pro2db.CheckData
pro2check = pro2db.DbsSaveDb
pro2db = "Nothing"

nMsg = pro2.MsgCount

if nMsg > 0:
    print(nMsg)
    for i in range(0, nMsg):
        print("Error message:", pro2.MsgText(i))
        
pro2run = pro2.RunCalcs(db_path)
pro2db = pro2.OpenDatabase(db_path)

if nMsg > 0:
    print(nMsg)
    for i in range(0, nMsg):
        print("Error message:", pro2.MsgText(i))

###OUTPUTS####
print()
print("Liquid Oxygen")

Stream_21 = pro2db.ActivateObject("Stream","21")
Oxygen_composition_Stream_21 = Stream_21.GetAttribute("TotalComposition",2)
Flowrate_Stream_21 = Stream_21.GetAttribute("TotalMolarRate")
Oxygen_flowrate_21 =Flowrate_Stream_21*Oxygen_composition_Stream_21 #Kmol/s
Oxygen_flowrate_21 =Oxygen_flowrate_21*32/(1000) # Ton/s

print("Flowrate_Stream_21",Flowrate_Stream_21)
print("Oxygen_flowrate_21",Oxygen_flowrate_21)
print("Oxygen_composition_Stream_21",Oxygen_composition_Stream_21)


######Liquid oxygen####
print()
Stream_1 = pro2db.ActivateObject("Stream","1")
Flowrate_Stream_1 = Stream_1.GetAttribute("TotalMolarRate")
print("Air")
print("Flowrate_Stream_1",Flowrate_Stream_1)

Oxygen_composition_Stream_1 = Stream_1.GetAttribute("TotalComposition",2)
Oxygen_flowrate_1 =Flowrate_Stream_1*Oxygen_composition_Stream_1
Oxygen_flowrate_1 =Oxygen_flowrate_1*32/(1000) # Ton/s
print("Oxygen_flowrate_1",Oxygen_flowrate_1)

Liq_Oxygen_Recovery = Oxygen_flowrate_21/Oxygen_flowrate_1*100


print("Recovery: ",Liq_Oxygen_Recovery)

### Gaseous Oxygen 
Stream_23 = pro2db.ActivateObject("Stream","23")
Flowrate_Stream_23 = Stream_23.GetAttribute("TotalMolarRate")
Oxygen_composition_Stream_23 = Stream_23.GetAttribute("TotalComposition",2)
Oxygen_flowrate_23 =Flowrate_Stream_23*Oxygen_composition_Stream_23
Oxygen_flowrate_23 =Oxygen_flowrate_23*32/(1000) # Ton/s

print("total recovery ", Liq_Oxygen_Recovery +Oxygen_flowrate_23*100/Oxygen_flowrate_1 )


###Heat from Exchangers####
Exchangers_name =["E6","E5","E7"]  #Exchangers in pump configuration
Exchangers_duty =np.zeros(len(Exchangers_name))
for i in range(len(Exchangers_name)):
    Exchanger = pro2db.ActivateObject("Hx",Exchangers_name[i]) 
    Exchangers_duty[i] = Exchanger.GetAttribute("DutyCalc")

###Expander###
PropEX1 = pro2db.ActivateObject("Expander","EX1")
WorkEX1 = PropEX1.GetAttribute("WorkActualCalc")

Efficiency = (WorkEX1/np.sum(Exchangers_duty))*100

print("Efficiency: ",Efficiency)

### Argon Flowate####
Stream_AL1 = pro2db.ActivateObject("Stream","AL1")
Argon_composition_Stream_AL1 = Stream_AL1.GetAttribute("TotalComposition",1)
Flowrate_Stream_AL1 = Stream_AL1.GetAttribute("TotalMolarRate")
Argon_flowrate_AL1 =39.948*Flowrate_Stream_AL1*Argon_composition_Stream_AL1/(1000)  #Ton/s
print()
print("Argon_flowrate_AL1", Argon_flowrate_AL1)


### Nitrogen Flowate####
Stream_20 = pro2db.ActivateObject("Stream","20")
Nitrogen_composition_Stream_20 = Stream_20.GetAttribute("TotalComposition",0)
Flowrate_Stream_20 = Stream_20.GetAttribute("TotalMolarRate")
Nitrogen_flowrate_20 =Flowrate_Stream_20*28*Nitrogen_composition_Stream_20/(1000) # Ton/s
print()
print("Nitrogen_flowrate_20", Nitrogen_flowrate_20)




#### Column ###
    
#HP Column
HP_Column = pro2db.ActivateObject("ColumnIn","HP")
HP_NumTr = HP_Column.GetAttribute("NumberOfTrays") - 1

TrayTemperature = HP_Column.GetAttribute("TrayTemperatures",-1)

# TrayTemperature = TrayTemperature - 273.15

# LiqRates = HP_Column.GetAttribute("TrayTotalLiqRates", -1) #mole/sec

# VapRates = HP_Column.GetAttribute("TrayTotalVaporRates", -1) #mole/sec

HP_Reboiler = HP_Column.GetAttribute("HeaterDuties", 0)

HP_Vapor = Reflux_vector[0]

HP_Diameter_list = HP_Column.GetAttribute("TrayDiameters", -1)
HP_Diameter = max(HP_Diameter_list)
if HP_Diameter <0:
    HP_Diameter = 3.048 #m
HP_Pressure = HP_Column.GetAttribute("TopTrayPressure")
HP_Tray_Spacing = 0.3048 #m  The usual one

HP_Volume= (np.pi*(HP_Diameter)**2)/4*HP_NumTr*HP_Tray_Spacing # Volumen de la Torre (m3)
HP_Asi = ((np.pi*(HP_Diameter)**2)/4) # Area de Plato(m2)
HP_Ptray = HP_Pressure*1.01 # bar


#LP Column
LP_Tray_Spacing = 0.3048 #m  The usual one
LP_Column = pro2db.ActivateObject("ColumnIn","LP")
LP_NumTr = LP_Column.GetAttribute("NumberOfTrays") - 1
LP_Diameter_list = LP_Column.GetAttribute("TrayDiameters", -1)
LP_Diameter = max(LP_Diameter_list)
if LP_Diameter <0:
    LP_Diameter = 3.524 #m
LP_Pressure = LP_Column.GetAttribute("TopTrayPressure")

LP_Volume= (np.pi*(LP_Diameter)**2)/4*LP_NumTr*LP_Tray_Spacing # Volumen de la Torre (m3)
LP_Asi = ((np.pi*(LP_Diameter)**2)/4) # Area de Plato(m2)
LP_Ptray = LP_Pressure # bar

#ARG Column
ARG_Column = pro2db.ActivateObject("ColumnIn","ARG")
ARG_NumTr = ARG_Column.GetAttribute("NumberOfTrays") - 1

ARG_Tray_spacing = 0.6098 #m
ARG_Diameter_list = ARG_Column.GetAttribute("TrayDiameters", -1)
ARG_Diameter = max(ARG_Diameter_list)
if ARG_Diameter <0:
    ARG_Diameter = 1.829 #m
ARG_Pressure = ARG_Column.GetAttribute("TopTrayPressure")


ARG_Volume= (np.pi*(ARG_Diameter)**2)/4*ARG_NumTr*ARG_Tray_spacing # Volumen de la Torre (m3)
ARG_Asi = ((np.pi*(ARG_Diameter)**2)/4) # Area de Plato(m2)
ARG_Ptray = ARG_Pressure*1.01 # bar


Column_Volumes  = [HP_Volume,LP_Volume,ARG_Volume]  
Column_Pressure = [HP_Pressure,LP_Pressure,ARG_Pressure]  
Column_FM =1
Total_cost_Columns =sum(Towers_cost(Column_Volumes,Column_Pressure,Column_FM))
Column_trays = ["Valve","Sieve","Sieve"]
Column_Area = np.array([HP_Diameter,LP_Diameter,ARG_Diameter])**2*np.pi/4
Column_NumTr =[HP_NumTr,LP_NumTr,ARG_NumTr]

Total_Tray_cost =Tray_cost(Column_Area[0] ,Column_NumTr[0],1,Column_trays[0])+Tray_cost(Column_Area[1] ,Column_NumTr[1],1,Column_trays[1])+Tray_cost(Column_Area[2] ,Column_NumTr[2],1,Column_trays[2])


##Heat exchangers



Total_heat_exchangers_name = ["DTAD","E1","E2","E5","E6","E7","E8"]
Total_heat_exchangers_A = np.zeros(len(Total_heat_exchangers_name))
Total_heat_exchangers_pressure = np.zeros(len(Total_heat_exchangers_name))
Total_heat_exchangers_duty =np.zeros(len(Total_heat_exchangers_name)) #KJ/s
Total_heat_exchangers_U = np.ones(len(Total_heat_exchangers_name))*5 #kW/(K.m2)   Revisar las U para intercambiador

for i in range(len(Total_heat_exchangers_name)):
   
    Exchanger = pro2db.ActivateObject("Hx",Total_heat_exchangers_name[i])
    Total_heat_exchangers_duty[i] = Exchanger.GetAttribute("DutyCalc")    #KJ/s
    print ("Duty Calc",Total_heat_exchangers_duty[i])
    LMTD = Exchanger.GetAttribute("LmtdCalc") 
    
    Total_heat_exchangers_pressure[i] =Exchanger.GetAttribute("HotInletPressure")
    Total_heat_exchangers_pressure[i]=Total_heat_exchangers_pressure[i]/100 # bar
    UA = Total_heat_exchangers_duty[i]/LMTD
    Total_heat_exchangers_A[i] = UA/Total_heat_exchangers_U[i] # Area del intercambiador (m2)

FM = 3.1   ## Revisar materiales de cada intercambiador
Total_heat_exchangers_cost = np.sum(Exchanger_function(Total_heat_exchangers_A,Total_heat_exchangers_pressure,FM))

L =["E3","E4"]




print("Total heat_exchangers_cost:",Total_heat_exchangers_cost)



### Compressors####

### Revisar COM

Compressors_name = ["C1","C2","C3"]
Compressors_Work = np.zeros(len(Compressors_name))
Compressors_Pressure =np.zeros(len(Compressors_name))
for i in range(len(Compressors_name)):
    Compressor = pro2db.ActivateObject("Compressor",Compressors_name[i])
    Compressors_Work[i] = Compressor.GetAttribute("WorkActualCalc")
    Compressors_Pressure[i] = Compressor.GetAttribute("PressOutCalc")
    Compressors_Pressure[i] = Compressors_Pressure[i]/100 # bar

Total_cost_compressor = sum(Compressor_Cost(Compressors_Pressure-1,Compressors_Work,FM) )

COM_compressor=pro2db.ActivateObject("Compressor","COM")
COM_work = COM_compressor.GetAttribute("WorkActualCalc")
COM_stream=pro2db.ActivateObject("Stream","4")
COM_pressure = COM_stream.GetAttribute("Pressure")

Total_cost_compressor = Total_cost_compressor + Compressor_Cost(COM_pressure,COM_work,FM)



print("Compressors cost: ",Total_cost_compressor)


## Expander####

Expanders_name =["EXP","EX1"]
Expanders_Work = np.zeros(len(Expanders_name))
Expanders_Pressure =np.zeros(len(Expanders_name))
for i in range(len(Expanders_name)):
    Expander = pro2db.ActivateObject("Expander",Expanders_name[i])
    Expanders_Work[i] = Expander.GetAttribute("WorkActualCalc")
    Expanders_Pressure[i] = Expander.GetAttribute("PressOutCalc")/100 # bar
Total_cost_expander = sum(Expander_cost(Expanders_Pressure-1,Expanders_Work,FM) )


### FlashDrum
Flash_Drum=pro2db.ActivateObject("Flash","DEW")
Flash_Duty=Flash_Drum.GetAttribute("DutyCalc") 
print("Flash_Duty: ",Flash_Duty)

###MSHE###
MSHE_1_volume=10 #m^3
MSHE_1_stream=pro2db.ActivateObject("Stream","8_R1")
MSHE_1_pressure= MSHE_1_stream.GetAttribute("Pressure")
MSHE_1_pressure =MSHE_1_pressure/100
MSHE_1_COST = MSHE_COST(MSHE_1_volume,MSHE_1_pressure)
MSHE_2_volume=10 #m^3
MSHE_2_stream=pro2db.ActivateObject("Stream","4")
MSHE_2_pressure= MSHE_2_stream.GetAttribute("Pressure")
MSHE_2_pressure =MSHE_2_pressure/100
MSHE_2_COST = MSHE_COST(MSHE_2_volume,MSHE_2_pressure)




#### CashFlow ####




Oxygen_price = 1100 #$/ton
Nitrogen_price = 3400 #$/ton
Argon_price  = 222.66 #$/ton


CF = Oxygen_price*Oxygen_flowrate_21+Nitrogen_price*Nitrogen_flowrate_20+Argon_price*Argon_flowrate_AL1   # $/s
CF =  CF*3600*24*365 #$/year
print("CF", CF, "$/año")

CP_total = Total_cost_compressor+Total_heat_exchangers_cost +Total_cost_expander+Total_cost_Columns+Total_Tray_cost+MSHE_1_COST+MSHE_2_COST
CAPEX = 1.18*CP_total
print("CAPEX: ",CAPEX)

### Raw Material Cost ####
Air_cost= 10 #$/ton
RMC = (Flowrate_Stream_1*Air_cost*29/1000)*3600*24*365 #$/year
print("RMC:",RMC)

##### Utility cost based on CAPCOST 2017
UC_HX = 8.49  #low refrigeration USD/GJ
UC_elec = 18.72 #electricity USD/GJ
Total_work_compressor = sum(Compressors_Work)
Total_exchangers_duty = sum(Total_heat_exchangers_duty)

UC = (Total_exchangers_duty/(10**6))*UC_HX+ (Total_work_compressor/(10**6))*UC_elec
UC = UC*3600*24*365 #$/year
print("UC:", UC)
#OPEX per year

OPEX = 0.18*CAPEX + 1.23*(RMC+UC) 
print("OPEX: ", OPEX)
print("Total CF", CF-OPEX)
print(pro2.GenerateReport(db_path))

#Cerrar la simulación
COMdisconnect(pro2, pro2db)




