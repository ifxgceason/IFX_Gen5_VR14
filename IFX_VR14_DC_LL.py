
#import library
import clr
import sys
import time
import datetime
import numpy as np
import pandas as pd
import os
import matplotlib.pyplot as plt
import ifx

# set Gen5 folder
dll_path=r"C:\GEN5\PythonScripts"
sys.path.append(dll_path)

# import Gen5 API
from Common.APIs import DataAPI
from Common.APIs import DisplayAPI
from Common.Models import TriggerModel
from Common.Models import FanModel
from Common.Models import RawDataModel
from Common.Models import DataModel
from Common.Models import DisplayModel
from Common.Enumerations import HorizontalScale
from Common.Enumerations import ScoplessChannel
from Common.Enumerations import Cursors
from Common.APIs import MeasurementAPI
from Common.APIs import GeneratorAPI
#from CommonBL import *
from Common.Enumerations import PowerState
from Common.Enumerations import Transition
from Common.Enumerations import PowerState
from Common.Enumerations import Transition
from Common.Enumerations import Protocol
from Common.Enumerations import RailTestMode
from System import String, Char, Int32, UInt16, Boolean, Array, Byte, Double
from System.Collections.Generic import List
from System.Collections.Generic import Dictionary

# import excel
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook

gen = GeneratorAPI()
rails = gen.GetRails()
display = DisplayAPI()
Measurment = MeasurementAPI()
data=DataAPI()

# define function
# eason.chen@infineon.com

#What Rails are available?
def check_rails_avaliable():
    print("Displaying rails available")
    print("------------------------------------------")
    t = gen.GetRails()
    for i in range(0, len(t)):
            print(t[i].Name, "Vid Table - ", t[i].VID, " Max Current - ",
                      t[i].MaxCurrent, "; Vsense - ", t[i].VSenseName)
            print("Load sections: " + t[i].LoadSections)
            for k in t[i].ProtoParm.Keys:
                    print("      ", k, " ", t[i].ProtoParm[k])
    return None
	
def dcLoadlinetest(DRIVE_ONE_RAIL,svid_bus,svid_addr,Vout,Iout_currents,dtime,svid_reg_value):
    MeasurementAPI().ClearAllMeasurements()
    results=pd.DataFrame()
    # Change Fan Speed
    data.SetFanSpeed(10)

    # Assign rails to Tabs
    gen.AssignRailToDriveOne(DRIVE_ONE_RAIL)
    #gen.AssignRailToDriveTwo(DRIVE_TWO_RAIL)

    # Set H Scale
    display.SetHorizontalScale(HorizontalScale.Scale10us)

    # Set trigger
    data.SetTrigger(4, 0, 0, 0, 0, False)

    # Set Voltage
    gen.SetVoltageForRail(DRIVE_ONE_RAIL, Vout, Transition.Fast)

    # Enable Voltage/Current mean
    time.sleep(0.25)
    display.Ch1_2Rail(DRIVE_ONE_RAIL)
    time.sleep(0.25)
    Measurment.MeasureCurrentMean(DRIVE_ONE_RAIL)
    Measurment.MeasureVoltageMean(DRIVE_ONE_RAIL)
    Measurment.MeasureEPod(DRIVE_ONE_RAIL)
    Measurment.MeasureVoltageAmplitude(DRIVE_ONE_RAIL)
    Measurment.MeasurePersistentVoltageMinMax(DRIVE_ONE_RAIL)
    Measurment.MeasurePersistentCurrentMinMax(DRIVE_ONE_RAIL)

    # Enable scope traces
    display.SetChannel(ScoplessChannel.Ch1, True)
    display.SetChannel(ScoplessChannel.Ch2, True)

    # Set vertical scale
    display.SetVerticalVoltageScale(0, 2)
    display.SetVerticalCurrentScale(0, 500)

    tt = Measurment.GetEPodOnce(DRIVE_ONE_RAIL)
    try:
        VIN = (float)(tt[20:25])
    except ValueError:
        print('Note - No Epod detected')
      

    ## turn on VR13.HC mode
    print("Turn on VR13.HC bit")
    data.SetSvidCmdWrite(0,5,0x2A,1,5,2,0,3)
    data.SetSvidCmdWrite(0,6,0xC2,1,5,2,0,3)

    #set Vout voltage.
    print(f'set Vout to {Vout}V')
    gen.SetVoltageForRail(DRIVE_ONE_RAIL,Vout,Transition.Slow)

    result_temp={}
    print("")
    print('  ..................................................')
    print('  start to run LL test')
    print('  ..................................................')
	
    rail_name=DRIVE_ONE_RAIL
    if rail_name=="VCCFA_EHV_FIVRA":
        
        gen.AssignRailToDriveOne("VCCFAEHVFIVRANW")
        gen.AssignRailToDriveTwo("VCCFAEHVFIVRANE")
        gen.AssignRailToDriveThree("VCCFAEHVFIVRASW")
        gen.AssignRailToDriveFour("VCCFAEHVFIVRASE")
        gen.SetVoltageForRail("VCCFAEHVFIVRASE", Vout, Transition.Fast)
        for Iout_current in Iout_currents:
            EHV_FIVRA_list=["VCCFA_EHV_FIVRA_NW","VCCFA_EHV_FIVRA_NE","VCCFA_EHV_FIVRA_SW","VCCFA_EHV_FIVRA_SE"]
            #EHV_FIVRA_load_dict={"VCCFA_EHV_FIVRA_NW":"VCCFAEHVFIVRANW","VCCFA_EHV_FIVRA_NE":"VCCFAEHVFIVRANE","VCCFA_EHV_FIVRA_SW":"VCCFAEHVFIVRASW","VCCFA_EHV_FIVRA_SE":"VCCFAEHVFIVRASE"}
            # Set the subcurrent for EHV_FIVRA only
            Iout_EHV=0
            
            gen.Generator1SVSC("VCCFAEHVFIVRANW", Iout_current/4, True)
            gen.Generator2SVSC("VCCFAEHVFIVRANE", Iout_current/4, True)        
            gen.Generator3SVSC("VCCFAEHVFIVRASW", Iout_current/4, True)
            gen.Generator4SVSC("VCCFAEHVFIVRASE", Iout_current/4, True)                    
            
            print("====")
            print(f'Iout={Iout_current}Amps')
            # Reset the max and min
            time.sleep(1)
            
            Measurment.ResetPersistentVoltageMeasurement(DRIVE_ONE_RAIL)

            time.sleep(1)
       
            Vout = Measurment.GetVoltageMeanOnce(DRIVE_ONE_RAIL)
            
            
            tt = Measurment.GetEPodOnce(DRIVE_ONE_RAIL)

            #get vout ripple
            raw_measurement = Measurment.GetPersistentVoltageMinMaxOnce(DRIVE_ONE_RAIL)
            phigh = float(raw_measurement.split()[1][:-1])
            plow = float(raw_measurement.split()[3][:-1])
            Vripple = float((phigh - plow)*1000) #change to mV
            time.sleep(0.2)
            
            ## Get PWR_IN(1Bh)
            PWR_IN=ifx.getsvidregvalue(svid_addr,svid_bus,0x1B)
            time.sleep(0.2)
            
            ## Get SVID_Iout(15h)
            SVID_Iout=ifx.getsvidregvalue(svid_addr,svid_bus,0x15)
            time.sleep(0.2)

            ## Get SVID_Iout(71h)
            SVID_Iout_71h=ifx.getsvidregvalue(svid_addr,svid_bus,0x71)

            ## set reg to Python list
            result_temp['Iout']=Iout_current
            result_temp['Vout']=float(Vout.replace("V",""))
            result_temp['Ripple']=float(Vripple)
            result_temp['Input Voltage']=0
            result_temp['Read PWR_IN']=PWR_IN
            result_temp['Read Iout 15h']=SVID_Iout
            result_temp['Read Iout 71h']=SVID_Iout_71h


            ##add list to pd
            results=results.append(result_temp,ignore_index=True)
            results=pd.DataFrame(results,columns=['Iout','Vout','Ripple','Input Voltage','Read PWR_IN','Read Iout 15h','Read Iout 71h','''"Iout_Value"'''])
            print(f"trun off loading and cooling {dtime}Sec") 
            
            time.sleep(dtime)
            
            time.sleep(dtime)
            gen.Generator1SVSC("VCCFAEHVFIVRANW",0, True)
            gen.Generator2SVSC("VCCFAEHVFIVRANE",0, True)
            gen.Generator3SVSC("VCCFAEHVFIVRASW",0, True)
            gen.Generator4SVSC("VCCFAEHVFIVRASE",0, True)

    elif rail_name=="PVCCD_HV":
        
        gen.AssignRailToDriveOne("PVCC_HV")
        gen.AssignRailToDriveTwo("PVCCDHV0123")
        gen.AssignRailToDriveThree("PVCCDHV4567")
        gen.SetVoltageForRail("PVCCD_HV", Vout, Transition.Fast)
        for Iout_current in Iout_currents:
            PVCCDHV_list=["PVCCHD0123","PVCCDHV4567"]
            #EHV_FIVRA_load_dict={"VCCFA_EHV_FIVRA_NW":"VCCFAEHVFIVRANW","VCCFA_EHV_FIVRA_NE":"VCCFAEHVFIVRANE","VCCFA_EHV_FIVRA_SW":"VCCFAEHVFIVRASW","VCCFA_EHV_FIVRA_SE":"VCCFAEHVFIVRASE"}
            # Set the subcurrent for EHV_FIVRA only
            Iout_EHV=0
            gen.Generator1SVSC("PVCCD_HV", 0, True)
            gen.Generator2SVSC("PVCCDHV0123", Iout_current/2, True)
            gen.Generator3SVSC("PVCCDHV4567", Iout_current/2, True)        
                   
            
            print("====")
            print(f'Iout={Iout_current}Amps')
            # Reset the max and min
            time.sleep(1)
            
            Measurment.ResetPersistentVoltageMeasurement(DRIVE_ONE_RAIL)

            time.sleep(1)
       
            Vout = Measurment.GetVoltageMeanOnce(DRIVE_ONE_RAIL)
            
            
            tt = Measurment.GetEPodOnce(DRIVE_ONE_RAIL)

            #get vout ripple
            raw_measurement = Measurment.GetPersistentVoltageMinMaxOnce(DRIVE_ONE_RAIL)
            phigh = float(raw_measurement.split()[1][:-1])
            plow = float(raw_measurement.split()[3][:-1])
            Vripple = float((phigh - plow)*1000) #change to mV
            time.sleep(0.2)
            
            ## Get PWR_IN(1Bh)
            PWR_IN=ifx.getsvidregvalue(svid_addr,svid_bus,0x1B)
            time.sleep(0.2)
            
            ## Get SVID_Iout(15h)
            SVID_Iout=ifx.getsvidregvalue(svid_addr,svid_bus,0x15)
            time.sleep(0.2)

            ## Get SVID_Iout(71h)
            SVID_Iout_71h=ifx.getsvidregvalue(svid_addr,svid_bus,0x71)

            ## set reg to Python list
            result_temp['Iout']=Iout_current
            result_temp['Vout']=float(Vout.replace("V",""))
            result_temp['Ripple']=float(Vripple)
            result_temp['Input Voltage']=0
            result_temp['Read PWR_IN']=PWR_IN
            result_temp['Read Iout 15h']=SVID_Iout
            result_temp['Read Iout 71h']=SVID_Iout_71h


            ##add list to pd
            results=results.append(result_temp,ignore_index=True)
            results=pd.DataFrame(results,columns=['Iout','Vout','Ripple','Input Voltage','Read PWR_IN','Read Iout 15h','Read Iout 71h','''"Iout_Value"'''])
            print(f"trun off loading and cooling {dtime}Sec") 
            
            time.sleep(dtime)
            
            time.sleep(dtime)
            gen.Generator2SVSC("PVCCDHV0123",0, True)
            gen.Generator3SVSC("PVCCDHV4567",0, True)

    else:
        for Iout_current in Iout_currents:
            
            # Set the current
            gen.Generator1SVSC(DRIVE_ONE_RAIL, Iout_current, True)
            print("====")
            print(f'Iout={Iout_current}Amps')
            # Reset the max and min
            time.sleep(1)
            Measurment.ResetPersistentCurrentMeasurement(DRIVE_ONE_RAIL)
            Measurment.ResetPersistentVoltageMeasurement(DRIVE_ONE_RAIL)

            time.sleep(1)
       
            Vout = Measurment.GetVoltageMeanOnce(DRIVE_ONE_RAIL)
            Iout = Measurment.GetCurrentMeanOnce(DRIVE_ONE_RAIL)
            
            tt = Measurment.GetEPodOnce(DRIVE_ONE_RAIL)

            #get vout ripple
            raw_measurement = Measurment.GetPersistentVoltageMinMaxOnce(DRIVE_ONE_RAIL)
            phigh = float(raw_measurement.split()[1][:-1])
            plow = float(raw_measurement.split()[3][:-1])
            Vripple = float((phigh - plow)*1000) #change to mV
            time.sleep(0.2)
            
            ## Get PWR_IN(1Bh)
            PWR_IN=ifx.getsvidregvalue(svid_addr,svid_bus,0x1B)
            time.sleep(0.2)
            
            ## Get SVID_Iout(15h)
            SVID_Iout=ifx.getsvidregvalue(svid_addr,svid_bus,0x15)
            time.sleep(0.2)

            ## Get SVID_Iout(71h)
            SVID_Iout_71h=ifx.getsvidregvalue(svid_addr,svid_bus,0x71)

            ## set reg to Python list
            result_temp['Iout']=float(Iout.replace("A",""))
            result_temp['Vout']=float(Vout.replace("V",""))
            result_temp['Ripple']=float(Vripple)
            result_temp['Input Voltage']=0
            result_temp['Read PWR_IN']=PWR_IN
            result_temp['Read Iout 15h']=SVID_Iout
            result_temp['Read Iout 71h']=SVID_Iout_71h


            ##add list to pd
            results=results.append(result_temp,ignore_index=True)
            results=pd.DataFrame(results,columns=['Iout','Vout','Ripple','Input Voltage','Read PWR_IN','Read Iout 15h','Read Iout 71h','''"Iout_Value"'''])
            print(f"trun off loading and cooling {dtime}Sec") 
            gen.Generator1SVSC(DRIVE_ONE_RAIL, 0, True)
            time.sleep(dtime)

         
    
    return results

def vr14_ifx_dc(rail_name="VCCIN",vout_list=[1.83,1.73],icc_max=100,cool_down_delay=1,LL_point=[1,3,5,7],excel=True):
        print(f"dll verison = {ifx.version()}")

        for vout in vout_list:
            svid_reg_dict={'Pin_Max':0x2E,
                   'Pin_Max_Add':0x51,
                   'Icc_Max':0x21,
                   'Icc_Max_Add':0x50,
                   'DC_LL':0x23,
                   'DC_LL_Fine':0x36,
                   'Capability':0x6,
                   'Protocol_id':0x5,
                   'ext_capability':0x9,
                   'VR14_capability':0x50,
                   'VIDo_max':0x0A}
            svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(rail_name)    
            svid_reg_value=ifx.get_all_svid_reg(svid_addr,svid_bus,svid_reg_dict)
            Iout=list() 

        
            theTime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename="IFX-VR-"+rail_name+"-"+str(vout)+"V_"+theTime+".xlsx"
            dtime=cool_down_delay

            Iout=LL_point
            df2=dcLoadlinetest(rail_name,svid_bus,svid_addr,vout,Iout,dtime,svid_reg_value)
            svid_reg_df=pd.DataFrame.from_dict(svid_reg_dict,orient='index',columns=['svid_command_code'])
            svid_reg_df=svid_reg_df.reset_index()
            svid_reg_df['svid_command_code']=svid_reg_df['svid_command_code'].apply(int).apply(hex).apply(str)
            
            df1=pd.DataFrame.from_dict(svid_reg_value,orient='index',columns=['svid_reg_value'])
            df1["svid_reg_value"]=df1["svid_reg_value"].apply(int).apply(hex).apply(str)
            #df1=df1.drop(['Value'],axis=1)
            df1=df1.reset_index()
            
            df3=pd.concat([svid_reg_df,df1],axis=1)
            #df1.rename(columns={'index':'command','Hex':'abc'})
            
            ifx.df1_df2_to_excel(df3,df2,filename,"Reg value","LoadLine")

            ifx.excelSelect(filename,excel)

        print("completed")

if __name__ == "__main__":


     vr14_ifx_dc("PVCCD_HV",[1.1],136,2,[0,1.3,3,5,8,10,13,16,18,21,23,26],excel=True)
    
    




    

    


