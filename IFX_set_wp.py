# ---------------------------- GEN5 VRTT - SVID VECTORS ------------------------------
# ------------------------------------------------------------------------------------
#   Intel Confidential 
# ------------------------------------------------------------------------------------
# - This script executes SVID vectors creating a voltage staircase 
# - SVID command-delay pairs can be changed by altering arguments in "vector.CreateVector()" 
# - Response of each SVID transaction is displayed on python shell
# ------------------------------------------------------------------------------------
#   Rev 1.0   - 5/15/2018
#           - Initial versioning.
#   pulled latest from GEN5CONTROLLER\NEXTGEN\PythonApplication2

import clr
import sys
import time

dll_path=r"C:\GEN5\PythonScripts"
sys.path.append(dll_path)

from Common.APIs import DataAPI
from Common.APIs import VectorAPI
from Common.APIs import MainAPI
from Common.APIs import DisplayAPI
from Common.Models import TriggerModel
from Common.Models import FanModel
from Common.Models import SVIDModel
from Common.Models import RawDataModel
from Common.Models import DataModel
from Common.Models import DisplayModel
from Common.Enumerations import HorizontalScale
from Common.Enumerations import ScoplessChannel
from Common.Enumerations import Cursors
from Common.APIs import MeasurementAPI
from Common.APIs import GeneratorAPI
from Common.Enumerations import PowerState
from Common.Enumerations import Transition
from Common.Enumerations import Protocol
from Common.Enumerations import RailTestMode
from Common.Enumerations import FRAME
from Common.Enumerations import SVIDBURSTTRIGGER
from Common.Enumerations import ScoplessChannel
from System import String, Char, Int32, UInt16, Boolean, Array, Byte, Double
from System.Collections.Generic import List

gen = GeneratorAPI()
display = DisplayAPI()
data = DataAPI()
Meas = MeasurementAPI()
vector = VectorAPI()
api = MainAPI()

def Load_All_Vectors_function(repeat=0, svid_bus=0):
    vector.LoadAllVectors(repeat, svid_bus)

def PrintVector(index,SVID_address,SVID_command,SVID_payload,SVID_delay,Vector_active):
    if Vector_active == 1:
        print(index+1,'\t', SVID_address,'\t', SVID_command,'\t', hex(SVID_payload),'\t', SVID_delay)
    else:
        index = 0
    return index

# Reset all vectors
def ResetAllVectors():
    vector.ResetVectors()
    

def add_vector_into_vector(index,SVID_address,SVID_command,SVID_payload,SVID_delay,vector_active=1):
    vector.CreateVector(index,2,SVID_address,SVID_command,SVID_payload,2,3,int(SVID_delay/40),vector_active)
    print(index+1,'\t', SVID_address,'\t', SVID_command,'\t', hex(SVID_payload),'\t', SVID_delay)

def run_vector():
    vector.LoadAllVectors(0,1)
    time.sleep(0.5)
    vector.ExecuteVectors(0)
    time.sleep(0.5)
    ResetAllVectors()

def SetPreemptiveVector():
    ON=1
    OFF=0
    print('Vectors set as : ')
    print('-------------------------------------------------------------------')
    print('# ', 'ADDRESS ', 'COMMAND ', 'DATA ', 'DELAY,ns')
    
#create vector #1
    add_vector_into_vector(index=0,SVID_address=0,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
#create vector #2
    add_vector_into_vector(index=1,SVID_address=0,SVID_command=6,SVID_payload=0x86,SVID_delay=400)
#create vector #3
    add_vector_into_vector(index=2,SVID_address=1,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
#create vector #4
    add_vector_into_vector(index=3,SVID_address=1,SVID_command=6,SVID_payload=0xAB,SVID_delay=400)
#create vector #5
    add_vector_into_vector(index=4,SVID_address=2,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
#create vector #6
    add_vector_into_vector(index=5,SVID_address=2,SVID_command=6,SVID_payload=0x97,SVID_delay=400)
#create vector #7
    add_vector_into_vector(index=6,SVID_address=0,SVID_command=5,SVID_payload=0x3C,SVID_delay=400)
#create vector #8
    add_vector_into_vector(index=7,SVID_address=0,SVID_command=6,SVID_payload=0x72,SVID_delay=400)
#create vector #9
    index = 8
    length = 0
    SVID_address = 1  #VccSA
    SVID_command = 5 # 1=setVIDfast, 2=setVIDslow, 3=setVIDdecay, 4=setPS, 5=setREGaddr, 6=setREGdata, 7=getREGdata
    SVID_payload = 0x3C #SetRegAddr
    SVID_delay = 400 #ns
    Vector_active = ON
    vector.CreateVector(index,2,SVID_address,SVID_command,SVID_payload,2,3,int(SVID_delay/40),Vector_active)
    if Vector_active == 1:
        length = PrintVector(index,SVID_address,SVID_command,SVID_payload,SVID_delay,Vector_active)
#create vector #10
    index = 9
    SVID_address = 1  #VccSA
    SVID_command = 6 #setRegData
    SVID_payload = 0x79 #0.85V
    SVID_delay = 400 #ns
    Vector_active = ON
    vector.CreateVector(index,2,SVID_address,SVID_command,SVID_payload,2,3,int(SVID_delay/40),Vector_active)
    if Vector_active == 1:
        length = PrintVector(index,SVID_address,SVID_command,SVID_payload,SVID_delay,Vector_active)
#create vector #11
    index = index+1
    length = 0
    SVID_address = 2 #VccIO
    SVID_command = 5 # 1=setVIDfast, 2=setVIDslow, 3=setVIDdecay, 4=setPS, 5=setREGaddr, 6=setREGdata, 7=getREGdata
    SVID_payload = 0x3C #SetRegAddr
    SVID_delay = 400 #ns
    Vector_active = ON
    vector.CreateVector(index,2,SVID_address,SVID_command,SVID_payload,2,3,int(SVID_delay/40),Vector_active)
    if Vector_active == 1:
        length = PrintVector(index,SVID_address,SVID_command,SVID_payload,SVID_delay,Vector_active)
#create vector #12
    index = index+1
    SVID_address = 2  #VccIO
    SVID_command = 6 #setRegData
    SVID_payload = 0x8D #0.95V
    SVID_delay = 400 #ns
    Vector_active = ON
    vector.CreateVector(index,2,SVID_address,SVID_command,SVID_payload,2,3,int(SVID_delay/40),Vector_active)
    if Vector_active == 1:
        length = PrintVector(index,SVID_address,SVID_command,SVID_payload,SVID_delay,Vector_active)

    vector.LoadAllVectors(0,1)
    time.sleep(0.5)
    vector.ExecuteVectors(0)
    vector.ReadVectorResponses()
    vResult = data.GetVectorDataOnce()
    print('-------------------------------------------------------------------')
    print('SVID RESPONSE')
    print('-------------------------------------------------------------------')
    print('#\tACK\tDATA\tPARITY')
    print('-------------------------------------------------------------------')
    for i in range (0,length+1):
       print(i+1,'\t',str(vResult.AckResp[i]),'\t',hex(vResult.DataResp[i]),'\t',str(vResult.ParityValResp[i]))
    return vResult.DataResp


def set_WP_init(rail1,rail2):
    ########## End of imports
    print("##############################################################################")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    print("                              SVID VECTORS                                    ")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    print("##############################################################################")
    print("")


    ## What Rails are available?
    print ("Displaying rails available");
    print ("-----------------------------------------------------------------------------------------------------------------------------"); 
    t=gen.GetRails()
    for i in range (0, len(t)):
        print(t[i].Name, "Vid Table - ", t[i].VID, " Max Current - ", t[i].MaxCurrent, "; Vsense - ", t[i].VSenseName)
        print("Load sections: " + t[i].LoadSections)
        for k in t[i].ProtoParm.Keys:
                    print("      ", k, " ", t[i].ProtoParm[k])
    print ("");
    print (""); 
    print ("Test Data");
    print ('-------------------------------------------------------------------');


    #####Assign rails to Tabs
    gen.AssignRailToDriveOne(rail1)    
    gen.AssignRailToDriveTwo(rail2)

    # Display channel
    display.Ch1_2Rail(rail1)

    # Display digital signals
    display.DisplayDigitalSignal("SVID1ALERT")
    display.DisplayDigitalSignal("CSO1#")

    ####Enable Scope Traces
    display.SetChannel(ScoplessChannel.Ch1, True)
    display.SetChannel(ScoplessChannel.Ch2, True)

    f = FanModel()
    f.IsSynchronized = True

    svid = SVIDModel()
    svid.IsSynchronized = True

    t = TriggerModel()
    t.IsSynchronized = True

    # Setting scaling
    hscale = HorizontalScale.Scale5ms;

    cursor = Cursors.Ch3Cr1;



    # Create vectors, define delay between vectors, load into FPGA and execute once.
    # Arguments in order -
    # Index, FrameStart, VR Addr, SVID Cmd, SVID Payload, Parity (0-odd, 1-even, 2-Auto), FrameEnd, Delay, Status (1-active, 0-Inactive)


        


            
    lc_list = []
    lc_list.append("LC4")
    lc_list.append("LC5")


    ##Change Fan Speed
    data.SetFanSpeed(10)    

    #####Assign rails to Tabs
    gen.AssignRailToDriveOne(rail1)    
    gen.AssignRailToDriveTwo(rail2)

    #Set Scale
    display.SetHorizontalScale(HorizontalScale.Scale200us)
    display.SetVerticalCurrentScale(0, 330)
    display.SetVerticalVoltageScale(0, 2)

    ## Set trigger to "SVID Burst"
    data.SetTrigger(12, 0, 0, 0, 0, False)

    # Main loop 
    # This command runs vectors once 
    #SetVectorsStairCase();
    #SetPreemptiveVector();

    # Repeat this command 20 times
    #for i in range(0,20):
    #    RepeatVectorsStaircase()

    # Reset all vectors
    ResetAllVectors();
    print ('-------------------------------------------------------------------')
    print ("")
    print ("Test complete");

if __name__ =="__main__":
    set_WP_init("VCCIN","PVCCFA_EHV")
