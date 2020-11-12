import sys
import pythoncom
pythoncom.CoInitialize() # to fix issue between clr and QFileDialog
from PyQt5.QtWidgets import QMainWindow,QApplication,QFileDialog
from PyQt5.QtCore import QThread,pyqtSignal

import Ui_PyQT_IFX_VR14
import ifx
import time

if ifx.detect_process("Gen5.exe"):
    pass
else:
    print("please run Gen5 first")
    time.sleep(15)
    sys.exit()

import IFX_VR14_DC_LL
import IFX_VR14_3D
import IFX_set_wp
import json

class myprogpressbar(QThread):
    updatebar=pyqtSignal(int)
    def __init__(self):
        QThread.__init__(self)
    def __del__(self):
        self.wait()
    def run(self):
        while True:
            for index in range(1, 101):
                self.updatebar.emit(index)
                time.sleep(0.1)
    def stop(self):
        self.terminate()

class dc_thread(QThread):
    dc_to_enable_abort = pyqtSignal(bool)    
    def __init__(self):
        QThread.__init__(self)
        #self.dc_to_enable_abort = pyqtSignal(bool)
        
    def __del__(self):
        self.wait()
    def run(self):
        self.dc_to_enable_abort.emit(True)
        IFX_VR14_DC_LL.vr14_ifx_dc(myWin.rail_name,myWin.vout_list,1,myWin.cool_down_delay,myWin.iout_list,myWin.excel, myWin.hc_bit)

    def stop(self):
        self.terminate()

class vr3d_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)
    def __init__(self):
        dc_thread.__init__(self)
    def __del__(self):
        self.wait()
    def run(self):
        IFX_VR14_3D.vr14_3d(myWin.rail_name,myWin.vout_list_3d,myWin.start_freq,myWin.end_freq,myWin.icc_min,myWin.icc_max,myWin.freq_steps_per_decade,myWin.rise_time,myWin.cool_down_delay_3d,myWin.start_duty,myWin.end_duty,myWin.duty_step,myWin.excel)
        
    def stop(self):
        self.terminate()        

class dc_thermal_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)
    def __init__(self):
        dc_thread.__init__(self)
    def __del__(self):
        self.wait()
    def run(self):
        IFX_VR14_DC_LL.thermal_drop(myWin.rail_name,myWin.thermal_vout_list, myWin.thermal_iout, myWin.thermal_time_list)
        
    def stop(self):
        self.terminate()


class MyMainWindow(QMainWindow,Ui_PyQT_IFX_VR14.Ui_MainWindow):
    def __init__(self,parent=None):
        super(MyMainWindow,self).__init__(parent)
        self.setFixedSize(840, 940)
        self.setupUi(self)
        
        # reflash power rails to all combobox.
        self.current_power_rail_dict=ifx.detect_all_rails()
        self.reflash_power_rails()
        self.setwp_change_vid_selection_function()
        
        self.pushButton_8.clicked.connect(self.dc_thermal_function)
        self.pushButton_2.clicked.connect(self.dc_loadline_function)
        self.pushButton_7.clicked.connect(self.stop_all_thread)
        self.pushButton.clicked.connect(self.vr3d_function)
        self.pushButton_9.clicked.connect(self.check_vendor_id)
        #self.pushButton_5.clicked.connect(self.save_parameter_function)
        self.actionSave_Parameter_to_file.triggered.connect(self.save_parameter_function)
        #self.pushButton_6.clicked.connect(self.load_parameter_function)
        self.actionLoad_Parameter_to_GUI.triggered.connect(self.load_parameter_function)
        self.pushButton_4.clicked.connect(self.load_report_to_3d_function)
        self.pushButton_13.clicked.connect(self.set_wp_vector_function)
        self.pushButton_15.clicked.connect(self.clear_svid_alert_function)
                
        #set vector 
        self.comboBox_11.currentIndexChanged.connect(self.set_vector_line1)
        self.comboBox_13.currentIndexChanged.connect(self.set_vector_line2)
        self.comboBox_15.currentIndexChanged.connect(self.set_vector_line3)
        self.comboBox_17.currentIndexChanged.connect(self.set_vector_line4)
        self.comboBox.currentIndexChanged.connect(self.change_VID_selection_function)
        self.comboBox.highlighted.connect(self.change_VID_selection_function)
        self.comboBox_2.currentIndexChanged.connect(self.setwp_rail1_vid_combox_function)
        self.comboBox_3.currentIndexChanged.connect(self.setwp_rail2_vid_combox_function)
        self.comboBox_4.currentIndexChanged.connect(self.setwp_rail3_vid_combox_function)
        self.comboBox_5.currentIndexChanged.connect(self.setwp_rail4_vid_combox_function)
        self.comboBox_10.currentIndexChanged.connect(self.setwp_change_vid_selection_function)
        self.comboBox_12.currentIndexChanged.connect(self.setwp_change_vid_selection_function)
        self.comboBox_14.currentIndexChanged.connect(self.setwp_change_vid_selection_function)
        self.comboBox_16.currentIndexChanged.connect(self.setwp_change_vid_selection_function)        
        #self.comboBox_24.currentIndexChanged.connect(self.turn_off_check_button_function)
        
        #set command list in comboBox of Vector
        command_list=['setVIDfast', 'setVIDslow', 'setVIDdecay', 'setPS', 'setRegaddr', 'setREGdata', 'getREG']
        self.comboBox_11.addItems(command_list)
        self.comboBox_13.addItems(command_list)
        self.comboBox_15.addItems(command_list)
        self.comboBox_17.addItems(command_list)   
        


        self.pushButton_10.clicked.connect(self.set_wp_init_function)

        self.comboBox.activated.connect(self.qbox)
        self.lcdNumber.setHexMode()
        self.lcdNumber.display(0xFF)
        #self.pushButton_5.setEnabled(True)
        #self.pushButton_6.setEnabled(True)
        self.pushButton_4.setEnabled(True)

        self.tabWidget.setTabEnabled(1,True)
        self.tabWidget.setTabEnabled(2,False)

        #set SetWP rail combobox disable
        self.comboBox_2.setEnabled(True)
        self.comboBox_2.addItems(list(self.current_power_rail_dict.keys()))
        self.comboBox_3.setEnabled(True)
        self.comboBox_3.addItems(list(self.current_power_rail_dict.keys()))
        self.comboBox_4.setEnabled(True)
        self.comboBox_4.addItems(list(self.current_power_rail_dict.keys()))
        self.comboBox_5.setEnabled(True)
        self.comboBox_5.addItems(list(self.current_power_rail_dict.keys()))        
        ## init thread
        self.dc_loadline_thread=dc_thread()
        self.vr3d=vr3d_thread()
        self.dc_thermal_thread=dc_thermal_thread()
        self.myprogpressbar=myprogpressbar()

        ## set signal connection
        self.myprogpressbar.updatebar.connect(self.update_bar)
        
        self.start_up()
        
    def setwp_rail1_vid_combox_function(self):
        temp=self.current_power_rail_dict[self.comboBox_2.currentText()]
        self.comboBox_6.clear()
        self.comboBox_6.addItem(str(temp[2]))
    def setwp_rail2_vid_combox_function(self):
        temp=self.current_power_rail_dict[self.comboBox_3.currentText()]
        self.comboBox_7.clear()
        self.comboBox_7.addItem(str(temp[2]))        
    def setwp_rail3_vid_combox_function(self):
        temp=self.current_power_rail_dict[self.comboBox_4.currentText()]
        self.comboBox_8.clear()
        self.comboBox_8.addItem(str(temp[2]))
    def setwp_rail4_vid_combox_function(self):
        temp=self.current_power_rail_dict[self.comboBox_5.currentText()]
        self.comboBox_9.clear()
        self.comboBox_9.addItem(str(temp[2]))          
    def change_VID_selection_function(self):
        self.comboBox_24.clear()
        self.current_power_rail_dict=ifx.detect_all_rails()
        print(self.current_power_rail_dict)
        temp=self.current_power_rail_dict[self.comboBox.currentText()]
        print(f" powerrail {temp[2]}")
        self.comboBox_24.addItem(str(temp[2]))
        if self.comboBox_24.currentText()=='0':
            self.pushButton_9.setEnabled(False)
        else:
            self.pushButton_9.setEnabled(True)
        
    def reflash_power_rails(self):
        print("reflash")
        
        self.comboBox.clear()
        self.comboBox.addItems(list(self.current_power_rail_dict.keys()))
        
        
        
        item_length=self.comboBox.count()
        item_list=list()
        for i in range(0,item_length):
            item_list.append(self.comboBox.itemText(i)) 
        self.comboBox_10.clear()
        self.comboBox_12.clear()    
        self.comboBox_14.clear()
        self.comboBox_16.clear()  
        self.comboBox_10.addItems(item_list)
        self.comboBox_12.addItems(item_list)
        self.comboBox_14.addItems(item_list)
        self.comboBox_16.addItems(item_list)
        
        self.qbox()
        
    def clear_svid_alert_function(self):
        IFX_set_wp.ResetAllVectors()
        for SVID_addr in range(0, 4):
            IFX_set_wp.add_vector_into_vector(index=SVID_addr,SVID_address=SVID_addr,SVID_command=7,SVID_payload=0x10,SVID_delay=200)
            
        IFX_set_wp.run_vector()
        
    def setwp_change_vid_selection_function(self):
        self.comboBox_23.clear()
        temp=self.current_power_rail_dict[self.comboBox_10.currentText()]
        self.comboBox_23.addItem(str(temp[2]))

        self.comboBox_25.clear()
        temp=self.current_power_rail_dict[self.comboBox_12.currentText()]
        self.comboBox_25.addItem(str(temp[2]))

        self.comboBox_26.clear()
        temp=self.current_power_rail_dict[self.comboBox_14.currentText()]
        self.comboBox_26.addItem(str(temp[2])) 

        self.comboBox_27.clear()
        temp=self.current_power_rail_dict[self.comboBox_16.currentText()]
        self.comboBox_27.addItem(str(temp[2]))         

    def set_wp_vector_function(self):
        vector_index=0
        IFX_set_wp.ResetAllVectors()

        #vector item_1        
        if self.checkBox.isChecked()==1:
            temp=ifx.VID_10mV_to_hex(float(self.lineEdit_60.text())*1000)
            item_1_payload_data=int(temp,16)
            item_1_delay=int(self.lineEdit_81.text())
            item_1_svid_addr, item_1_svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_10.currentText())
            if self.comboBox_11.currentIndex()<=1:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_1_svid_addr,SVID_command=(self.comboBox_11.currentIndex()+1),SVID_payload=item_1_payload_data,SVID_delay=item_1_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_1_svid_addr,SVID_command=(self.comboBox_11.currentIndex()+1),SVID_payload=int(self.lineEdit_59.text(),16),SVID_delay=item_1_delay)
            vector_index+=1
        #vector item_2
        if self.checkBox_7.isChecked()==1:
            temp=ifx.VID_10mV_to_hex(float(self.lineEdit_62.text())*1000)
            item_2_payload_data=int(temp,16)
            item_2_delay=int(self.lineEdit_86.text())
            item_2_svid_addr, item_2_svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_12.currentText())
            if self.comboBox_13.currentIndex()<=1:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_2_svid_addr,SVID_command=(self.comboBox_13.currentIndex()+1),SVID_payload=item_2_payload_data,SVID_delay=item_2_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_2_svid_addr,SVID_command=(self.comboBox_13.currentIndex()+1),SVID_payload=int(self.lineEdit_61.text(),16),SVID_delay=item_2_delay)
            vector_index+=1
            
        #vector item_3
        if self.checkBox_8.isChecked()==1:
            temp=ifx.VID_10mV_to_hex(float(self.lineEdit_64.text())*1000)
            item_3_payload_data=int(temp,16)
            item_3_delay=int(self.lineEdit_85.text())
            item_3_svid_addr, item_3_svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_14.currentText())
            if self.comboBox_15.currentIndex()<=1:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_3_svid_addr,SVID_command=(self.comboBox_15.currentIndex()+1),SVID_payload=item_3_payload_data,SVID_delay=item_3_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_3_svid_addr,SVID_command=(self.comboBox_15.currentIndex()+1),SVID_payload=int(self.lineEdit_63.text(),16),SVID_delay=item_3_delay)
            vector_index+=1
            
        if self.checkBox_9.isChecked()==1:
            temp=ifx.VID_10mV_to_hex(float(self.lineEdit_66.text())*1000)
            item_4_payload_data=int(temp,16)
            item_4_delay=int(self.lineEdit_83.text())
            item_4_svid_addr, item_4_svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_16.currentText())
            if self.comboBox_17.currentIndex()<=1:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_4_svid_addr,SVID_command=(self.comboBox_17.currentIndex()+1),SVID_payload=item_4_payload_data,SVID_delay=item_4_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index,SVID_address=item_4_svid_addr,SVID_command=(self.comboBox_17.currentIndex()+1),SVID_payload=int(self.lineEdit_65.text(),16),SVID_delay=item_4_delay)
            vector_index+=1
            
        if vector_index ==0:
            pass
            #IFX_set_wp.ResetAllVectors()
         
        if self.comboBox_23.currentText()!="0" and self.comboBox_25.currentText()!="0":
            IFX_set_wp.run_vector()
        else:
            print("VID table is 0")
            #IFX_set_wp.ResetAllVectors()

    def set_vector_line1(self, i):
        if i<=1:
            self.lineEdit_59.setEnabled(False)
            self.lineEdit_60.setEnabled(True)
        else:
            self.lineEdit_59.setEnabled(True)
            self.lineEdit_60.setEnabled(False)
            
    def set_vector_line2(self, i):
        if i<=1:
            self.lineEdit_61.setEnabled(False)
            self.lineEdit_62.setEnabled(True)
        else:
            self.lineEdit_61.setEnabled(True)
            self.lineEdit_62.setEnabled(False)

    def set_vector_line3(self, i):
        if i<=1:
            self.lineEdit_63.setEnabled(False)
            self.lineEdit_64.setEnabled(True)
        else:
            self.lineEdit_63.setEnabled(True)
            self.lineEdit_64.setEnabled(False)
            
    def set_vector_line4(self, i):
        if i<=1:
            self.lineEdit_65.setEnabled(False)
            self.lineEdit_66.setEnabled(True)
        else:
            self.lineEdit_65.setEnabled(True)
            self.lineEdit_66.setEnabled(False)
        
    def set_wp_init_function(self):
        IFX_set_wp.ResetAllVectors()
        # rail_1
        wp1=ifx.VID_10mV_to_hex(float(self.lineEdit_20.text())*1000,self.comboBox_6.currentText() )
        print(self.comboBox_6.currentText())
        rail_1_wp1_new=int(wp1,16)
        wp2=ifx.VID_10mV_to_hex(float(self.lineEdit_21.text())*1000,self.comboBox_6.currentText())
        rail_1_wp2_new=int(wp2,16)

        # rail_2
        wp1=ifx.VID_10mV_to_hex(float(self.lineEdit_23.text())*1000,self.comboBox_7.currentText())
        rail_2_wp1_new=int(wp1,16)
        wp2=ifx.VID_10mV_to_hex(float(self.lineEdit_24.text())*1000,self.comboBox_7.currentText())
        rail_2_wp2_new=int(wp2,16)

        # rail_3
        wp1=ifx.VID_10mV_to_hex(float(self.lineEdit_26.text())*1000,self.comboBox_8.currentText())
        rail_3_wp1_new=int(wp1,16)
        wp2=ifx.VID_10mV_to_hex(float(self.lineEdit_27.text())*1000,self.comboBox_8.currentText())
        rail_3_wp2_new=int(wp2,16)

        # rail_4
        wp1=ifx.VID_10mV_to_hex(float(self.lineEdit_29.text())*1000,self.comboBox_9.currentText())
        rail_4_wp1_new=int(wp1,16)
        wp2=ifx.VID_10mV_to_hex(float(self.lineEdit_30.text())*1000,self.comboBox_9.currentText())
        rail_4_wp2_new=int(wp2,16)

        #set rail_1 wp1
        rail1_svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_2.currentText())
        
        IFX_set_wp.add_vector_into_vector(index=0,SVID_address=rail1_svid_addr,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=1,SVID_address=rail1_svid_addr,SVID_command=6,SVID_payload=rail_1_wp1_new,SVID_delay=400)
        #set rail_1 wp2      
        IFX_set_wp.add_vector_into_vector(index=2,SVID_address=rail1_svid_addr,SVID_command=5,SVID_payload=0x3C,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=3,SVID_address=rail1_svid_addr,SVID_command=6,SVID_payload=rail_1_wp2_new,SVID_delay=400)

        #set rail_2 wp1
        rail2_svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_3.currentText())
        IFX_set_wp.add_vector_into_vector(index=4,SVID_address=rail2_svid_addr,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=5,SVID_address=rail2_svid_addr,SVID_command=6,SVID_payload=rail_2_wp1_new,SVID_delay=400)
        #set rail_2 wp2      
        IFX_set_wp.add_vector_into_vector(index=6,SVID_address=rail2_svid_addr,SVID_command=5,SVID_payload=0x3C,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=7,SVID_address=rail2_svid_addr,SVID_command=6,SVID_payload=rail_2_wp2_new,SVID_delay=400)

        #set rail_3 wp1
        rail3_svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_4.currentText())       
        IFX_set_wp.add_vector_into_vector(index=8,SVID_address=rail3_svid_addr,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=9,SVID_address=rail3_svid_addr,SVID_command=6,SVID_payload=rail_3_wp1_new,SVID_delay=400)
        #set rail_3 wp2      
        IFX_set_wp.add_vector_into_vector(index=10,SVID_address=rail3_svid_addr,SVID_command=5,SVID_payload=0x3C,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=11,SVID_address=rail3_svid_addr,SVID_command=6,SVID_payload=rail_3_wp2_new,SVID_delay=400)

        #set rail_4 wp1
        rail4_svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox_5.currentText())
        IFX_set_wp.add_vector_into_vector(index=12,SVID_address=rail4_svid_addr,SVID_command=5,SVID_payload=0x3B,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=13,SVID_address=rail4_svid_addr,SVID_command=6,SVID_payload=rail_4_wp1_new,SVID_delay=400)
        #set rail_4 wp2      
        IFX_set_wp.add_vector_into_vector(index=14,SVID_address=rail4_svid_addr,SVID_command=5,SVID_payload=0x3C,SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(index=15,SVID_address=rail4_svid_addr,SVID_command=6,SVID_payload=rail_4_wp2_new,SVID_delay=400)

        IFX_set_wp.run_vector()
#        time.sleep(0.5)
#        IFX_set_wp.ResetAllVectors()
    def load_report_to_3d_function(self):
        dlg=QFileDialog(self, 'Open File', '.', 'Excel report (*.xlsx);;All Files (*)')
        if dlg.exec_():
            filenames=dlg.selectedFiles()
            with open(filenames[0],'r') as j:
                ifx.plt_vmax(filenames[0])
        
    def update_bar(self,data):
        self.progressBar.setValue(data)
        
    def save_parameter_function(self):
        self.update_GUI()
        if self.tabWidget.currentIndex() == 0:
            self.parameter_dict={"rail_name":self.comboBox.currentText(),
                                 "LL_vout_list":self.lineEdit_10.text(),
                                 "LL_iout_list":self.lineEdit_11.text(),
                                 "LL_cool_down_delay":self.lineEdit_14.text(),
                                 
                                 "3d_vout_list":self.lineEdit_16.text(),
                                 "3d_cool_down_delay":self.lineEdit_15.text(),
                                 "3d_start_freq":self.lineEdit.text(),
                                 "3d_end_freq":self.lineEdit_2.text(),
                                 "3d_rise_time":self.lineEdit_6.text(),
                                 "3d_fall_time":self.lineEdit_12.text(),
                                 "3d_icc_min":self.lineEdit_4.text(),
                                 "3d_icc_max":self.lineEdit_3.text(),
                                 "3d_freq_steps_per_decade":self.lineEdit_5.text(),
                                 "3d_duty_step":self.lineEdit_9.text(),
                                 "3d_start_duty":self.lineEdit_8.text(),
                                 "3d_end_duty":self.lineEdit_7.text(),
                                 
                                 }
    
            filename_with_path=QFileDialog.getSaveFileName(self, 'Save File', '.', 'JSON Files (*.json)')
            save_filename=filename_with_path[0]
            if save_filename!="":
                with open(save_filename, 'w') as fp:            
                    json.dump(self.parameter_dict, fp)
        else:
            print(self.tabWidget.currentIndex)
            print("only can save first tab now")
    def load_parameter_function(self):
        
        dlg=QFileDialog(self, 'Open File', '.', 'JSON Files (*.json);;All Files (*)')
        if self.tabWidget.currentIndex() == 0:
            if dlg.exec_():
                filenames=dlg.selectedFiles()
    
                with open(filenames[0],'r') as j:
                    json_data=json.load(j)
                    print(json_data)
                self.comboBox.setCurrentText(json_data["rail_name"])
                self.lineEdit_10.setText(json_data["LL_vout_list"])
                self.lineEdit_11.setText(json_data["LL_iout_list"])
                self.lineEdit_14.setText(json_data["LL_cool_down_delay"])
                
                self.lineEdit_16.setText(json_data["3d_vout_list"])
                self.lineEdit_15.setText(json_data["3d_cool_down_delay"])
                self.lineEdit.setText(json_data["3d_start_freq"])
                self.lineEdit_2.setText(json_data["3d_end_freq"])
                self.lineEdit_6.setText(json_data["3d_rise_time"])
                self.lineEdit_12.setText(json_data["3d_fall_time"])
                self.lineEdit_4.setText(json_data["3d_icc_min"])
                self.lineEdit_3.setText(json_data["3d_icc_max"])
                self.lineEdit_5.setText(json_data["3d_freq_steps_per_decade"])
                self.lineEdit_9.setText(json_data["3d_duty_step"])
                self.lineEdit_8.setText(json_data["3d_start_duty"])
                self.lineEdit_7.setText(json_data["3d_end_duty"])
        else:
            print("only support first page")
            
    
    def update_GUI(self):
        ## get GUI import
        self.rail_name=self.comboBox.currentText()
        self.vout_list=eval("["+str(self.lineEdit_10.text())+"]")
        self.iout_list=eval("["+str(self.lineEdit_11.text())+"]")
        self.cool_down_delay=int(self.lineEdit_14.text())
        self.excel=self.checkBox_2.isChecked()
        self.hc_bit=self.checkBox_10.isChecked()

        ##get GUI import for 3D parameter      
        self.vout_list_3d=eval("["+str(self.lineEdit_16.text())+"]")
        self.start_freq=float(self.lineEdit.text())
        self.end_freq=int(self.lineEdit_2.text())
        self.rise_time=int(self.lineEdit_6.text())
        self.fall_time=int(self.lineEdit_12.text())
        self.icc_min=int(self.lineEdit_4.text())
        self.icc_max=int(self.lineEdit_3.text())
        self.freq_steps_per_decade=int(self.lineEdit_5.text())
        self.duty_step=int(self.lineEdit_9.text())
        self.start_duty=int(self.lineEdit_8.text())
        self.end_duty=int(self.lineEdit_7.text())
        self.cool_down_delay_3d=int(self.lineEdit_15.text())

        ## get GUI import from thermal drop
        self.thermal_vout_list=eval("["+str(self.lineEdit_17.text())+"]")
        self.thermal_time_list=eval("["+str(self.lineEdit_19.text())+"]")
        self.thermal_iout=float(self.lineEdit_18.text())
        
    def vr3d_function(self):
        self.update_GUI()
        self.runing_items()
        self.vr3d.start()
        self.myprogpressbar.start()

    def dc_thermal_function(self):
        self.update_GUI()
        self.runing_items()
        self.dc_thermal_thread.start()
        self.myprogpressbar.start()
        
    def dc_loadline_function(self):
        self.update_GUI()
        self.runing_items()        
        self.dc_loadline_thread.start()
        self.myprogpressbar.start()

    def stop_all_thread(self):
        print("====")
        print("abort test")
        self.start_up()
        self.dc_loadline_thread.stop()
        self.vr3d.stop()
        self.dc_thermal_thread.stop()
        self.myprogpressbar.stop()

    def check_vendor_id(self):
        svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox.currentText())
        self.vendor_id=ifx.get_svid_reg(svid_addr,svid_bus,0)
        if self.vendor_id == '13' :
            self.pushButton_8.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton.setEnabled(True)
            self.tabWidget.setTabEnabled(1,True)

        else:
            self.start_up()
            self.pushButton_9.setEnabled(True)
        self.lcdNumber.display(self.vendor_id)

    def qbox(self):
        
        #self.pushButton_9.setEnabled(True)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(True)
        self.tabWidget.setTabEnabled(1,False)


        
    def start_up(self):
        self.comboBox.setEnabled(True)
        self.pushButton_9.setEnabled(False)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(False)
        self.tabWidget.setTabEnabled(1,False)
        self.tabWidget.setTabEnabled(2,False)
        self.lcdNumber.display(0xFF)
        
    def runing_items(self):
        self.comboBox.setEnabled(False)
        self.pushButton_9.setEnabled(False)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(True)
        self.lcdNumber.display(0x00)

if __name__ == "__main__":
    app=QApplication(sys.argv)

    myWin=MyMainWindow()

    myWin.show()

    sys.exit(app.exec_())
