import sys
from PyQt5.QtWidgets import QMainWindow,QApplication#,QTableView
from PyQt5.QtCore import QThread,pyqtSignal#,QTime
import Ui_PyQT_IFX_VR14
import ifx
import time

if ifx.detect_process("Gen5.exe"):
    pass
else:
    print("please run Gen5 frist")
    time.sleep(15)
    sys.exit()

import IFX_VR14_DC_LL
import IFX_VR14_3D

class dc_thread(QThread):
    dc_to_enable_abort = pyqtSignal(bool)    
    def __init__(self):
        QThread.__init__(self)
        #self.dc_to_enable_abort = pyqtSignal(bool)
        
    def __del__(self):
        self.wait()
    def run(self):
        self.dc_to_enable_abort.emit(True)
        IFX_VR14_DC_LL.vr14_ifx_dc(myWin.rail_name,myWin.vout_list,1,myWin.cool_down_delay,myWin.iout_list,myWin.excel)

    def stop(self):
        self.terminate()

class vr3d_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)
    def __init__(self):
        dc_thread.__init__(self)
    def __del__(self):
        self.wait()
    def run(self):
        IFX_VR14_3D.vr14_3d(myWin.rail_name,myWin.vout_list_3d,myWin.start_freq,myWin.end_freq,myWin.icc_min,myWin.icc_max,myWin.freq_steps_per_decade,myWin.rise_time,myWin.cool_down_delay,myWin.start_duty,myWin.end_duty,myWin.duty_step,myWin.excel)
        
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
        self.setFixedSize(800, 750)
        self.setupUi(self)
        
        self.pushButton_8.clicked.connect(self.dc_thermal_function)
        self.pushButton_2.clicked.connect(self.dc_loadline_function)
        self.pushButton_7.clicked.connect(self.stop_all_thread)
        self.pushButton.clicked.connect(self.vr3d_function)
        self.pushButton_9.clicked.connect(self.check_vendor_id)
        self.comboBox.activated.connect(self.qbox)
        self.lcdNumber.setHexMode()
        self.lcdNumber.display(0xFF)
        ## init thread
        self.dc_loadline_thread=dc_thread()
        self.vr3d=vr3d_thread()
        self.dc_thermal_thread=dc_thermal_thread()

        ## set signal connection
        #self.dc_loadline_thread.dc_to_enable_abort.connect(self.set_abort_enable)

        self.start_up()
    def update_GUI(self):
        ## get GUI import
        self.rail_name=self.comboBox.currentText()
        self.vout_list=eval("["+str(self.lineEdit_10.text())+"]")
        self.iout_list=eval("["+str(self.lineEdit_11.text())+"]")
        self.cool_down_delay=int(self.lineEdit_14.text())
        self.excel=self.checkBox_2.isChecked()

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

    def dc_thermal_function(self):
        self.update_GUI()
        self.runing_items()
        self.dc_thermal_thread.start()
        
    def dc_loadline_function(self):
        self.update_GUI()
        self.runing_items()        
        self.dc_loadline_thread.start()

    def stop_all_thread(self):
        print("====")
        print("abort test")
        self.start_up()
        self.dc_loadline_thread.stop()
        self.vr3d.stop()
        self.dc_thermal_thread.stop()

    def check_vendor_id(self):
        svid_addr,svid_bus=ifx.rail_name_to_svid_parameter(self.comboBox.currentText())
        self.vendor_id=ifx.get_svid_reg(svid_addr,svid_bus,0)
        if self.vendor_id == '13':
            self.pushButton_8.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton.setEnabled(True)
        else:
            self.start_up()
            self.pushButton_9.setEnabled(True)
        self.lcdNumber.display(self.vendor_id)
      
    def qbox(self):
        self.pushButton_9.setEnabled(True)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(True)
        
    def start_up(self):
        self.comboBox.setEnabled(True)
        self.pushButton_9.setEnabled(False)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(False)
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
