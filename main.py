import sys
from PyQt5.QtWidgets import QMainWindow,QApplication,QTableView
from PyQt5.QtCore import QThread,pyqtSignal,QTime
import Ui_PyQT_IFX_VR14
#import PdQtClass
import time
import datetime
import os
import pandas as pd
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
        #print("DC")
    def stop(self):
        self.terminate()

class vr3d_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)
    def __init__(self):
        dc_thread.__init__(self)
    def __del__(self):
        self.wait()
    def run(self):
        #print(f"{myWin.vout_list_3d},{myWin.start_freq},{myWin.end_freq},{myWin.rise_time},{myWin.fall_time},{myWin.icc_min},{myWin.icc_max},{myWin.freq_steps_per_decade},{myWin.duty_step},{myWin.duty_step},{myWin.start_duty},{myWin.end_duty},{myWin.cool_down_delay_3d}")
        IFX_VR14_3D.vr14_3d(myWin.rail_name,myWin.vout_list_3d,myWin.start_freq,myWin.end_freq,myWin.icc_min,myWin.icc_max,myWin.freq_steps_per_decade,myWin.rise_time,myWin.cool_down_delay,myWin.start_duty,myWin.end_duty,myWin.duty_step,myWin.excel)
        
    def stop(self):
        self.terminate()


class MyMainWindow(QMainWindow,Ui_PyQT_IFX_VR14.Ui_MainWindow):
    def __init__(self,parent=None):
        super(MyMainWindow,self).__init__(parent)
        self.setFixedSize(800, 750)
        self.setupUi(self)

        self.pushButton_2.clicked.connect(self.dc_loadline_function)
        self.pushButton_7.clicked.connect(self.stop_all_thread)
        self.pushButton.clicked.connect(self.vr3d_function)
        ## init thread
        self.dc_loadline_thread=dc_thread()
        self.vr3d=vr3d_thread()

        ## set signal connection
        self.vr3d.vr3d_to_enable_abort.connect(self.set_abort_enable)
        self.dc_loadline_thread.dc_to_enable_abort.connect(self.set_abort_enable)
    def update_GUI(self):
        ## get GUI import
        self.rail_name=self.comboBox.currentText()
        self.vout_list=eval(self.lineEdit_10.text())
        self.iout_list=eval(self.lineEdit_11.text())
        self.cool_down_delay=int(self.lineEdit_14.text())
        self.excel=self.checkBox_2.isChecked()

        ##get GUI import for 3D parameter      
        self.vout_list_3d=eval(self.lineEdit_16.text())
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
    def vr3d_function(self):
        self.update_GUI()
        self.vr3d.start()
    def set_abort_enable(self):
        pass
        
    def dc_loadline_function(self):

        self.update_GUI()   
        self.dc_loadline_thread.start()

    def stop_all_thread(self):
        
        self.dc_loadline_thread.stop()
        self.vr3d.stop()

 


if __name__ == "__main__":
    app=QApplication(sys.argv)

    myWin=MyMainWindow()

    myWin.show()

    sys.exit(app.exec_())
