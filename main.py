import sys
from PyQt5.QtWidgets import QMainWindow,QApplication,QTableView
from PyQt5.QtCore import QThread,pyqtSignal,QTime
import Ui_PyQT_IFX_VR14
#import PdQtClass
import time
import datetime
import os
import pandas as pd
#import IFX_VR14_DC_LL

class dc_thread(QThread):
    dc_to_enable_abort = pyqtSignal(bool)    
    def __init__(self):
        QThread.__init__(self)
        #self.dc_to_enable_abort = pyqtSignal(bool)
        
    def __del__(self):
        self.wait()
    def run(self):
        self.dc_to_enable_abort.emit(True)
        #IFX_VR14_DC_LL.vr14_ifx_dc(myWin.rail_name,myWin.vout_list,1,myWin.cool_down_delay,myWin.iout_list,myWin.excel)
        for i in range(1000):
            self.dc_to_enable_abort.emit(True)
            print(f"thread {myWin.iout_list}")
            
            time.sleep(0.01)
    def stop(self):
        self.terminate()

class vr3d_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)
    def __init__(self):
        dc_thread.__init__(self)
    def __del__(self):
        self.wait()
    def run(self):
        #IFX_VR14_DC_LL.vr14_ifx_dc(myWin.rail_name,myWin.vout_list,1,myWin.cool_down_delay,myWin.iout_list,myWin.excel)
        for i in range(10):
            self.vr3d_to_enable_abort.emit(True)
            print(f"thread {myWin.vout_list}")
            time.sleep(0.01)
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
