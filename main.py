import pythoncom
import sys
pythoncom.CoInitialize()  # to fix issue between clr and QFileDialog

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
import json
import IFX_set_wp
import IFX_VR14_3D
import IFX_VR14_DC_LL
import time
import fivra_3d
import ifx
import Ui_PyQT_IFX_VR14



if ifx.detect_process("Gen5.exe"):
    pass
else:
    print("please run Gen5 first")
    time.sleep(15)
    sys.exit()


class myprogpressbar(QThread):
    updatebar = pyqtSignal(int)

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
        IFX_VR14_DC_LL.vr14_ifx_dc(myWin.rail_name, myWin.vout_list, 1,
                                   myWin.cool_down_delay, myWin.iout_list, myWin.excel, myWin.hc_bit)

    def stop(self):
        self.terminate()


class vr3d_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)

    def __init__(self):
        dc_thread.__init__(self)

    def __del__(self):
        self.wait()

    def run(self):
        IFX_VR14_3D.vr14_3d(myWin.rail_name, myWin.vout_list_3d, myWin.start_freq, myWin.end_freq, myWin.icc_min, myWin.icc_max,
                            myWin.freq_steps_per_decade, myWin.rise_time, myWin.cool_down_delay_3d, myWin.start_duty, myWin.end_duty, myWin.duty_step, myWin.excel)

    def stop(self):
        self.terminate()


class dc_thermal_thread(QThread):
    vr3d_to_enable_abort = pyqtSignal(bool)

    def __init__(self):
        dc_thread.__init__(self)

    def __del__(self):
        self.wait()

    def run(self):
        IFX_VR14_DC_LL.thermal_drop(
            myWin.rail_name, myWin.thermal_vout_list, myWin.thermal_iout, myWin.thermal_time_list)

    def stop(self):
        self.terminate()


class fivra_3d_thread(QThread):
    #vr3d_to_enable_abort = pyqtSignal(bool)
    def __init__(self):
        dc_thread.__init__(self)

    def __del__(self):
        self.wait()

    def run(self):
        fivra_3d.fivra_3d(target_voltage=myWin.fivra_3d_target_vout,
                          start_frequency=myWin.fivra_3d_start_frequency,
                          end_frequency=myWin.fivra_3d_end_frequency,
                          sample_per_decade=myWin.fivra_3d_sample_per_decade,
                          nw_current_low=myWin.fivra_3d_nw_current_low,
                          nw_current_high=myWin.fivra_3d_nw_current_high,
                          ne_current_low=myWin.fivra_3d_ne_current_low,
                          ne_current_high=myWin.fivra_3d_ne_current_high,
                          sw_current_low=myWin.fivra_3d_sw_current_low,
                          sw_current_high=myWin.fivra_3d_sw_current_high,
                          se_current_low=myWin.fivra_3d_se_current_low,
                          se_current_high=myWin.fivra_3d_se_current_high,
                          nw_slew_rate=myWin.fivra_3d_nw_slew_rate,
                          ne_slew_rate=myWin.fivra_3d_ne_slew_rate,
                          sw_slew_rate=myWin.fivra_3d_sw_slew_rate,
                          se_slew_rate=myWin.fivra_3d_se_slew_rate,
                          cooldown_delay=myWin.fivra_3d_cooldown_delay,
                          duty_cycle_list=myWin.fivra_3d_duty_cycle_list,
                          debug=True,
                          )

    def stop(self):
        self.terminate()

    def stop(self):
        self.terminate()


class MyMainWindow(QMainWindow, Ui_PyQT_IFX_VR14.Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setFixedSize(840, 940)
        self.setupUi(self)

        # reflash power rails to all combobox.
        self.current_power_rail_dict = ifx.detect_all_rails()
        self.reflash_power_rails()
        self.setwp_change_vid_selection_function()
        self.pushButton_14.clicked.connect(self.set_wp_dc_load_function)
        self.set_wp_dc_load_status = False

        self.pushButton_8.clicked.connect(self.dc_thermal_function)
        self.pushButton_2.clicked.connect(self.dc_loadline_function)
        self.pushButton_7.clicked.connect(self.stop_all_thread)
        self.pushButton.clicked.connect(self.vr3d_function)
        self.pushButton_9.clicked.connect(self.check_vendor_id)
        # self.pushButton_5.clicked.connect(self.save_parameter_function)
        self.actionSave_Parameter_to_file.triggered.connect(
            self.save_parameter_function)
        # self.pushButton_6.clicked.connect(self.load_parameter_function)
        self.actionLoad_Parameter_to_GUI.triggered.connect(
            self.load_parameter_function)
        self.pushButton_4.clicked.connect(self.load_report_to_3d_function)
        self.pushButton_13.clicked.connect(self.set_wp_vector_function)
        self.pushButton_15.clicked.connect(self.clear_svid_alert_function)

        self.pushButton_17.clicked.connect(self.fivra_loadline_function)
        self.pushButton_18.clicked.connect(self.fivra_3d_function)

        # set vector
        self.checkBox_11.stateChanged.connect(self.select_unslect_all_items)
        self.comboBox_11.currentIndexChanged.connect(self.set_vector_line1)
        self.comboBox_13.currentIndexChanged.connect(self.set_vector_line2)
        self.comboBox_15.currentIndexChanged.connect(self.set_vector_line3)
        self.comboBox_17.currentIndexChanged.connect(self.set_vector_line4)
        self.comboBox_19.currentIndexChanged.connect(self.set_vector_line5)
        self.comboBox_22.currentIndexChanged.connect(self.set_vector_line6)
        self.comboBox_37.currentIndexChanged.connect(self.set_vector_line7)
        self.comboBox_48.currentIndexChanged.connect(self.set_vector_line8)
        self.comboBox_50.currentIndexChanged.connect(self.set_vector_line9)
        self.comboBox_53.currentIndexChanged.connect(self.set_vector_line10)
        self.comboBox_55.currentIndexChanged.connect(self.set_vector_line11)
        self.comboBox_57.currentIndexChanged.connect(self.set_vector_line12)
        #self.comboBox.setStyleSheet("background-color: yellow")
        self.comboBox.currentIndexChanged.connect(
            self.change_VID_selection_function)
        self.comboBox.highlighted.connect(self.change_VID_selection_function)
        self.comboBox_2.currentIndexChanged.connect(
            self.setwp_rail1_vid_combox_function)
        self.comboBox_3.currentIndexChanged.connect(
            self.setwp_rail2_vid_combox_function)
        self.comboBox_4.currentIndexChanged.connect(
            self.setwp_rail3_vid_combox_function)
        self.comboBox_5.currentIndexChanged.connect(
            self.setwp_rail4_vid_combox_function)
        self.comboBox_10.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_12.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_14.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_16.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_18.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_21.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_20.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_49.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_51.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_52.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_54.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)
        self.comboBox_56.currentIndexChanged.connect(
            self.setwp_change_vid_selection_function)

        # self.comboBox_24.currentIndexChanged.connect(self.turn_off_check_button_function)

        # set command list in comboBox of Vector
        command_list = ['setVIDfast', 'setVIDslow', 'setVIDdecay', 'setPS',
                        'setRegaddr', 'setREGdata', 'getREG', 'testMode', 'setWP']
        self.comboBox_11.addItems(command_list)
        self.comboBox_13.addItems(command_list)
        self.comboBox_15.addItems(command_list)
        self.comboBox_17.addItems(command_list)
        self.comboBox_19.addItems(command_list)
        self.comboBox_22.addItems(command_list)
        self.comboBox_37.addItems(command_list)
        self.comboBox_48.addItems(command_list)
        self.comboBox_50.addItems(command_list)
        self.comboBox_53.addItems(command_list)
        self.comboBox_55.addItems(command_list)
        self.comboBox_57.addItems(command_list)

        self.pushButton_10.clicked.connect(self.set_wp_init_function)

        self.comboBox.activated.connect(self.qbox)
        self.lcdNumber.setHexMode()
        self.lcdNumber.display(0xFF)
        # self.pushButton_5.setEnabled(True)
        # self.pushButton_6.setEnabled(True)
        self.pushButton_4.setEnabled(True)

        self.tabWidget.setTabEnabled(1, True)
        self.tabWidget.setTabEnabled(2, False)

        # set SetWP rail combobox disable
        self.comboBox_2.setEnabled(True)
        self.comboBox_2.addItems(list(self.current_power_rail_dict.keys()))
        self.comboBox_3.setEnabled(True)
        self.comboBox_3.addItems(list(self.current_power_rail_dict.keys()))
        self.comboBox_4.setEnabled(True)
        self.comboBox_4.addItems(list(self.current_power_rail_dict.keys()))
        self.comboBox_5.setEnabled(True)
        self.comboBox_5.addItems(list(self.current_power_rail_dict.keys()))
        # init thread
        self.dc_loadline_thread = dc_thread()
        self.vr3d = vr3d_thread()
        self.fivra_3d_thread = fivra_3d_thread()
        self.dc_thermal_thread = dc_thermal_thread()
        self.myprogpressbar = myprogpressbar()

        # set signal connection
        self.myprogpressbar.updatebar.connect(self.update_bar)

        self.start_up()

    def fivra_loadline_function(self):
        print("fivra_LL")
        pass

    def fivra_3d_function(self):
        print("fivra_3d")
        self.update_GUI()
        self.runing_items()
        self.fivra_3d_thread.start()
        self.myprogpressbar.start()

    def setwp_rail1_vid_combox_function(self):
        temp = self.current_power_rail_dict[self.comboBox_2.currentText()]
        # self.comboBox_6.clear()
        # self.comboBox_6.addItem(str(temp[2]))

    def setwp_rail2_vid_combox_function(self):
        temp = self.current_power_rail_dict[self.comboBox_3.currentText()]
        # self.comboBox_7.clear()
        # self.comboBox_7.addItem("5")
        # self.comboBox_7.addItem(str(temp[2]))

    def setwp_rail3_vid_combox_function(self):
        temp = self.current_power_rail_dict[self.comboBox_4.currentText()]
        # self.comboBox_8.clear()
        # self.comboBox_8.addItem(str(temp[2]))

    def setwp_rail4_vid_combox_function(self):
        temp = self.current_power_rail_dict[self.comboBox_5.currentText()]
        # self.comboBox_9.clear()
        # self.comboBox_9.addItem(str(temp[2]))

    def change_VID_selection_function(self):
        self.comboBox_24.clear()
        self.current_power_rail_dict = ifx.detect_all_rails()
        print(self.current_power_rail_dict)
        temp = self.current_power_rail_dict[self.comboBox.currentText()]
        print(f" powerrail {temp[2]}")
        self.comboBox_24.addItem(str(temp[2]))
        if self.comboBox_24.currentText() == '0':
            pass
            # for VCCD_HV fixed VID table
            # self.pushButton_9.setEnabled(False)
        else:
            self.pushButton_9.setEnabled(True)

    def reflash_power_rails(self):
        print("reflash")

        self.comboBox.clear()
        self.comboBox.addItems(list(self.current_power_rail_dict.keys()))

        item_length = self.comboBox.count()
        item_list = list()
        for i in range(0, item_length):
            item_list.append(self.comboBox.itemText(i))
        # item_list.append("allCall_0xF")
        self.comboBox_10.clear()
        self.comboBox_12.clear()
        self.comboBox_14.clear()
        self.comboBox_16.clear()
        self.comboBox_18.clear()
        self.comboBox_21.clear()
        self.comboBox_20.clear()
        self.comboBox_49.clear()
        self.comboBox_51.clear()
        self.comboBox_52.clear()
        self.comboBox_54.clear()
        self.comboBox_56.clear()
        self.comboBox_10.addItems(item_list)
        self.comboBox_12.addItems(item_list)
        self.comboBox_14.addItems(item_list)
        self.comboBox_16.addItems(item_list)
        self.comboBox_18.addItems(item_list)
        self.comboBox_21.addItems(item_list)
        self.comboBox_20.addItems(item_list)
        self.comboBox_49.addItems(item_list)
        self.comboBox_51.addItems(item_list)
        self.comboBox_52.addItems(item_list)
        self.comboBox_54.addItems(item_list)
        self.comboBox_56.addItems(item_list)

        self.qbox()

    def clear_svid_alert_function(self):
        IFX_set_wp.ResetAllVectors()
        for SVID_addr in range(0, 4):
            IFX_set_wp.add_vector_into_vector(
                index=SVID_addr, SVID_address=SVID_addr, SVID_command=7, SVID_payload=0x10, SVID_delay=200)

        IFX_set_wp.run_vector()

    def setwp_change_vid_selection_function(self):
        pass
        '''
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
        '''

    def set_wp_vector_function(self):
        vector_index = 0
        IFX_set_wp.ResetAllVectors()

        # vector item_1
        if self.checkBox.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_60.text())*1000, str(self.comboBox_23.currentText()))
            item_1_payload_data = int(temp, 16)
            item_1_delay = int(self.lineEdit_81.text())
            item_1_svid_addr, item_1_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_10.currentText())
            if self.comboBox_11.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_1_svid_addr, SVID_command=(
                    self.comboBox_11.currentIndex()+1), SVID_payload=item_1_payload_data, SVID_delay=item_1_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_1_svid_addr, SVID_command=(
                    self.comboBox_11.currentIndex()+1), SVID_payload=int(self.lineEdit_59.text(), 16), SVID_delay=item_1_delay)
            vector_index += 1
        # vector item_2
        if self.checkBox_7.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_62.text())*1000, str(self.comboBox_25.currentText()))
            item_2_payload_data = int(temp, 16)
            item_2_delay = int(self.lineEdit_86.text())
            item_2_svid_addr, item_2_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_12.currentText())
            if self.comboBox_13.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_2_svid_addr, SVID_command=(
                    self.comboBox_13.currentIndex()+1), SVID_payload=item_2_payload_data, SVID_delay=item_2_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_2_svid_addr, SVID_command=(
                    self.comboBox_13.currentIndex()+1), SVID_payload=int(self.lineEdit_61.text(), 16), SVID_delay=item_2_delay)
            vector_index += 1

        # vector item_3
        if self.checkBox_8.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_64.text())*1000, str(self.comboBox_26.currentText()))
            item_3_payload_data = int(temp, 16)
            item_3_delay = int(self.lineEdit_85.text())
            item_3_svid_addr, item_3_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_14.currentText())
            if self.comboBox_15.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_3_svid_addr, SVID_command=(
                    self.comboBox_15.currentIndex()+1), SVID_payload=item_3_payload_data, SVID_delay=item_3_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_3_svid_addr, SVID_command=(
                    self.comboBox_15.currentIndex()+1), SVID_payload=int(self.lineEdit_63.text(), 16), SVID_delay=item_3_delay)
            vector_index += 1
        # vector item_4
        if self.checkBox_9.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_66.text())*1000, str(self.comboBox_27.currentText()))
            item_4_payload_data = int(temp, 16)
            item_4_delay = int(self.lineEdit_83.text())
            item_4_svid_addr, item_4_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_16.currentText())
            if self.comboBox_17.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_4_svid_addr, SVID_command=(
                    self.comboBox_17.currentIndex()+1), SVID_payload=item_4_payload_data, SVID_delay=item_4_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_4_svid_addr, SVID_command=(
                    self.comboBox_17.currentIndex()+1), SVID_payload=int(self.lineEdit_65.text(), 16), SVID_delay=item_4_delay)
            vector_index += 1

        # vector item_5
        if self.checkBox_17.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_68.text())*1000, str(self.comboBox_28.currentText()))
            item_5_payload_data = int(temp, 16)
            item_5_delay = int(self.lineEdit_80.text())
            item_5_svid_addr, item_5_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_18.currentText())
            if self.comboBox_19.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_5_svid_addr, SVID_command=(
                    self.comboBox_19.currentIndex()+1), SVID_payload=item_5_payload_data, SVID_delay=item_5_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_5_svid_addr, SVID_command=(
                    self.comboBox_19.currentIndex()+1), SVID_payload=int(self.lineEdit_67.text(), 16), SVID_delay=item_5_delay)
            vector_index += 1

          # vector item_6
        if self.checkBox_18.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_70.text())*1000, str(self.comboBox_29.currentText()))
            item_6_payload_data = int(temp, 16)
            item_6_delay = int(self.lineEdit_84.text())
            item_6_svid_addr, item_6_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_21.currentText())
            if self.comboBox_22.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_6_svid_addr, SVID_command=(
                    self.comboBox_22.currentIndex()+1), SVID_payload=item_6_payload_data, SVID_delay=item_6_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_6_svid_addr, SVID_command=(
                    self.comboBox_22.currentIndex()+1), SVID_payload=int(self.lineEdit_69.text(), 16), SVID_delay=item_6_delay)
            vector_index += 1

          # vector item_7
        if self.checkBox_19.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_72.text())*1000, str(self.comboBox_30.currentText()))
            item_7_payload_data = int(temp, 16)
            item_7_delay = int(self.lineEdit_82.text())
            item_7_svid_addr, item_7_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_20.currentText())
            if self.comboBox_37.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_7_svid_addr, SVID_command=(
                    self.comboBox_37.currentIndex()+1), SVID_payload=item_7_payload_data, SVID_delay=item_7_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_7_svid_addr, SVID_command=(
                    self.comboBox_37.currentIndex()+1), SVID_payload=int(self.lineEdit_71.text(), 16), SVID_delay=item_7_delay)
            vector_index += 1

          # vector item_8
        if self.checkBox_29.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_98.text())*1000, str(self.comboBox_31.currentText()))
            item_8_payload_data = int(temp, 16)
            item_8_delay = int(self.lineEdit_113.text())
            item_8_svid_addr, item_8_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_49.currentText())
            if self.comboBox_48.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_8_svid_addr, SVID_command=(
                    self.comboBox_48.currentIndex()+1), SVID_payload=item_8_payload_data, SVID_delay=item_8_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_8_svid_addr, SVID_command=(
                    self.comboBox_48.currentIndex()+1), SVID_payload=int(self.lineEdit_97.text(), 16), SVID_delay=item_8_delay)
            vector_index += 1

          # vector item_9
        if self.checkBox_30.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_99.text())*1000, str(self.comboBox_32.currentText()))
            item_9_payload_data = int(temp, 16)
            item_9_delay = int(self.lineEdit_114.text())
            item_9_svid_addr, item_9_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_51.currentText())
            if self.comboBox_50.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_9_svid_addr, SVID_command=(
                    self.comboBox_50.currentIndex()+1), SVID_payload=item_9_payload_data, SVID_delay=item_9_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_9_svid_addr, SVID_command=(
                    self.comboBox_50.currentIndex()+1), SVID_payload=int(self.lineEdit_100.text(), 16), SVID_delay=item_9_delay)
            vector_index += 1

          # vector item_10
        if self.checkBox_31.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_102.text())*1000, str(self.comboBox_33.currentText()))
            item_10_payload_data = int(temp, 16)
            item_10_delay = int(self.lineEdit_112.text())
            item_10_svid_addr, item_10_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_52.currentText())
            if self.comboBox_53.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_10_svid_addr, SVID_command=(
                    self.comboBox_53.currentIndex()+1), SVID_payload=item_10_payload_data, SVID_delay=item_10_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_10_svid_addr, SVID_command=(
                    self.comboBox_53.currentIndex()+1), SVID_payload=int(self.lineEdit_101.text(), 16), SVID_delay=item_10_delay)
            vector_index += 1

          # vector item_11
        if self.checkBox_32.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_103.text())*1000, str(self.comboBox_34.currentText()))
            item_11_payload_data = int(temp, 16)
            item_11_delay = int(self.lineEdit_116.text())
            item_11_svid_addr, item_11_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_54.currentText())
            if self.comboBox_55.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_11_svid_addr, SVID_command=(
                    self.comboBox_55.currentIndex()+1), SVID_payload=item_11_payload_data, SVID_delay=item_11_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_11_svid_addr, SVID_command=(
                    self.comboBox_55.currentIndex()+1), SVID_payload=int(self.lineEdit_104.text(), 16), SVID_delay=item_11_delay)
            vector_index += 1

          # vector item_12
        if self.checkBox_33.isChecked() == 1:
            temp = ifx.VID_10mV_to_hex(
                float(self.lineEdit_106.text())*1000, str(self.comboBox_35.currentText()))
            item_12_payload_data = int(temp, 16)
            item_12_delay = int(self.lineEdit_115.text())
            item_12_svid_addr, item_12_svid_bus = ifx.rail_name_to_svid_parameter(
                self.comboBox_56.currentText())
            if self.comboBox_57.currentIndex() <= 1:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_12_svid_addr, SVID_command=(
                    self.comboBox_56.currentIndex()+1), SVID_payload=item_12_payload_data, SVID_delay=item_12_delay)
            else:
                IFX_set_wp.add_vector_into_vector(index=vector_index, SVID_address=item_12_svid_addr, SVID_command=(
                    self.comboBox_56.currentIndex()+1), SVID_payload=int(self.lineEdit_105.text(), 16), SVID_delay=item_12_delay)
            vector_index += 1

        if vector_index == 0:
            pass
            # IFX_set_wp.ResetAllVectors()

        if self.comboBox_23.currentText() != "0" and self.comboBox_25.currentText() != "0":
            IFX_set_wp.run_vector()
        else:
            print("VID table is 0")
            # IFX_set_wp.ResetAllVectors()

    def set_wp_dc_load_function(self):
        if self.set_wp_dc_load_status == False:
            IFX_VR14_DC_LL.set_wp_inti_DC_load(self.comboBox_2.currentText(), int(self.lineEdit_22.text()),
                                               self.comboBox_3.currentText(), int(self.lineEdit_25.text()),
                                               self.comboBox_4.currentText(), int(self.lineEdit_28.text()),
                                               self.comboBox_5.currentText(), int(self.lineEdit_31.text()),
                                               True)
            self.pushButton_14.setStyleSheet("background-color: red")
            self.pushButton_14.setText("loading is On now")
            self.set_wp_dc_load_status = True
        else:
            self.pushButton_14.setStyleSheet("background-color: #01FF00;")
            IFX_VR14_DC_LL.set_wp_inti_DC_load(self.comboBox_2.currentText(), int(self.lineEdit_22.text()),
                                               self.comboBox_3.currentText(), int(self.lineEdit_25.text()),
                                               self.comboBox_4.currentText(), int(self.lineEdit_28.text()),
                                               self.comboBox_5.currentText(), int(self.lineEdit_31.text()),
                                               False)
            self.pushButton_14.setText("loading is off now")
            self.set_wp_dc_load_status = False

    def select_unslect_all_items(self):
        self.checkBox.setChecked(self.checkBox_11.isChecked())
        self.checkBox_7.setChecked(self.checkBox_11.isChecked())
        self.checkBox_8.setChecked(self.checkBox_11.isChecked())
        self.checkBox_9.setChecked(self.checkBox_11.isChecked())
        self.checkBox_17.setChecked(self.checkBox_11.isChecked())
        self.checkBox_18.setChecked(self.checkBox_11.isChecked())
        self.checkBox_19.setChecked(self.checkBox_11.isChecked())
        self.checkBox_29.setChecked(self.checkBox_11.isChecked())
        self.checkBox_30.setChecked(self.checkBox_11.isChecked())
        self.checkBox_31.setChecked(self.checkBox_11.isChecked())
        self.checkBox_32.setChecked(self.checkBox_11.isChecked())
        self.checkBox_33.setChecked(self.checkBox_11.isChecked())

    def set_vector_line1(self, i):
        if i <= 2:
            self.lineEdit_59.setEnabled(False)
            self.lineEdit_60.setEnabled(True)
        else:
            self.lineEdit_59.setEnabled(True)
            self.lineEdit_60.setEnabled(False)

    def set_vector_line2(self, i):
        if i <= 2:
            self.lineEdit_61.setEnabled(False)
            self.lineEdit_62.setEnabled(True)
        else:
            self.lineEdit_61.setEnabled(True)
            self.lineEdit_62.setEnabled(False)

    def set_vector_line3(self, i):
        if i <= 2:
            self.lineEdit_63.setEnabled(False)
            self.lineEdit_64.setEnabled(True)
        else:
            self.lineEdit_63.setEnabled(True)
            self.lineEdit_64.setEnabled(False)

    def set_vector_line4(self, i):
        if i <= 2:
            self.lineEdit_65.setEnabled(False)
            self.lineEdit_66.setEnabled(True)
        else:
            self.lineEdit_65.setEnabled(True)
            self.lineEdit_66.setEnabled(False)

    def set_vector_line5(self, i):
        if i <= 2:
            self.lineEdit_67.setEnabled(False)
            self.lineEdit_68.setEnabled(True)
        else:
            self.lineEdit_67.setEnabled(True)
            self.lineEdit_68.setEnabled(False)

    def set_vector_line6(self, i):
        if i <= 2:
            self.lineEdit_69.setEnabled(False)
            self.lineEdit_70.setEnabled(True)
        else:
            self.lineEdit_69.setEnabled(True)
            self.lineEdit_70.setEnabled(False)

    def set_vector_line7(self, i):
        if i <= 2:
            self.lineEdit_71.setEnabled(False)
            self.lineEdit_72.setEnabled(True)
        else:
            self.lineEdit_71.setEnabled(True)
            self.lineEdit_72.setEnabled(False)

    def set_vector_line8(self, i):
        if i <= 2:
            self.lineEdit_97.setEnabled(False)
            self.lineEdit_98.setEnabled(True)
        else:
            self.lineEdit_97.setEnabled(True)
            self.lineEdit_98.setEnabled(False)

    def set_vector_line9(self, i):
        if i <= 2:
            self.lineEdit_100.setEnabled(False)
            self.lineEdit_99.setEnabled(True)
        else:
            self.lineEdit_100.setEnabled(True)
            self.lineEdit_99.setEnabled(False)

    def set_vector_line10(self, i):
        if i <= 2:
            self.lineEdit_101.setEnabled(False)
            self.lineEdit_102.setEnabled(True)
        else:
            self.lineEdit_101.setEnabled(True)
            self.lineEdit_102.setEnabled(False)

    def set_vector_line11(self, i):
        if i <= 2:
            self.lineEdit_104.setEnabled(False)
            self.lineEdit_103.setEnabled(True)
        else:
            self.lineEdit_104.setEnabled(True)
            self.lineEdit_103.setEnabled(False)

    def set_vector_line12(self, i):
        if i <= 2:
            self.lineEdit_105.setEnabled(False)
            self.lineEdit_106.setEnabled(True)
        else:
            self.lineEdit_105.setEnabled(True)
            self.lineEdit_106.setEnabled(False)

    def set_wp_init_function(self):
        IFX_set_wp.ResetAllVectors()
        # rail_1
        wp1 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_20.text())*1000, self.comboBox_6.currentText())
        print(self.comboBox_6.currentText())
        rail_1_wp1_new = int(wp1, 16)
        wp2 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_21.text())*1000, self.comboBox_6.currentText())
        rail_1_wp2_new = int(wp2, 16)

        # rail_2
        wp1 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_23.text())*1000, self.comboBox_7.currentText())
        rail_2_wp1_new = int(wp1, 16)
        wp2 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_24.text())*1000, self.comboBox_7.currentText())
        rail_2_wp2_new = int(wp2, 16)

        # rail_3
        wp1 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_26.text())*1000, self.comboBox_8.currentText())
        rail_3_wp1_new = int(wp1, 16)
        wp2 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_27.text())*1000, self.comboBox_8.currentText())
        rail_3_wp2_new = int(wp2, 16)

        # rail_4
        wp1 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_29.text())*1000, self.comboBox_9.currentText())
        rail_4_wp1_new = int(wp1, 16)
        wp2 = ifx.VID_10mV_to_hex(
            float(self.lineEdit_30.text())*1000, self.comboBox_9.currentText())
        rail_4_wp2_new = int(wp2, 16)

        # set rail_1 wp1
        rail1_svid_addr, svid_bus = ifx.rail_name_to_svid_parameter(
            self.comboBox_2.currentText())

        IFX_set_wp.add_vector_into_vector(
            index=0, SVID_address=rail1_svid_addr, SVID_command=5, SVID_payload=0x3B, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=1, SVID_address=rail1_svid_addr, SVID_command=6, SVID_payload=rail_1_wp1_new, SVID_delay=400)
        # set rail_1 wp2
        IFX_set_wp.add_vector_into_vector(
            index=2, SVID_address=rail1_svid_addr, SVID_command=5, SVID_payload=0x3C, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=3, SVID_address=rail1_svid_addr, SVID_command=6, SVID_payload=rail_1_wp2_new, SVID_delay=400)

        # set rail_2 wp1
        rail2_svid_addr, svid_bus = ifx.rail_name_to_svid_parameter(
            self.comboBox_3.currentText())
        IFX_set_wp.add_vector_into_vector(
            index=4, SVID_address=rail2_svid_addr, SVID_command=5, SVID_payload=0x3B, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=5, SVID_address=rail2_svid_addr, SVID_command=6, SVID_payload=rail_2_wp1_new, SVID_delay=400)
        # set rail_2 wp2
        IFX_set_wp.add_vector_into_vector(
            index=6, SVID_address=rail2_svid_addr, SVID_command=5, SVID_payload=0x3C, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=7, SVID_address=rail2_svid_addr, SVID_command=6, SVID_payload=rail_2_wp2_new, SVID_delay=400)

        # set rail_3 wp1
        rail3_svid_addr, svid_bus = ifx.rail_name_to_svid_parameter(
            self.comboBox_4.currentText())
        IFX_set_wp.add_vector_into_vector(
            index=8, SVID_address=rail3_svid_addr, SVID_command=5, SVID_payload=0x3B, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=9, SVID_address=rail3_svid_addr, SVID_command=6, SVID_payload=rail_3_wp1_new, SVID_delay=400)
        # set rail_3 wp2
        IFX_set_wp.add_vector_into_vector(
            index=10, SVID_address=rail3_svid_addr, SVID_command=5, SVID_payload=0x3C, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=11, SVID_address=rail3_svid_addr, SVID_command=6, SVID_payload=rail_3_wp2_new, SVID_delay=400)

        # set rail_4 wp1
        rail4_svid_addr, svid_bus = ifx.rail_name_to_svid_parameter(
            self.comboBox_5.currentText())
        IFX_set_wp.add_vector_into_vector(
            index=12, SVID_address=rail4_svid_addr, SVID_command=5, SVID_payload=0x3B, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=13, SVID_address=rail4_svid_addr, SVID_command=6, SVID_payload=rail_4_wp1_new, SVID_delay=400)
        # set rail_4 wp2
        IFX_set_wp.add_vector_into_vector(
            index=14, SVID_address=rail4_svid_addr, SVID_command=5, SVID_payload=0x3C, SVID_delay=400)
        IFX_set_wp.add_vector_into_vector(
            index=15, SVID_address=rail4_svid_addr, SVID_command=6, SVID_payload=rail_4_wp2_new, SVID_delay=400)

        IFX_set_wp.run_vector()
#        time.sleep(0.5)
#        IFX_set_wp.ResetAllVectors()

    def load_report_to_3d_function(self):
        dlg = QFileDialog(self, 'Open File', '.',
                          'Excel report (*.xlsx);;All Files (*)')
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            with open(filenames[0], 'r') as j:
                ifx.plt_vmax(filenames[0])

    def update_bar(self, data):
        self.progressBar.setValue(data)

    def save_parameter_function(self):
        self.update_GUI()

        if self.tabWidget.currentIndex() == 0:
            self.parameter_dict = {"rail_name": self.comboBox.currentText(),
                                   "LL_vout_list": self.lineEdit_10.text(),
                                   "LL_iout_list": self.lineEdit_11.text(),
                                   "LL_cool_down_delay": self.lineEdit_14.text(),

                                   "3d_vout_list": self.lineEdit_16.text(),
                                   "3d_cool_down_delay": self.lineEdit_15.text(),
                                   "3d_start_freq": self.lineEdit.text(),
                                   "3d_end_freq": self.lineEdit_2.text(),
                                   "3d_rise_time": self.lineEdit_6.text(),
                                   "3d_fall_time": self.lineEdit_12.text(),
                                   "3d_icc_min": self.lineEdit_4.text(),
                                   "3d_icc_max": self.lineEdit_3.text(),
                                   "3d_freq_steps_per_decade": self.lineEdit_5.text(),
                                   "3d_duty_step": self.lineEdit_9.text(),
                                   "3d_start_duty": self.lineEdit_8.text(),
                                   "3d_end_duty": self.lineEdit_7.text(),

                                   }
        elif self.tabWidget.currentIndex() == 1:
            print("save vector")
            self.parameter_dict = {"set_wp_init_rail_name1": self.comboBox_2.currentText(),
                                   "set_wp_init_rail_name2": self.comboBox_3.currentText(),
                                   "set_wp_init_rail_name3": self.comboBox_4.currentText(),
                                   "set_wp_init_rail_name4": self.comboBox_5.currentText(),
                                   "set_wp_init_rail_name1_wp1": self.lineEdit_20.text(),
                                   "set_wp_init_rail_name1_wp2": self.lineEdit_21.text(),
                                   "set_wp_init_rail_name2_wp1": self.lineEdit_23.text(),
                                   "set_wp_init_rail_name2_wp2": self.lineEdit_24.text(),
                                   "set_wp_init_rail_name3_wp1": self.lineEdit_26.text(),
                                   "set_wp_init_rail_name3_wp2": self.lineEdit_27.text(),
                                   "set_wp_init_rail_name4_wp1": self.lineEdit_29.text(),
                                   "set_wp_init_rail_name4_wp2": self.lineEdit_30.text(),
                                   "set_wp_init_rail_vid_table1": self.comboBox_6.currentText(),
                                   "set_wp_init_rail_vid_table2": self.comboBox_7.currentText(),
                                   "set_wp_init_rail_vid_table3": self.comboBox_8.currentText(),
                                   "set_wp_init_rail_vid_table4": self.comboBox_9.currentText(),
                                   # vector section
                                   "set_wp_vector_line1_checked": self.checkBox.isChecked(),
                                   "set_wp_vector_line2_checked": self.checkBox_7.isChecked(),
                                   "set_wp_vector_line3_checked": self.checkBox_8.isChecked(),
                                   "set_wp_vector_line4_checked": self.checkBox_9.isChecked(),
                                   "set_wp_vector_line5_checked": self.checkBox_17.isChecked(),
                                   "set_wp_vector_line6_checked": self.checkBox_18.isChecked(),
                                   "set_wp_vector_line7_checked": self.checkBox_19.isChecked(),
                                   "set_wp_vector_line8_checked": self.checkBox_29.isChecked(),
                                   "set_wp_vector_line9_checked": self.checkBox_30.isChecked(),
                                   "set_wp_vector_line10_checked": self.checkBox_31.isChecked(),
                                   "set_wp_vector_line11_checked": self.checkBox_32.isChecked(),
                                   "set_wp_vector_line12_checked": self.checkBox_33.isChecked(),
                                   "set_wp_vector_line1_railname": self.comboBox_9.currentText(),
                                   "set_wp_vector_line2_railname": self.comboBox_12.currentText(),
                                   "set_wp_vector_line3_railname": self.comboBox_14.currentText(),
                                   "set_wp_vector_line4_railname": self.comboBox_16.currentText(),
                                   "set_wp_vector_line5_railname": self.comboBox_18.currentText(),
                                   "set_wp_vector_line6_railname": self.comboBox_21.currentText(),
                                   "set_wp_vector_line7_railname": self.comboBox_20.currentText(),
                                   "set_wp_vector_line8_railname": self.comboBox_49.currentText(),
                                   "set_wp_vector_line9_railname": self.comboBox_51.currentText(),
                                   "set_wp_vector_line10_railname": self.comboBox_52.currentText(),
                                   "set_wp_vector_line11_railname": self.comboBox_54.currentText(),
                                   "set_wp_vector_line12_railname": self.comboBox_56.currentText(),
                                   "set_wp_vector_line1_command": self.comboBox_11.currentText(),
                                   "set_wp_vector_line2_command": self.comboBox_13.currentText(),
                                   "set_wp_vector_line3_command": self.comboBox_15.currentText(),
                                   "set_wp_vector_line4_command": self.comboBox_17.currentText(),
                                   "set_wp_vector_line5_command": self.comboBox_19.currentText(),
                                   "set_wp_vector_line6_command": self.comboBox_22.currentText(),
                                   "set_wp_vector_line7_command": self.comboBox_37.currentText(),
                                   "set_wp_vector_line8_command": self.comboBox_48.currentText(),
                                   "set_wp_vector_line9_command": self.comboBox_50.currentText(),
                                   "set_wp_vector_line10_command": self.comboBox_53.currentText(),
                                   "set_wp_vector_line11_command": self.comboBox_55.currentText(),
                                   "set_wp_vector_line12_command": self.comboBox_57.currentText(),
                                   "set_wp_vector_line1_payload": self.lineEdit_59.text(),
                                   "set_wp_vector_line2_payload": self.lineEdit_61.text(),
                                   "set_wp_vector_line3_payload": self.lineEdit_63.text(),
                                   "set_wp_vector_line4_payload": self.lineEdit_65.text(),
                                   "set_wp_vector_line5_payload": self.lineEdit_67.text(),
                                   "set_wp_vector_line6_payload": self.lineEdit_69.text(),
                                   "set_wp_vector_line7_payload": self.lineEdit_71.text(),
                                   "set_wp_vector_line8_payload": self.lineEdit_97.text(),
                                   "set_wp_vector_line9_payload": self.lineEdit_100.text(),
                                   "set_wp_vector_line10_payload": self.lineEdit_101.text(),
                                   "set_wp_vector_line11_payload": self.lineEdit_104.text(),
                                   "set_wp_vector_line12_payload": self.lineEdit_105.text(),
                                   "set_wp_vector_line1_payload_voltage": self.lineEdit_60.text(),
                                   "set_wp_vector_line2_payload_voltage": self.lineEdit_62.text(),
                                   "set_wp_vector_line3_payload_voltage": self.lineEdit_64.text(),
                                   "set_wp_vector_line4_payload_voltage": self.lineEdit_66.text(),
                                   "set_wp_vector_line5_payload_voltage": self.lineEdit_68.text(),
                                   "set_wp_vector_line6_payload_voltage": self.lineEdit_70.text(),
                                   "set_wp_vector_line7_payload_voltage": self.lineEdit_72.text(),
                                   "set_wp_vector_line8_payload_voltage": self.lineEdit_98.text(),
                                   "set_wp_vector_line9_payload_voltage": self.lineEdit_99.text(),
                                   "set_wp_vector_line10_payload_voltage": self.lineEdit_102.text(),
                                   "set_wp_vector_line11_payload_voltage": self.lineEdit_103.text(),
                                   "set_wp_vector_line12_payload_voltage": self.lineEdit_106.text(),
                                   "set_wp_vector_line1_vid_table": self.comboBox_23.currentText(),
                                   "set_wp_vector_line2_vid_table": self.comboBox_25.currentText(),
                                   "set_wp_vector_line3_vid_table": self.comboBox_26.currentText(),
                                   "set_wp_vector_line4_vid_table": self.comboBox_27.currentText(),
                                   "set_wp_vector_line5_vid_table": self.comboBox_28.currentText(),
                                   "set_wp_vector_line6_vid_table": self.comboBox_29.currentText(),
                                   "set_wp_vector_line7_vid_table": self.comboBox_30.currentText(),
                                   "set_wp_vector_line8_vid_table": self.comboBox_31.currentText(),
                                   "set_wp_vector_line9_vid_table": self.comboBox_32.currentText(),
                                   "set_wp_vector_line10_vid_table": self.comboBox_33.currentText(),
                                   "set_wp_vector_line11_vid_table": self.comboBox_34.currentText(),
                                   "set_wp_vector_line12_vid_table": self.comboBox_35.currentText(),
                                   "set_wp_vector_line1_delay": self.lineEdit_81.text(),
                                   "set_wp_vector_line2_delay": self.lineEdit_86.text(),
                                   "set_wp_vector_line3_delay": self.lineEdit_85.text(),
                                   "set_wp_vector_line4_delay": self.lineEdit_83.text(),
                                   "set_wp_vector_line5_delay": self.lineEdit_80.text(),
                                   "set_wp_vector_line6_delay": self.lineEdit_84.text(),
                                   "set_wp_vector_line7_delay": self.lineEdit_82.text(),
                                   "set_wp_vector_line8_delay": self.lineEdit_113.text(),
                                   "set_wp_vector_line9_delay": self.lineEdit_114.text(),
                                   "set_wp_vector_line10_delay": self.lineEdit_112.text(),
                                   "set_wp_vector_line11_delay": self.lineEdit_116.text(),
                                   "set_wp_vector_line12_delay": self.lineEdit_115.text(),
                                   "set_wp_loop_times": self.lineEdit_58.text(),
                                   }
        elif self.tabWidget.currentIndex() == 2:
            self.parameter_dict = {"fivra_3d_target_vout": self.lineEdit_41.text(),
                                   "fivra_3d_start_frequency": self.lineEdit_42.text(),
                                   "fivra_3d_end_frequency": self.lineEdit_36.text(),
                                   "fivra_3d_sample_per_decade": self.lineEdit_45.text(),
                                   "fivra_3d_nw_current_low": self.lineEdit_47.text(),
                                   "fivra_3d_nw_current_high": self.lineEdit_48.text(),
                                   "fivra_3d_ne_current_low": self.lineEdit_51.text(),
                                   "fivra_3d_ne_current_high": self.lineEdit_50.text(),
                                   "fivra_3d_sw_current_low": self.lineEdit_55.text(),
                                   "fivra_3d_sw_current_high": self.lineEdit_54.text(),
                                   "fivra_3d_se_current_low": self.lineEdit_57.text(),
                                   "fivra_3d_se_current_high": self.lineEdit_53.text(),
                                   "fivra_3d_nw_slew_rate": self.lineEdit_49.text(),
                                   "fivra_3d_ne_slew_rate": self.lineEdit_52.text(),
                                   "fivra_3d_sw_slew_rate": self.lineEdit_56.text(),
                                   "fivra_3d_se_slew_rate": self.lineEdit_73.text(),
                                   "fivra_3d_cooldown_delay": self.lineEdit_35.text(),
                                   "fivra_3d_duty_cycle_list": self.lineEdit_40.text(),
                                   }

        else:
            pass
        filename_with_path = QFileDialog.getSaveFileName(
            self, 'Save File', '.', 'JSON Files (*.json)')
        save_filename = filename_with_path[0]
        if save_filename != "":
            with open(save_filename, 'w') as fp:
                #json.dump(self.parameter_dict, fp)
                fp.write(json.dumps(self.parameter_dict, indent=4))
        # else:
            # print(self.tabWidget.currentIndex)
            #print("only can save first tab now")

        self.parameter_dict.clear()

    def load_parameter_function(self):

        try:
            dlg = QFileDialog(self, 'Open File', '.',
                              'JSON Files (*.json);;All Files (*)')
            if self.tabWidget.currentIndex() == 0:
                if dlg.exec_():
                    filenames = dlg.selectedFiles()

                    with open(filenames[0], 'r') as j:
                        json_data = json.load(j)
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
                    self.lineEdit_5.setText(
                        json_data["3d_freq_steps_per_decade"])
                    self.lineEdit_9.setText(json_data["3d_duty_step"])
                    self.lineEdit_8.setText(json_data["3d_start_duty"])
                    self.lineEdit_7.setText(json_data["3d_end_duty"])
            elif self.tabWidget.currentIndex() == 1:
                if dlg.exec_():
                    filenames = dlg.selectedFiles()

                    with open(filenames[0], 'r') as j:
                        json_data = json.load(j)
                        print(json_data)
                    self.comboBox_2.setCurrentText(
                        json_data["set_wp_init_rail_name1"])
                    self.comboBox_3.setCurrentText(
                        json_data["set_wp_init_rail_name2"])
                    self.comboBox_4.setCurrentText(
                        json_data["set_wp_init_rail_name3"])
                    self.comboBox_5.setCurrentText(
                        json_data["set_wp_init_rail_name4"])
                    self.lineEdit_20.setText(
                        json_data["set_wp_init_rail_name1_wp1"])
                    self.lineEdit_21.setText(
                        json_data["set_wp_init_rail_name1_wp2"])
                    self.lineEdit_23.setText(
                        json_data["set_wp_init_rail_name2_wp1"])
                    self.lineEdit_24.setText(
                        json_data["set_wp_init_rail_name2_wp2"])
                    self.lineEdit_26.setText(
                        json_data["set_wp_init_rail_name3_wp1"])
                    self.lineEdit_27.setText(
                        json_data["set_wp_init_rail_name3_wp2"])
                    self.lineEdit_29.setText(
                        json_data["set_wp_init_rail_name4_wp1"])
                    self.lineEdit_30.setText(
                        json_data["set_wp_init_rail_name4_wp1"])
                    self.comboBox_6.setCurrentText(
                        json_data["set_wp_init_rail_vid_table1"])
                    self.comboBox_7.setCurrentText(
                        json_data["set_wp_init_rail_vid_table2"])
                    self.comboBox_8.setCurrentText(
                        json_data["set_wp_init_rail_vid_table3"])
                    self.comboBox_9.setCurrentText(
                        json_data["set_wp_init_rail_vid_table4"])

                    self.checkBox.setChecked(
                        json_data["set_wp_vector_line1_checked"])
                    self.checkBox_7.setChecked(
                        json_data["set_wp_vector_line2_checked"])
                    self.checkBox_8.setChecked(
                        json_data["set_wp_vector_line3_checked"])
                    self.checkBox_9.setChecked(
                        json_data["set_wp_vector_line4_checked"])
                    self.checkBox_17.setChecked(
                        json_data["set_wp_vector_line5_checked"])
                    self.checkBox_18.setChecked(
                        json_data["set_wp_vector_line6_checked"])
                    self.checkBox_19.setChecked(
                        json_data["set_wp_vector_line7_checked"])
                    self.checkBox_29.setChecked(
                        json_data["set_wp_vector_line8_checked"])
                    self.checkBox_30.setChecked(
                        json_data["set_wp_vector_line9_checked"])
                    self.checkBox_31.setChecked(
                        json_data["set_wp_vector_line10_checked"])
                    self.checkBox_32.setChecked(
                        json_data["set_wp_vector_line11_checked"])
                    self.checkBox_33.setChecked(
                        json_data["set_wp_vector_line12_checked"])
                    self.comboBox_10.setCurrentText(
                        json_data["set_wp_vector_line1_railname"])
                    self.comboBox_12.setCurrentText(
                        json_data["set_wp_vector_line2_railname"])
                    self.comboBox_14.setCurrentText(
                        json_data["set_wp_vector_line3_railname"])
                    self.comboBox_16.setCurrentText(
                        json_data["set_wp_vector_line4_railname"])
                    self.comboBox_18.setCurrentText(
                        json_data["set_wp_vector_line5_railname"])
                    self.comboBox_21.setCurrentText(
                        json_data["set_wp_vector_line6_railname"])
                    self.comboBox_20.setCurrentText(
                        json_data["set_wp_vector_line7_railname"])
                    self.comboBox_49.setCurrentText(
                        json_data["set_wp_vector_line8_railname"])
                    self.comboBox_51.setCurrentText(
                        json_data["set_wp_vector_line9_railname"])
                    self.comboBox_52.setCurrentText(
                        json_data["set_wp_vector_line10_railname"])
                    self.comboBox_54.setCurrentText(
                        json_data["set_wp_vector_line11_railname"])
                    self.comboBox_56.setCurrentText(
                        json_data["set_wp_vector_line12_railname"])
                    self.comboBox_11.setCurrentText(
                        json_data["set_wp_vector_line1_command"])
                    self.comboBox_13.setCurrentText(
                        json_data["set_wp_vector_line2_command"])
                    self.comboBox_15.setCurrentText(
                        json_data["set_wp_vector_line3_command"])
                    self.comboBox_17.setCurrentText(
                        json_data["set_wp_vector_line4_command"])
                    self.lineEdit_59.setText(
                        json_data["set_wp_vector_line1_payload"])
                    self.lineEdit_61.setText(
                        json_data["set_wp_vector_line2_payload"])
                    self.lineEdit_63.setText(
                        json_data["set_wp_vector_line3_payload"])
                    self.lineEdit_65.setText(
                        json_data["set_wp_vector_line4_payload"])
                    self.lineEdit_67.setText(
                        json_data["set_wp_vector_line5_payload"])
                    self.lineEdit_69.setText(
                        json_data["set_wp_vector_line6_payload"])
                    self.lineEdit_71.setText(
                        json_data["set_wp_vector_line7_payload"])
                    self.lineEdit_97.setText(
                        json_data["set_wp_vector_line8_payload"])
                    self.lineEdit_100.setText(
                        json_data["set_wp_vector_line9_payload"])
                    self.lineEdit_101.setText(
                        json_data["set_wp_vector_line10_payload"])
                    self.lineEdit_104.setText(
                        json_data["set_wp_vector_line11_payload"])
                    self.lineEdit_105.setText(
                        json_data["set_wp_vector_line12_payload"])
                    self.lineEdit_60.setText(
                        json_data["set_wp_vector_line1_payload_voltage"])
                    self.lineEdit_62.setText(
                        json_data["set_wp_vector_line2_payload_voltage"])
                    self.lineEdit_64.setText(
                        json_data["set_wp_vector_line3_payload_voltage"])
                    self.lineEdit_66.setText(
                        json_data["set_wp_vector_line4_payload_voltage"])
                    self.lineEdit_68.setText(
                        json_data["set_wp_vector_line5_payload_voltage"])
                    self.lineEdit_70.setText(
                        json_data["set_wp_vector_line6_payload_voltage"])
                    self.lineEdit_72.setText(
                        json_data["set_wp_vector_line7_payload_voltage"])
                    self.lineEdit_98.setText(
                        json_data["set_wp_vector_line8_payload_voltage"])
                    self.lineEdit_99.setText(
                        json_data["set_wp_vector_line9_payload_voltage"])
                    self.lineEdit_102.setText(
                        json_data["set_wp_vector_line10_payload_voltage"])
                    self.lineEdit_103.setText(
                        json_data["set_wp_vector_line11_payload_voltage"])
                    self.lineEdit_106.setText(
                        json_data["set_wp_vector_line12_payload_voltage"])
                    self.comboBox_23.setCurrentText(
                        json_data["set_wp_vector_line1_vid_table"])
                    self.comboBox_25.setCurrentText(
                        json_data["set_wp_vector_line2_vid_table"])
                    self.comboBox_26.setCurrentText(
                        json_data["set_wp_vector_line3_vid_table"])
                    self.comboBox_27.setCurrentText(
                        json_data["set_wp_vector_line4_vid_table"])
                    self.comboBox_28.setCurrentText(
                        json_data["set_wp_vector_line5_vid_table"])
                    self.comboBox_29.setCurrentText(
                        json_data["set_wp_vector_line6_vid_table"])
                    self.comboBox_30.setCurrentText(
                        json_data["set_wp_vector_line7_vid_table"])
                    self.comboBox_31.setCurrentText(
                        json_data["set_wp_vector_line8_vid_table"])
                    self.comboBox_32.setCurrentText(
                        json_data["set_wp_vector_line9_vid_table"])
                    self.comboBox_33.setCurrentText(
                        json_data["set_wp_vector_line10_vid_table"])
                    self.comboBox_34.setCurrentText(
                        json_data["set_wp_vector_line11_vid_table"])
                    self.comboBox_35.setCurrentText(
                        json_data["set_wp_vector_line12_vid_table"])
                    self.lineEdit_81.setText(
                        json_data["set_wp_vector_line1_delay"])
                    self.lineEdit_86.setText(
                        json_data["set_wp_vector_line2_delay"])
                    self.lineEdit_85.setText(
                        json_data["set_wp_vector_line3_delay"])
                    self.lineEdit_83.setText(
                        json_data["set_wp_vector_line4_delay"])
                    self.lineEdit_80.setText(
                        json_data["set_wp_vector_line5_delay"])
                    self.lineEdit_84.setText(
                        json_data["set_wp_vector_line6_delay"])
                    self.lineEdit_82.setText(
                        json_data["set_wp_vector_line7_delay"])
                    self.lineEdit_113.setText(
                        json_data["set_wp_vector_line8_delay"])
                    self.lineEdit_114.setText(
                        json_data["set_wp_vector_line9_delay"])
                    self.lineEdit_112.setText(
                        json_data["set_wp_vector_line10_delay"])
                    self.lineEdit_116.setText(
                        json_data["set_wp_vector_line11_delay"])
                    self.lineEdit_115.setText(
                        json_data["set_wp_vector_line12_delay"])
                    self.lineEdit_58.setText(
                        json_data.get("set_wp_loop_times", '0'))
            elif self.tabWidget.currentIndex() == 2:
                if dlg.exec_():
                    filenames = dlg.selectedFiles()

                    with open(filenames[0], 'r') as j:
                        json_data = json.load(j)
                        print(json_data)
                    self.lineEdit_41.setText(
                        json_data["fivra_3d_target_vout"])
                    self.lineEdit_42.setText(
                        json_data["fivra_3d_start_frequency"])
                    self.lineEdit_36.setText(
                        json_data["fivra_3d_end_frequency"])
                    self.lineEdit_45.setText(
                        json_data["fivra_3d_sample_per_decade"])
                    self.lineEdit_47.setText(
                        json_data["fivra_3d_nw_current_low"])
                    self.lineEdit_48.setText(
                        json_data["fivra_3d_nw_current_high"])

                    self.lineEdit_51.setText(
                        json_data["fivra_3d_ne_current_low"])
                    self.lineEdit_50.setText(
                        json_data["fivra_3d_ne_current_high"])
                    self.lineEdit_55.setText(
                        json_data["fivra_3d_sw_current_low"])
                    self.lineEdit_54.setText(
                        json_data["fivra_3d_sw_current_high"])
                    self.lineEdit_57.setText(
                        json_data["fivra_3d_se_current_low"])
                    self.lineEdit_53.setText(
                        json_data["fivra_3d_se_current_high"])
                    self.lineEdit_49.setText(
                        json_data["fivra_3d_nw_slew_rate"])
                    self.lineEdit_52.setText(
                        json_data["fivra_3d_ne_slew_rate"])
                    self.lineEdit_56.setText(
                        json_data["fivra_3d_sw_slew_rate"])
                    self.lineEdit_73.setText(
                        json_data["fivra_3d_se_slew_rate"])
                    self.lineEdit_35.setText(
                        json_data["fivra_3d_cooldown_delay"])
                    self.lineEdit_40.setText(
                        json_data["fivra_3d_duty_cycle_list"])

            else:
                QMessageBox.about(
                    self, "Warning", "It can't support this page")
        except:
            QMessageBox.about(
                self, "Warning", "The JSON file can't match this page, please select another .JSON file")

    def update_GUI(self):
        # get GUI import
        self.rail_name = self.comboBox.currentText()
        self.vout_list = eval("["+str(self.lineEdit_10.text())+"]")
        self.iout_list = eval("["+str(self.lineEdit_11.text())+"]")
        self.cool_down_delay = int(self.lineEdit_14.text())
        self.excel = self.checkBox_2.isChecked()
        self.hc_bit = self.checkBox_10.isChecked()

        # get GUI import for 3D parameter
        self.vout_list_3d = eval("["+str(self.lineEdit_16.text())+"]")
        self.start_freq = float(self.lineEdit.text())
        self.end_freq = int(self.lineEdit_2.text())
        self.rise_time = int(self.lineEdit_6.text())
        self.fall_time = int(self.lineEdit_12.text())
        self.icc_min = int(self.lineEdit_4.text())
        self.icc_max = int(self.lineEdit_3.text())
        self.freq_steps_per_decade = int(self.lineEdit_5.text())
        self.duty_step = int(self.lineEdit_9.text())
        self.start_duty = int(self.lineEdit_8.text())
        self.end_duty = int(self.lineEdit_7.text())
        self.cool_down_delay_3d = int(self.lineEdit_15.text())

        # get GUI import from thermal drop
        self.thermal_vout_list = eval("["+str(self.lineEdit_17.text())+"]")
        self.thermal_time_list = eval("["+str(self.lineEdit_19.text())+"]")
        self.thermal_iout = float(self.lineEdit_18.text())

        # TODO: get GUI paramter for FIVRA page
        self.fivra_3d_target_vout = float(self.lineEdit_41.text())
        self.fivra_3d_start_frequency = float(self.lineEdit_42.text())
        self.fivra_3d_end_frequency = float(self.lineEdit_36.text())
        self.fivra_3d_sample_per_decade = float(self.lineEdit_45.text())
        self.fivra_3d_nw_current_low = float(self.lineEdit_47.text())
        self.fivra_3d_nw_current_high = float(self.lineEdit_48.text())
        self.fivra_3d_ne_current_low = float(self.lineEdit_51.text())
        self.fivra_3d_ne_current_high = float(self.lineEdit_50.text())
        self.fivra_3d_sw_current_low = float(self.lineEdit_55.text())
        self.fivra_3d_sw_current_high = float(self.lineEdit_54.text())
        self.fivra_3d_se_current_low = float(self.lineEdit_57.text())
        self.fivra_3d_se_current_high = float(self.lineEdit_53.text())
        self.fivra_3d_nw_slew_rate = float(self.lineEdit_49.text())
        self.fivra_3d_ne_slew_rate = float(self.lineEdit_52.text())
        self.fivra_3d_sw_slew_rate = float(self.lineEdit_56.text())
        self.fivra_3d_se_slew_rate = float(self.lineEdit_73.text())
        self.fivra_3d_cooldown_delay = float(self.lineEdit_35.text())
        self.fivra_3d_duty_cycle_list = eval(
            "["+str(self.lineEdit_40.text())+"]")

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
        self.fivra_3d_thread.stop()
        self.myprogpressbar.stop()

    def check_vendor_id(self):
        svid_addr, svid_bus = ifx.rail_name_to_svid_parameter(
            self.comboBox.currentText())
        self.vendor_id = ifx.get_svid_reg(svid_addr, svid_bus, 0)
        if self.vendor_id == '13':
            self.pushButton_8.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton.setEnabled(True)
            self.tabWidget.setTabEnabled(1, True)

            self.tabWidget.setTabEnabled(2, True)

        else:
            self.start_up()
            self.pushButton_9.setEnabled(True)
        self.lcdNumber.display(self.vendor_id)

    def qbox(self):

        # self.pushButton_9.setEnabled(True)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(True)
        self.tabWidget.setTabEnabled(1, False)

    def start_up(self):
        self.comboBox.setEnabled(True)
        self.pushButton_9.setEnabled(False)
        self.pushButton.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_7.setEnabled(False)
        self.tabWidget.setTabEnabled(1, False)
        self.tabWidget.setTabEnabled(2, False)
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
    app = QApplication(sys.argv)

    myWin = MyMainWindow()

    myWin.show()

    sys.exit(app.exec_())
