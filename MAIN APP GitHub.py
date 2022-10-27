# -*- coding: utf-8 -*-
"""
Created on Sun Sep  4 11:21:03 2022

@author: xavie
"""
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QWidget, QAction, QTabWidget,QVBoxLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import pandas as pd
import numpy as np
import math
import matplotlib.pyplot as plt
from mpl_toolkits import mplot3d
import matplotlib.pyplot as plt
import os
from zipfile import ZipFile
import smtplib
from email.message import EmailMessage
#Email Tab
Email = []
Password = []
#Traverse reduction tab
Folder_loc = []
File_name = []
path_and_file = ''
#Stn coodinates tab
Station = []
Easting = []
Northing = []
Height = []
#Observations tab
Point_at_Stn_list_inport = []
Set_list_inport = []
Face_list_inport = []
Point_ID_list_inport = []
String_list_inport = []
HzBearing_list_inport = []
VzBearing_list_inport = []
HzDist_list_inport = []
SlDist_list_inport = []
Hi_xyz = []
Ht_xyz = []
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Detail and Traverse Survey Processing'
        self.left = 0
        self.top = 0
        self.width = 300
        self.height = 200
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.table_widget = MyTableWidget(self)
        self.setCentralWidget(self.table_widget)
        self.show() 
class MyTableWidget(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)
        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tab4 = QWidget()
        self.tab5 = QWidget()
        self.tab6 = QWidget()
        self.tabs.resize(300,200)
        # Add tabs
        self.tabs.addTab(self.tab1,"Email")
        self.tabs.addTab(self.tab2,"Traverse reductions")
        self.tabs.addTab(self.tab3,"Detail reductions")
        self.tabs.addTab(self.tab4,"Station Coordinates")
        self.tabs.addTab(self.tab5,"Observations")
        self.tabs.addTab(self.tab6,"Data Managment")
        # Create first tab
        self.tab1.layout = QFormLayout(self)
        self.Email_label = QLabel("Email: ", self)
        self.tab1.layout.addRow(self.Email_label)
        self.tab1.setLayout(self.tab1.layout)
        def Submit_email():
            a = self.Email_line.text()
            Email.append(a)
            a = self.Email_line.clear()
            x = Email[-1]
            print('Your email address is: ', x)
        self.Email_line = QLineEdit(self)
        self.tab1.layout.addRow(self.Email_line)
        self.tab1.setLayout(self.tab1.layout)
        self.Password_label = QLabel("App specific password: ", self)
        self.tab1.layout.addRow(self.Password_label)
        self.tab1.setLayout(self.tab1.layout) 
        def Submit_Password():
            a = self.Password_line.text()
            Password.append(a)
            a = self.Password_line.clear()
            x = Password[-1]
            print('Your Password is: ', x)
        self.Password_line = QLineEdit(self)
        self.tab1.layout.addRow(self.Password_line)
        self.tab1.setLayout(self.tab1.layout)
        self.Submit_all_email = QPushButton("Submit all")
        self.tab1.layout.addRow(self.Submit_all_email)
        self.tab1.setLayout(self.tab1.layout)
        self.Submit_all_email.clicked.connect(Submit_email)
        self.Submit_all_email.clicked.connect(Submit_Password)
        def showdialog():
             dlg = QDialog()
             dlg.resize(700, 200)
             b1 = QLabel('Before you close this app for the detail survey you must first create \nan app specific password in your google account so the data can be sent \nto you. To create an app specific password, follow these steps: \nGo to the security tab in your google account: https://myaccount.google.com/security \n> Signing in to Google > App passwords > Select app > Other (Custom name) \n> “Python” >  Generate. Now Copy and paste the password into this app.',dlg)
             b1.move(50,50)
             dlg.setWindowTitle("What is an app specific password?")
             dlg.setWindowModality(Qt.ApplicationModal)
             dlg.exec_()
        self.WhatIs = QPushButton("(What is an app specific password?)")
        self.tab1.layout.addRow(self.WhatIs)
        self.tab1.setLayout(self.tab1.layout)
        self.WhatIs.clicked.connect(showdialog) 
        #TAB2
        self.tab2.layout = QFormLayout(self)
        self.Folder_Location_trav = QLabel("Where is the excel file? (copy and past file path here):", self)
        self.tab2.layout.addRow(self.Folder_Location_trav)
        self.tab2.setLayout(self.tab2.layout)
        def Submit_Folder_location():
            a = self.Folder_line_trav.text()
            a = a.replace("\\", "/")
            b = self.File_line_trav.text()
            b = '/' + b + '.xlsx' 
            path_and_file = a + b
            a = self.Folder_line_trav.clear()
            b = self.File_line_trav.clear()
            print('The location of the file is: ', path_and_file)
            #+++++++++++++++++++++++ALL IMPORTS ARE HERE+++++++++++++++++++++++++++++
            import pandas as pd
            import numpy as np
            import math
            import matplotlib.pyplot as plt
            from mpl_toolkits import mplot3d
            import matplotlib.pyplot as plt
            import os
            from zipfile import ZipFile
            import smtplib
            from email.message import EmailMessage
            import os
            import shutil
            #+++++++++++++++++++++++ALL IMPORTS ARE HERE+++++++++++++++++++++++++++++
            #+++++++++++++++++++++++ALL DEFINITIONS ARE HERE++++++++++++++++++++++++++
            def dms_to_dd(bearing):
                bearing = format(bearing, ".4f")
                bearing_str = str(bearing)
                bearing_deg_str = bearing_str[:-5]
                bearing_min_str = bearing_str[-4:-2]
                bearing_sec_str = bearing_str[-2:]
                bearing_deg = float(bearing_deg_str)
                bearing_min = float(bearing_min_str)
                bearing_sec = float(bearing_sec_str)
                DMS = bearing_deg + (float(bearing_min)/60) + (float(bearing_sec)/3600)
                return DMS
            def Horizontal_FLred(FL, FR):
                temp_face = FL - FR
                if temp_face < 0:
                    temp_face = temp_face + 180
                else: 
                    temp_face = temp_face - 180
                Hz_FLred = FL - 0.5*temp_face
                return Hz_FLred #
            def Vertical_FLred(FL, FR):
                if FL > 180.0000:
                    print("FL was likey observed in FR orientation. Please review the data")
                temp_face = FL + FR - 360
                V_FLred = FL - 0.5*temp_face
                return V_FLred
            def dd_to_rad(bearing):
                radian = bearing * 3.141592653589793/180
                return radian
            def rad_to_dms(bearing):
                radian = bearing * 3.141592653589793/180
                return radian
            def decdeg2dms(dd):
               is_positive = dd >= 0
               dd = abs(dd)
               minutes,seconds = divmod(dd*3600,60)
               degrees,minutes = divmod(minutes,60)
               degrees = degrees if is_positive else -degrees
               #seconds = round(seconds, 0)
               return (degrees,minutes,seconds)
            def TupleToDMS(tupleDMS):
                degrees = str(int(tupleDMS[0]))
                degrees = degrees + '° '
                minutes = str(int(tupleDMS[1]))
                minutes = (minutes + "' ")
                seconds = str(int(tupleDMS[2]))
                seconds = seconds + '"'
                DMS = degrees + minutes + seconds
                return DMS
            def dmsP_to_dmsStr(bearing): 
                dd = dms_to_dd(bearing)
                x = decdeg2dms(dd)
                DMS_Str = TupleToDMS(x)
                return DMS_Str
            def rad_to_dms_Str(bearing): #This function turns radiance into DMS (STRING FORMAT)
                dd = bearing * 180/3.141592653589793
                dms_tuple = decdeg2dms(dd)
                dms_str = TupleToDMS(dms_tuple)
                return dms_str
            def  DMStoSEC(DMS): #input must be in dd
                x = decdeg2dms(DMS)
                degrees = x[0]*3600
                minutes = x[1] * 60
                Seconds = degrees + minutes + x[2]
                Seconds = round(Seconds, 0)
                Seconds = int(Seconds)
                Seconds = str(Seconds) + '"'
                return Seconds #Output in seconds, type str
            def  DDtoDMS_Str(DMS): #input must be in dd
                x = decdeg2dms(DMS)
                DMS_Str = TupleToDMS(x)
                return DMS_Str #Output in seconds, type str
            #+++++++++++++++++++++++ALL DEFINITIONS ARE HERE++++++++++++++++++++++++++
            #+++++++++++++++++++++++Specifiing a directory++++++++++++++++++++++++++
            directory = os.getcwd()
            # print('Directory: ', directory)
            #+++++++++++++++++++++++Specifiing a directory++++++++++++++++++++++++++
            #+++++++++++++++++++++++Creating a temporery export location++++++++++++++
            Export_path = directory.replace("\\", "/")
            os.chdir(Export_path)
            ExportFolder = 'Exports'
            os.makedirs(ExportFolder)
            #+++++++++++++++++++++++Creating a temporery export location++++++++++++++
            #+++++++++++++++++++++++Importing the travers data file+++++++++++++++++++
            Station_table = pd.read_excel (path_and_file, sheet_name='Stations') #place "r" before the path string to address special character, such as '\'. Don't forget to put the file name at the end of the path + '.xlsx'
            Start_bearing_table = pd.read_excel (path_and_file, sheet_name='Start Bearing') #place "r" before the path string to address special character, such as '\'. Don't forget to put the file name at the end of the path + '.xlsx'
            Observations_table = pd.read_excel (path_and_file, sheet_name='Observations') #place "r" before the path string to address special character, such as '\'. Don't forget to put the file name at the end of the path + '.xlsx'
            #+++++++++++++++++++++++Importing the travers data file+++++++++++++++++++
            #+++++++++++++++++++++++Creating lists from the Station table+++++++++++++
            Start_Easting = Station_table['Easting (m)'].tolist()
            x = Start_Easting[0]
            Start_Easting = x
            Start_Northing = Station_table['Northing (m)'].tolist()
            x = Start_Northing[0]
            Start_Northing = x
            Start_Height = Station_table['Height (m)'].tolist()
            x = Start_Height[0]
            Start_Height = x
            #+++++++++++++++++++++++Creating lists from the Station table+++++++++++++
            #+++++++++++++++++Creating lists from the Start bearing table+++++++++++++
            Start_Bearing = Start_bearing_table['Bearing'].tolist()
            x = Start_Bearing[0]
            Start_Bearing = x
            #+++++++++++++++++Creating lists from the Start bearing table+++++++++++++
            #+++++++++++++++++++++++Creating lists from the Obs table+++++++++++++++++
            From_Stn = Observations_table['From Stn'].tolist()
            To_Stn = Observations_table['To Stn'].tolist()
            Face = Observations_table['Face'].tolist()
            Hz_bearing = Observations_table['Horizontal Bearing (dd.mmss)'].tolist()
            V_bearing = Observations_table['Vertical Bearing (dd.mmss)'].tolist()
            HD = Observations_table['Horizontal distance (m)'].tolist()
            SD = Observations_table['Slope distance (m)'].tolist()
            Hi = Observations_table['Height of Instrument (m)'].tolist()
            Ht = Observations_table['Height of Target (m)'].tolist()
            #+++++++++++++++++++++++Creating lists from the Obs table+++++++++++++++++
            #+++++++++++++++++++++++Bearings from DMS to DD compadable units++++++++++
            Hz_bearing_deg = []
            for i in Hz_bearing: 
                Hz_bearing_deg.append(dms_to_dd(i))
            V_bearing_deg = []
            for i in V_bearing: 
                V_bearing_deg.append(dms_to_dd(i))
            #+++++++++++++++++++++++Bearings from DMS to DD compadable units++++++++++
            #+++++++++++++++++++++++Reducing the bearings and distances++++++++++++++++
            first_value = 0
            second_value = 1
            Hz_bearing_red = []
            for i in Hz_bearing_deg:
                if second_value <= len(Hz_bearing_deg):
                    FL = Hz_bearing_deg[first_value]
                    FR = Hz_bearing_deg[second_value]
                    temp = Horizontal_FLred(FL, FR)
                    Hz_bearing_red.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            V_bearing_red = []
            for i in V_bearing_deg:
                if second_value <= len(V_bearing_deg):
                    FL = V_bearing_deg[first_value]
                    FR = V_bearing_deg[second_value]
                    temp = Vertical_FLred(FL, FR)
                    V_bearing_red.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break   
            first_value = 0
            second_value = 1
            HD_red = []
            for i in HD:
                if second_value <= len(HD):
                    dist1 = HD[first_value]
                    dist2 = HD[second_value]
                    temp = (dist1 + dist2)/2
                    HD_red.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            x = HD_red[0]
            HD_red.pop(0)
            HD_red.append(x)   
            first_value = 0
            second_value = 1
            SD_red = []
            for i in HD:
                if second_value <= len(SD):
                    dist1 = SD[first_value]
                    dist2 = SD[second_value]
                    temp = (dist1 + dist2)/2
                    SD_red.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            x = SD_red[0]
            SD_red.pop(0)
            SD_red.append(x)
            #+++++++++++++++++++++++Reducing the bearings and distances++++++++++++++++
            #+++++++++++++++++++++++Second reduction of distances++++++++++++++++++++++
            first_value = 0
            second_value = 1
            HD_red2 = []
            for i in HD_red:
                if second_value <= len(HD_red):
                    dist1 = HD_red[first_value]
                    dist2 = HD_red[second_value]
                    temp = (dist1 + dist2)/2
                    HD_red2.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            SD_red2 = []
            for i in SD_red:
                if second_value <= len(SD_red):
                    dist1 = SD_red[first_value]
                    dist2 = SD_red[second_value]
                    temp = (dist1 + dist2)/2
                    SD_red2.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            #+++++++++++++++++++++++Second reduction of distances++++++++++++++++++++++
            #+++++++++++++++++++++++Reducing the instrument and target hights++++++++++
            first_value = 0
            second_value = 1
            Hi_red = []
            for i in Hi:
                if second_value <= len(Hi):
                    dist1 = Hi[first_value]
                    dist2 = Hi[second_value]
                    temp = (dist1 + dist2)/2
                    Hi_red.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            Ht_red = []
            for i in Ht:
                if second_value <= len(Ht):
                    dist1 = Ht[first_value]
                    dist2 = Ht[second_value]
                    temp = (dist1 + dist2)/2
                    Ht_red.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            #+++++++++++++++++++++++Reducing the instrument and target hights++++++++++
            #+++++++++++++++++++++++Computing internal angles++++++++++++++++++++++++++
            first_value = 0
            second_value = 1
            Internal_angles = [] #DD
            for i in Hz_bearing_red:
                if second_value <= len(Hz_bearing_red):
                    bearing1 = Hz_bearing_red[first_value]
                    bearing2 = Hz_bearing_red[second_value]
                    temp = bearing2 - bearing1
                    Internal_angles.append(temp) #DD
                    first_value +=2
                    second_value +=2
                else:
                    break
            #+++++++++++++++++++++++Computing internal angles++++++++++++++++++++++++++
            #+++++++++++++++++++++++Choosing on direction of vertical anlges+++++++++++
            second_value = 1
            V_bearing_temp = [] #DD
            for i in V_bearing_red:
                if second_value <= len(V_bearing_red):
                        needed = V_bearing_red[second_value]
                        V_bearing_temp.append(needed)
                        second_value +=2
            V_bearing = V_bearing_temp #DD
            #+++++++++++++++++++++++Choosing on direction of vertical anlges+++++++++++
            #++++++++++Choosing on direction of instrument and target height+++++++++++
            second_value = 1
            Hi_red2 = []
            for i in Hi_red:
                if second_value <= len(Hi_red):
                        needed = Hi_red[second_value]
                        Hi_red2.append(needed)
                        second_value +=2
            second_value = 1
            Ht_red2 = []
            for i in Ht_red:
                if second_value <= len(Ht_red):
                        needed = Ht_red[second_value]
                        Ht_red2.append(needed)
                        second_value +=2
            #++++++++++Choosing on direction of instrument and target height+++++++++++
            #+++++++++++++++++++++++Bearings from DD to rad compadable units++++++++++
            Internal_angles_rad = [] #rad
            for i in Internal_angles: 
                Internal_angles_rad.append(dd_to_rad(i))
            V_bearing_rad = [] #rad
            for i in V_bearing: 
                V_bearing_rad.append(dd_to_rad(i))
            #+++++++++++++++++++++++Bearings from DD to rad compadable units++++++++++
            #+++++++++++++++++++++++Creating lists to go in a table+++++++++++++++++++
            Slot = ""
            To_Stn_red = To_Stn[::2]
            From_Stn_red = From_Stn[::2]
            From_Stn_red2 = From_Stn_red[::2]
            temp_From_Stn = []
            for i in From_Stn_red2:
                temp_From_Stn.append(i)
            x = temp_From_Stn[0]
            temp_From_Stn.pop(0)
            temp_From_Stn.append(x)
            From_To = []
            for i in range(len(From_Stn_red2)):
                x = From_Stn_red2[i] + ' to ' + temp_From_Stn[i]
                From_To.append(x)
            LineStation = []
            LineStation.append(From_Stn[0] + '(Known)')
            for i in range(len(From_To)):
                LineStation.append(From_To[i])
                LineStation.append(temp_From_Stn[i])
            LineStation.append(Slot)
            Internal_anlge_list = []
            Internal_anlge_list.append(Slot)
            Internal_anlge_list.append(Slot)
            for i in range(len(From_Stn_red2)):
                Internal_anlge_list.append(rad_to_dms_Str(Internal_angles_rad[i]))
                Internal_anlge_list.append(Slot)
            Bearing_list = []
            for i in range(len(LineStation)):
                Bearing_list.append(Slot)
            Bearing_list[1] = dmsP_to_dmsStr(Start_Bearing)
            V_bearing_list = []
            V_bearing_list.append(Slot)
            for i in range(len(V_bearing)):
                x = dd_to_rad(V_bearing[i])
                x = rad_to_dms_Str(x)
                V_bearing_list.append(x)
                V_bearing_list.append(Slot)
            V_bearing_list.append(Slot)
            HD_list = []
            HD_list.append(Slot)
            for i in range(len(HD_red2)):
                HD_list.append(HD_red2[i])
                HD_list.append(Slot)
            HD_list.append(Slot)
            SD_list = []
            SD_list.append(Slot)
            for i in range(len(SD_red2)):
                SD_list.append(SD_red2[i])
                SD_list.append(Slot)
            SD_list.append(Slot)
            blank = []
            for i in range(len(LineStation)):
                blank.append(Slot)
            Easting_list = []
            for i in range(len(LineStation)):
                Easting_list.append(Slot)
            Easting_list[0] = Start_Easting
            Northing_list = []
            for i in range(len(LineStation)):
                Northing_list.append(Slot)
            Northing_list[0] = Start_Northing
            Height_list = []
            for i in range(len(LineStation)):
                Height_list.append(Slot)
            Height_list[0] = Start_Height
            #+++++++++++++++++++++++Creating lists to go in a table+++++++++++++++++++
            #+++++++++++++++++++++++Exporting the reduced table+++++++++++++++++++++++
            Hz_bearing_red_DMS = []
            for i in range(len(Hz_bearing_red)): 
                x = Hz_bearing_red[i]
                x = dmsP_to_dmsStr(x)
                Hz_bearing_red_DMS.append(x)
            V_bearing_red_DMS = []
            for i in range(len(V_bearing_red)): 
                x = V_bearing_red[i]
                x = dmsP_to_dmsStr(x)
                V_bearing_red_DMS.append(x)
            reduced_table = pd.DataFrame(list(zip(From_Stn_red, To_Stn_red, Hz_bearing_red_DMS, V_bearing_red_DMS, HD_red, SD_red, Hi_red, Ht_red)),
                            columns =['From Stn', 'To Stn', 'Horizontal Bearing (DMS)', 'Vertical Bearing (DMS)', 'Horizontal distance (m)', 'Slope distance (m)', 'Height of Instrument (m)', 'Height of Target (m)'])
            file_name = Export_path + '/' + ExportFolder + '/Reduced table.xlsx'
            reduced_table.to_excel(file_name)
            #+++++++++++++++++++++++Exporting the reduced table+++++++++++++++++++++++
            #+++++++++++++++++Creating a text file to write in++++++++++++++++++++++++
            file_name = Export_path + '/' + ExportFolder + '/Working out.txt'
            f = open(file_name, "a")
            #+++++++++++++++++Creating a text file to write in++++++++++++++++++++++++
            #+++++++++++++++++Angular misclose++++++++++++++++++++++++++++++++++++++++
            n = len(From_Stn_red2)
            AM = sum(Internal_angles) - 180 * (n-2)
            AM_DMS = DMStoSEC(AM)
            sum_angles = DDtoDMS_Str(sum(Internal_angles))
            second_part = str(180 * (n-2)) + '° = ' + AM_DMS 
            f.write('Step 1.  Angular misclose (AM)')
            f.write('\n')
            f.write('\n')
            f.write('AM = Sum(Angles) - 180° x (n-2),    where n is the number of stations ')
            f.write('\n')
            sentence = sum_angles + " - " + "180° x (" + str(n) + " - 2)"  
            f.write(sentence)
            f.write('\n')
            sentence = sum_angles + ' - ' + second_part
            f.write(sentence)
            f.write('\n')
            f.write('\n')
            f.write('\n')
            #+++++++++++++++++Angular misclose++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Angular correction++++++++++++++++++++++++++++++++++++++++
            f.write('Step 2. Angular correction (AC)')
            f.write('\n')
            f.write('\n')
            f.write('AC = -(AM)/n')
            f.write('\n')
            AC = -(AM/n)
            AC_DMS = DMStoSEC(AC)
            sentence = 'AC = -(' +  AM_DMS + ')/' + str(n) + ' = -' + AC_DMS
            f.write(sentence)
            f.write('\n')
            f.write('\n')
            f.write('\n')
            #+++++++++++++++++Angular correction++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Adjust the angles++++++++++++++++++++++++++++++++++++++++
            f.write('Step 3. Adjust angles')
            f.write('\n')
            f.write('\n')
            AC_array = []
            for i in range(len(Internal_angles)):
                AC_array.append(AC)
            Adjusted_anlges = [x+y for x,y in zip(Internal_angles,AC_array)]
            Internal_anlge_DMS_Str = []
            for i in range(len(Internal_angles_rad)):
                Internal_anlge_DMS_Str.append(rad_to_dms_Str(Internal_angles_rad[i]))
            Adjusted_anlges_DMS_Str = []
            for i in range(len(Adjusted_anlges)):
                Adjusted_anlges_DMS_Str.append(DDtoDMS_Str(Adjusted_anlges[i]))
            for i in range(len(Adjusted_anlges)):
                sentence = temp_From_Stn[i] + ': ' + Internal_anlge_DMS_Str[i] + ' + (-' + AC_DMS + ') = ' + Adjusted_anlges_DMS_Str[i]
                f.write(sentence)
                f.write('\n')
            f.write('\n')
            f.write('\n')
            f.write('\n')
            #+++++++++++++++++Adjust the angles++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Compute Bearings++++++++++++++++++++++++++++++++++++++++
            f.write('Step 4. Compute Bearings ')
            f.write('\n')
            f.write('\n')
            Start_bearing_DMS_Str = Bearing_list[1]
            f.write('Bearing = previous Bearing ± 180° + angle between two points')
            f.write('\n')
            f.write('\n')
            sentence = 'START BEARING: ' + Start_bearing_DMS_Str
            f.write(sentence)
            f.write('\n')
            prev_bearing = dms_to_dd(Start_Bearing)
            bearings = []
            prev_bearings = []
            prev_bearings.append(prev_bearing)
            for i in range(len(Adjusted_anlges)):
                    bearing = prev_bearing + 180 + Adjusted_anlges[i]
                    if bearing < 0:
                        bearing = bearing + 360
                    elif bearing > 360: 
                        bearing = bearing - 360
                    elif bearing < -360: 
                        bearing = bearing + 720
                    elif bearing > 360: 
                        bearing = bearing - 720   
                    elif bearing < -720: 
                        bearing = bearing + 1080
                    elif bearing > 720: 
                        bearing = bearing - 1080
                    bearings.append(bearing)
                    prev_bearing = bearing
                    prev_bearings.append(prev_bearing)
            bearings_DMS_Str = []    
            for i in range(len(bearings)):
                bearings_DMS_Str.append(DDtoDMS_Str(bearings[i]))    
            prev_bearings_DMS_Str = []    
            for i in range(len(prev_bearings)):
                prev_bearings_DMS_Str.append(DDtoDMS_Str(prev_bearings[i]))
            temp_From_To = []
            for i in From_To:
                temp_From_To.append(i)
            temp_From_To.pop(0)
            temp_From_To.append('END   ')
            for i in range(len(temp_From_To)):
                sentence = temp_From_To[i] + ': ' + prev_bearings_DMS_Str[i] + ' ± 180° + ' + Adjusted_anlges_DMS_Str[i] + ' = ' + bearings_DMS_Str[i]
                f.write(sentence)
                f.write('\n')
            f.write('\n')
            Start_bearing_dd = dms_to_dd(Start_Bearing)
            if round(Start_bearing_dd, 4) == round(bearings[-1], 4):
                f.write('The START and the END Bearing are equal (or close too), as required.')
            else:
                f.write('THE START AND END BEARING ARE NOT EQUAL. CHECK DATA.')
            f.write('\n')
            f.write('\n')
            f.write('\n')
            #+++++++++++++++++Compute Bearings++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Compute ΔE & ΔN & ΔH++++++++++++++++++++++++++++++++++++++++
            f.write('Step 5. Compute the change in E, N & H')
            f.write('\n')
            f.write('\n')
            f.write('Delta E = HD x sin(Bearing)')
            f.write('\n')
            f.write('Dleta N = HD x cos(Bearing)')
            f.write('\n')
            f.write('Dleta H = SD x cos(Vertical angle) + (Hi - Ht)')
            f.write('\n')
            f.write('\n')
            bearings_rad = []
            for i in range(len(bearings)):
                x = dd_to_rad(i)
                bearings_rad.append(x)
            V_bearing_rad = []
            for i in range(len(V_bearing)):
                    x = dd_to_rad(i)
                    V_bearing_rad.append(x)   
            temp_bearings = []
            for i in bearings:
                temp_bearings.append(i)
            x = temp_bearings[-1]
            temp_bearings.pop(-1)
            bearings = [] 
            bearings.append(x)
            for i in temp_bearings:
                bearings.append(i)
            bearings_rad = []
            for i in bearings:
                bearings_rad.append(dd_to_rad(i))

            VBrad = []
            for i in range(len(V_bearing)):
                
                x = dd_to_rad(V_bearing[i])
                VBrad.append(x)

            DeltaE_list_unrounded = [x*math.sin(y) for x,y in zip(HD_red2,bearings_rad)]
            DeltaN_list_unrounded = [x*math.cos(y) for x,y in zip(HD_red2,bearings_rad)]
            DeltaH_list_unrounded = [x*math.cos(y) + (a - b) for x,y,a,b in zip(SD_red2,VBrad, Hi_red2, Ht_red2)]
            bearings_DMS_Str = []
            for i in range(len(bearings)):
                bearings_DMS_Str.append(DDtoDMS_Str(bearings[i]))
            bearings_DMS_Str = []
            for i in range(len(bearings)):
                bearings_DMS_Str.append(DDtoDMS_Str(bearings[i]))
            V_bearing_DMS_Str = []   
            for i in range(len(V_bearing)):
                V_bearing_DMS_Str.append(DDtoDMS_Str(V_bearing[i])) 
            DeltaE_list = []
            for i in DeltaE_list_unrounded:
                x = round(i, 3)
                DeltaE_list.append(x) 
            DeltaN_list = []
            for i in DeltaN_list_unrounded:
                x = round(i, 3)
                DeltaN_list.append(x) 
            DeltaH_list = []
            for i in DeltaH_list_unrounded:
                x = round(i, 3)
                DeltaH_list.append(x) 
            for i in range(len(DeltaE_list)): 
                sentence = From_To[i] + ': '
                f.write(sentence)
                f.write('\n')
                sentence = 'delta E = ' + str(HD_red2[i]) + ' x sin(' +  str(bearings_DMS_Str[i]) + ') = ' +  str(DeltaE_list[i])
                f.write(sentence)
                f.write('\n')
                sentence = 'delta N = ' + str(HD_red2[i]) + ' x cos(' +  str(bearings_DMS_Str[i]) + ') = ' +  str(DeltaN_list[i])
                f.write(sentence)
                f.write('\n')
                sentence = 'delta H = ' + str(round(SD_red2[i], 3)) + ' x cos(' + str(V_bearing_DMS_Str[i]) + ') + (' + str(Hi_red2[i]) + ' - '  + str(Hi_red2[i]) + ') = ' + str(DeltaH_list[i])
                f.write(sentence)
                f.write('\n')
                f.write('\n')
            f.write('\n')
            deltaE_sum = sum(DeltaE_list)
            deltaE_sum = round(deltaE_sum, 3)
            sentence = 'sum of delta E = ' + str(deltaE_sum)
            f.write(sentence)
            f.write('\n')
            deltaN_sum = sum(DeltaN_list)
            deltaN_sum = round(deltaN_sum, 3)
            sentence = 'sum of delta N = ' + str(deltaN_sum)
            f.write(sentence)
            f.write('\n')
            f.write('\n')
            f.write('\n')
            #+++++++++++++++++Compute ΔE & ΔN & ΔH++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Linear misclose++++++++++++++++++++++++++++++++++++++++
            f.write('Step 6. Compute the linear misclose (LM) and  proportional misclose (PM)')
            f.write('\n')
            f.write('\n')
            f.write('LM = sqrt((sum(delta E^2)) + (sum(delta N^2)))')
            f.write('\n')
            LM_m = math.sqrt((deltaE_sum**2)+(deltaN_sum**2)) #m
            LM_m = round(LM_m, 3) 
            LM_mm = LM_m*1000
            PM = sum(HD_red2)/LM_m
            sentence = 'LM = sqrt((' + str(deltaE_sum) + ') + (' + str(deltaN_sum) + '))'
            f.write(sentence)
            f.write('\n')
            sentence = 'LM = ' + str(LM_m) + 'm or ' + str(LM_mm) + 'mm'
            f.write(sentence)
            f.write('\n')
            f.write('\n')
            SUM_HD = sum(HD_red2)
            PM = SUM_HD/LM_m
            PM = round(PM, 3)
            acceptable = 0
            if PM > 15000:
                acceptable = 1    
            f.write('PM = sum(HD)/LM')
            f.write('\n')
            sentence = 'PM = ' + str(SUM_HD) + '/' + str(LM_m)
            f.write(sentence)
            f.write('\n')
            if acceptable == 1:
                sentence = 'PM = ' + str(PM) + ',   ' + str(PM) + ' > 15 000.'
                sentencetwo = 'Therefore the travers meets the 3rd order standards,' 
            else:
                sentence = 'PM = ' + str(PM) + ',   ' + str(PM) + ' < 15 000.'
                sentencetwo = 'Therefore the travers does not meet the 3rd order standards,' 
            f.write('\n')
            sentencethree = '(Standard for the Australian Survey Control Network Special Publication 1 (SP1) | Intergovernmental Committee on Surveying and Mapping, 2020).'
            f.write(sentence)
            f.write('\n')
            f.write(sentencetwo)
            f.write('\n')
            f.write(sentencethree)
            f.write('\n')
            f.write('\n')
            f.write('\n')
            f.write('\n')
            #+++++++++++++++++Linear misclose++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Cr delta E and Cr delta N++++++++++++++++++++++++++++++++++++++++
            f.write('Step 7. Compute the Corrections for delta E and delta N')
            f.write('\n')
            f.write('\n')
            f.write('Cr delta E = (-(HD)/sum(HD)) x mE, 	mE = sum(delta E)')
            f.write('\n')
            f.write('Cr delta N = (-(HD)/sum(HD)) x mN, 	mN = sum(delta N)')
            f.write('\n')
            f.write('\n')
            SUM_HD = sum(HD_red2)
            SUM_HD = round(SUM_HD, 3)
            mE = deltaE_sum
            mN = deltaN_sum
            Easting_Cr = []
            for i in range(len(HD_red2)):
                x = ((-1*HD_red2[i])/SUM_HD)*mE
                x = round(x, 5)
                Easting_Cr.append(x) 
            Northing_Cr = []
            for i in range(len(HD_red2)):
                x = ((-1*HD_red2[i])/SUM_HD)*mN
                x = round(x, 5)
                Northing_Cr.append(x)   
            for i in range(len(Northing_Cr)):
                sentence = From_To[i] + ': '
                f.write(sentence)
                f.write('\n')
                sentence = 'Cr delta E = (-(' +  str(HD_red2[i]) + ')/' + str(SUM_HD) + ') x ' + str(mE) + ' = ' + str(Easting_Cr[i])
                f.write(sentence)
                f.write('\n')
                sentence = 'Cr delta N = (-(' +  str(HD_red2[i]) + ')/' + str(SUM_HD) + ') x ' + str(mN) + ' = ' + str(Northing_Cr[i])
                f.write(sentence)
                f.write('\n')
                f.write('\n')
            f.write('\n')
            #+++++++++++++++++Cr delta E and Cr delta N++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Coordinates for E, N & H++++++++++++++++++++++++++++++++++++++++
            if PM >= 15000:
                f.write('E = delta E + Cr delta E + Coordinate at the previous point')
                f.write('\n')
                f.write('N = delta N + Cr delta N + Coordinate at the previous point')
                f.write('\n')
                f.write('H = delta H + Cr delta H + Coordinate at the previous point')
                f.write('\n')
                f.write('\n')
            else:
                f.write('E = delta E + Coordinate at the previous point')
                f.write('\n')
                f.write('N = delta N + Coordinate at the previous point')
                f.write('\n')
                f.write('H = delta H + Coordinate at the previous point')
                f.write('\n')
                f.write('\n')
            sentence = str(From_Stn_red2[0]) + '(Known):'
            f.write(sentence)
            f.write('\n')
            sentence = 'E: = ' + str(Start_Easting)
            f.write(sentence)
            f.write('\n')
            sentence = 'N: = ' + str(Start_Northing)
            f.write(sentence)
            f.write('\n')
            sentence = 'H: = ' + str(Start_Height)
            f.write(sentence)
            f.write('\n')
            f.write('\n')
            Easting = []
            Northing = []
            Height = []
            Easting.append(Start_Easting)
            Northing.append(Start_Northing)
            Height.append(Start_Height)
            last_DeltaE_list = []
            for i in DeltaE_list:
                last_DeltaE_list.append(i)
            last_DeltaE_list.pop(-1)
            last_DeltaN_list = []
            for i in DeltaN_list:
                last_DeltaN_list.append(i)
            last_DeltaN_list.pop(-1)
            last_DeltaH_list = []
            for i in DeltaH_list:
                last_DeltaH_list.append(i)
            last_DeltaH_list.pop(-1)
            last_Easting_Cr = []
            for i in Easting_Cr:
                last_Easting_Cr.append(i)
            last_Easting_Cr.pop(-1)
            last_Northing_Cr = []
            for i in Northing_Cr:
                last_Northing_Cr.append(i)
            last_Northing_Cr.pop(-1)
            last_From_Stn = []
            for i in From_Stn_red2:
                last_From_Stn.append(i)
            last_From_Stn.pop(0)
            E_coord_prev = Start_Easting
            N_coord_prev = Start_Northing
            H_coord_prev = Start_Height
            E_prevs = []
            E_prevs.append(Start_Easting)
            Eastings = []
            for i in range(len(last_DeltaE_list)):
                x = last_DeltaE_list[i] + last_Easting_Cr[i] + E_coord_prev
                x = round(x, 3)
                Eastings.append(x) 
                E_coord_prev = x
                E_prevs.append(E_coord_prev)
            N_prevs = []
            N_prevs.append(Start_Northing)
            Northings = []
            for i in range(len(last_DeltaN_list)):
                x = last_DeltaN_list[i] + last_Northing_Cr[i] + N_coord_prev
                x = round(x, 3)
                Northings.append(x) 
                N_coord_prev = x
                N_prevs.append(N_coord_prev)
            H_prevs = []
            H_prevs.append(Start_Height)
            Heights = []
            
            
            for i in range(len(DeltaH_list)):
                x = DeltaH_list[i] + H_coord_prev
            
            
            # for i in range(len(last_DeltaH_list)):
            #     x = last_DeltaH_list[i] + H_coord_prev
                
                
                x = round(x, 3)
                Heights.append(x) 
                H_coord_prev = x
                H_prevs.append(H_coord_prev)
            new_Eastings = []
            for i in E_prevs:
                new_Eastings.append(i)
            new_Northings = []
            for i in N_prevs:
                new_Northings.append(i)
            new_Heights = []
            for i in H_prevs:
                new_Heights.append(i)
            E_prevs.pop(-1)
            N_prevs.pop(-1)
            H_prevs.pop(-1)
            for i in range(len(last_DeltaE_list)):
                sentence = str(last_From_Stn[i]) + ': '
                f.write(sentence)
                f.write('\n')
                sentence = 'E = ' + str(last_DeltaE_list[i]) + ' + ' + str(last_Easting_Cr[i]) + ' + ' + str(E_prevs[i]) + ' = ' + str(Eastings[i])
                f.write(sentence)
                f.write('\n')
                sentence = 'N = ' + str(last_DeltaN_list[i]) + ' + ' + str(last_Northing_Cr[i]) + ' + ' + str(N_prevs[i]) + ' = ' + str(Northings[i])
                f.write(sentence)
                f.write('\n')
                sentence = 'H = ' + str(last_DeltaH_list[i]) + ' + ' + str(H_prevs[i]) + ' = ' + str(Heights[i])
                f.write(sentence)
                f.write('\n')
                f.write('\n')
            f.write('\n')
            f.write('\n')
            f.write('\n')
            f.close()
            #+++++++++++++++++Coordinates for E, N & H++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Plotting the data++++++++++++++++++++++++++++++++++++++++
            fig,ax1 = plt.subplots()
            Plot_E = []
            for i in new_Eastings:
                Plot_E.append(i)
            Plot_E.append(new_Eastings[0])
            Plot_N = []
            for i in new_Northings:
                Plot_N.append(i)
            Plot_N.append(new_Northings[0])
            x = Plot_E
            y = Plot_N
            line = ax1.plot(x, y,color='y', alpha=0.5, linewidth=2.2,label='Observation line',zorder=9)
            for i in range(len(new_Eastings)):
                plt.text(new_Eastings[i] + new_Eastings[i]*0.002 , new_Northings[i] + new_Northings[i]*0.002 , From_Stn_red2[i] + ', Height (m): ' + str(new_Heights[i]), fontsize = 5) 
            plt.title('Close Looped Traverse (Arbitrary Datum)', fontsize = 15)
            plt.xlabel('Eastings (m)', fontsize = 15)
            plt.ylabel('Northings (m)', fontsize = 15)
            ax1.scatter(new_Eastings, new_Northings, marker='x',s=80,color='black',alpha=1,label='Traverse Mark',zorder=10) 
            ax1.legend(loc='upper left')
            max_E = max(new_Eastings)
            min_E = min(new_Eastings)
            max_N = max(new_Northings)
            min_N = min(new_Northings)
            plt.xlim(min_E - min_E * 0.01, max_E + max_E * 0.035)
            plt.ylim(min_N - min_N * 0.01, max_N + max_N * 0.01)
            plt.savefig(Export_path + '/' + ExportFolder + "/Close Looped Traverse Plot.pdf")
            plt.show()
            plt.close()
            #+++++++++++++++++Plotting the data++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++Data excel table for the Bowditch adjustment++++++
            Adjusted_anlges_rad = []
            for i in Adjusted_anlges: 
                Adjusted_anlges_rad.append(dd_to_rad(i))
            Adjusted_anlges_list = []
            Adjusted_anlges_list.append(Slot)
            Adjusted_anlges_list.append(Slot)
            for i in range(len(From_Stn_red2)):
                Adjusted_anlges_list.append(rad_to_dms_Str(Internal_angles_rad[i]))
                Adjusted_anlges_list.append(Slot)
            Bearings_list_with_endB = []
            for i in bearings_DMS_Str:
                Bearings_list_with_endB.append(i)
            Bearings_list_with_endB.append(bearings_DMS_Str[0])
            Bearing_list = []
            for i in range(len(Bearings_list_with_endB)):
                Bearing_list.append(Slot)
                Bearing_list.append(Bearings_list_with_endB[i])
            Delta_E_list = []
            for i in range(len(DeltaE_list)):
                Delta_E_list.append(Slot)
                Delta_E_list.append(DeltaE_list[i])
            Delta_E_list.append(Slot)
            Delta_E_list.append(Slot)
            Delta_N_list = []
            for i in range(len(DeltaN_list)):
                Delta_N_list.append(Slot)
                Delta_N_list.append(DeltaN_list[i])
            Delta_N_list.append(Slot)
            Delta_N_list.append(Slot)
            Delta_H_list = []
            for i in range(len(DeltaH_list)):
                Delta_H_list.append(Slot)
                Delta_H_list.append(DeltaH_list[i])
            Delta_H_list.append(Slot)
            Delta_H_list.append(Slot)
            Easting_Cr_list = []
            for i in range(len(Easting_Cr)):
                Easting_Cr_list.append(Slot)
                Easting_Cr_list.append(Easting_Cr[i])
            Easting_Cr_list.append(Slot)
            Easting_Cr_list.append(Slot)
            Northing_Cr_list = []
            for i in range(len(Northing_Cr)):
                Northing_Cr_list.append(Slot)
                Northing_Cr_list.append(Northing_Cr[i])
            Northing_Cr_list.append(Slot)
            Northing_Cr_list.append(Slot)
            Easting_list = []
            for i in range(len(new_Eastings)):
                Easting_list.append(new_Eastings[i])
                Easting_list.append(Slot)
            Easting_list.append(Slot)
            Easting_list.append(Slot)
            Northing_list = []
            for i in range(len(new_Northings)):
                Northing_list.append(new_Northings[i])
                Northing_list.append(Slot)
            Northing_list.append(Slot)
            Northing_list.append(Slot)
            Height_list = []
            for i in range(len(new_Heights)):
                Height_list.append(new_Heights[i])
                Height_list.append(Slot)
            Height_list.append(Slot)
            Height_list.append(Slot)
            Bowditch_table = pd.DataFrame(list(zip(LineStation, Internal_anlge_list, Adjusted_anlges_list, Bearing_list, V_bearing_list, HD_list, SD_list, Delta_E_list, Easting_Cr_list, Delta_N_list, Northing_Cr_list, Delta_H_list, Easting_list, Northing_list, Height_list)),
                            columns =['Line/Station', 'Internal Angle',  'Adjusted Anlges (DMS)',  'Bearing (DMS)', 'Vertical Angle (DMS)', 'HD (m)', 'SD (m)', 'delta E (m)', 'Cr delta E (m)','delta N (m)', 'Cr delta N (m)', 'delta H (m)', 'E (m)', 'N (m)', 'H (m)'])
            file_name = Export_path + '/' + ExportFolder + '/Bowditch Table.xlsx'
            Bowditch_table.to_excel(file_name)
            #+++++++++++++++++Data excel table for the Bowditch adjustment++++++
            #+++++++++++++++++sending the data++++++++++++++++++++++++++++++++++++++++
            zip_path = directory + '/Exports'
            email_path = directory + "/Exports.zip"
            def get_all_file_paths(directory):
                # initializing empty file paths list
                file_paths = []
                # crawling through directory and subdirectories
                for root, directories, files in os.walk(directory):
                    for filename in files:
                        # join the two strings in order to form the full filepath.
                        filepath = os.path.join(root, filename)
                        file_paths.append(filepath)
                # returning all file paths
                return file_paths        
            def main(): ##CODE USED from user: nkmk.me found at: https://note.nkmk.me/en/python-zipfile/#:~:text=To%20compress%20individual%20files%20into,'w'%20(write).
                # path to folder which needs to be zipped
                directory = zip_path
                # calling function to get all file paths in the directory
                file_paths = get_all_file_paths(directory)
                # printing the list of all files to be zipped
                print('Following files will be zipped:')
                for file_name in file_paths:
                    print(file_name)
                # writing files to a zipfile
                with ZipFile(email_path,'w') as zip:
                    # writing each file one by one
                    for file in file_paths:
                        zip.write(file)
                print('All files zipped successfully!')        
            if __name__ == "__main__":
                main()
            email = Email[-1]
            password = Password[-1]
            msg=EmailMessage() #CODE USED FROM Joska de Langen found at: https://realpython.com/python-send-email/
            msg['Subject'] = 'Traverse Survey Reductions'
            msg['From']='Xavier Ryan (with love)'
            msg['To']=email
            email_path = directory + "/Exports.zip"
            with open(email_path, "rb") as f:
                file_data=f.read()
                #print("File data in binary", file_data)
                file_name=f.name
                #print("File name is", file_name)
                msg.add_attachment(file_data,maintype="application",subtype="zip",filename=file_name)
            with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
                server.login(email, password)
                server.send_message(msg)
            print('Email Sent!!!')
            #+++++++++++++++++sending the data++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++++++++Deleting the temporery exports folder++++++++++++++
            Export_folder_location = Export_path + '/' + ExportFolder 
            Reduced_table_data_file_location = Export_path + '/' + ExportFolder + '/Reduced Table.xlsx'
            Working_out_file_location = Export_path + '/' + ExportFolder + '/Working out.txt'
            Zip_file_location = Export_path + '/Exports.zip'
            os.unlink(Export_path + '/' + ExportFolder + '/Bowditch Table.xlsx')
            os.unlink(Export_path + '/' + ExportFolder + "/Close Looped Traverse Plot.pdf")
            os.unlink(Zip_file_location)
            os.unlink(Reduced_table_data_file_location)
            os.unlink(Working_out_file_location)
            os.rmdir(Export_folder_location)
            #+++++++++++++++++++++++Deleting the temporery exports folder++++++++++++++
        self.Folder_line_trav = QLineEdit(self)
        self.tab2.layout.addRow(self.Folder_line_trav)
        self.tab2.setLayout(self.tab1.layout)
        self.File_name = QLabel('What is the excel file called?: ', self)
        self.tab2.layout.addRow(self.File_name)
        self.tab2.setLayout(self.tab2.layout)
        self.File_line_trav = QLineEdit(self)
        self.tab2.layout.addRow(self.File_line_trav)
        self.tab2.setLayout(self.tab1.layout)
        self.Submit_all_email = QPushButton("Process")
        self.tab2.layout.addRow(self.Submit_all_email)
        self.tab2.setLayout(self.tab1.layout)
        self.Submit_all_email.clicked.connect(Submit_Folder_location)
        #TAB3
        self.tab3.layout = QFormLayout(self)
        self.Folder_Location = QLabel("Where is the excel file? (copy and past file path here):", self)
        self.tab3.layout.addRow(self.Folder_Location)
        self.tab3.setLayout(self.tab3.layout)
        def Submit_Folder_location_detail():
            a = self.Folder_line.text()
            a = a.replace("\\", "/")
            b = self.File_line.text()
            b = '/' + b + '.xlsx'
            path_and_file = a + b
            a = self.Folder_line.clear()
            b = self.File_line.clear()
            print('The location of the file is: ', path_and_file) 
            import pandas as pd
            import numpy as np
            import math
            import matplotlib.pyplot as plt
            from mpl_toolkits import mplot3d
            import matplotlib.pyplot as plt
            import os
            from zipfile import ZipFile
            from email.message import EmailMessage
            import smtplib
            #+++++++++++++++++++++++ALL IMPORTS ARE HERE+++++++++++++++++++++++++++++
            #+++++++++++++++++++++++ALL DEFINITIONS ARE HERE+++++++++++++++++++++++++++++
            def dms_to_dd(bearing):
                bearing = format(bearing, ".4f")
                bearing_str = str(bearing)
                bearing_deg_str = bearing_str[:-5]
                bearing_min_str = bearing_str[-4:-2]
                bearing_sec_str = bearing_str[-2:]
                bearing_deg = float(bearing_deg_str)
                bearing_min = float(bearing_min_str)
                bearing_sec = float(bearing_sec_str)
                DMS = bearing_deg + (float(bearing_min)/60) + (float(bearing_sec)/3600)
                return DMS
            def Horizontal_FLred(FL, FR):
                temp_face = FL - FR
                if temp_face < 0:
                    temp_face = temp_face + 180
                else: 
                    temp_face = temp_face - 180
                Hz_FLred = FL - 0.5*temp_face
                return Hz_FLred #
            def Vertical_FLred(FL, FR):
                if FL > 180.0000:
                    print("FL was likey observed in FR orientation. Please review the data")
                temp_face = FL + FR - 360
                V_FLred = FL - 0.5*temp_face
                return V_FLred
            def dd_to_rad(bearing):
                radian = bearing * 3.141592653589793/180
                return radian
            def rad_to_dms(bearing):
                radian = bearing * 3.141592653589793/180
                return radian
            def decdeg2dms(dd):
               is_positive = dd >= 0
               dd = abs(dd)
               minutes,seconds = divmod(dd*3600,60)
               degrees,minutes = divmod(minutes,60)
               degrees = degrees if is_positive else -degrees
               #seconds = round(seconds, 0)
               return (degrees,minutes,seconds)
            def TupleToDMS(tupleDMS):
                degrees = str(int(tupleDMS[0]))
                degrees = degrees + '° '
                minutes = str(int(tupleDMS[1]))
                minutes = (minutes + "' ")
                seconds = str(int(tupleDMS[2]))
                seconds = seconds + '"'
                DMS = degrees + minutes + seconds
                return DMS
            def dmsP_to_dmsStr(bearing): 
                dd = dms_to_dd(bearing)
                x = decdeg2dms(dd)
                DMS_Str = TupleToDMS(x)
                return DMS_Str
            def rad_to_dms_Str(bearing): #This function turns radiance into DMS (STRING FORMAT)
                dd = bearing * 180/3.141592653589793
                dms_tuple = decdeg2dms(dd)
                dms_str = TupleToDMS(dms_tuple)
                return dms_str
            def  DMStoSEC(DMS): #input must be in dd
                x = decdeg2dms(DMS)
                degrees = x[0]*3600
                minutes = x[1] * 60
                Seconds = degrees + minutes + x[2]
                Seconds = round(Seconds, 0)
                Seconds = int(Seconds)
                Seconds = str(Seconds) + '"'
                return Seconds #Output in seconds, type str
            def  DDtoDMS_Str(DMS): #input must be in dd
                x = decdeg2dms(DMS)
                DMS_Str = TupleToDMS(x)
                return DMS_Str #Output in seconds, type str
            #+++++++++++++++++++++++ALL DEFINITIONS ARE HERE+++++++++++++++++++++++++++++
            #Here we are importing the detail table that contains the locations data of existing or established points as well as the observations data from the detail
            #+++++++++++++++++++++++Specifiing a directory++++++++++++++++++++++++++
            directory = os.getcwd()
            # print('Directory: ', directory)
            #+++++++++++++++++++++++Specifiing a directory++++++++++++++++++++++++++
            #+++++++++++++++++++++++Creating a temporery export location++++++++++++++
            Export_path = directory.replace("\\", "/")
            os.chdir(Export_path)
            ExportFolder = 'Exports'
            os.makedirs(ExportFolder)
            #+++++++++++++++++++++++Creating a temporery export location++++++++++++++                       
            Coordinates_table = pd.read_excel (path_and_file, sheet_name='Coordinates')
            Observations_table = pd.read_excel (path_and_file, sheet_name='Observations')
            #+++++++++++++++++++Importing all the data+++++++++++++++++++++++++++++++++++++++
            Point_at_Stn_list_inport = Observations_table['Station'].tolist()
            Set_list_inport = Observations_table['Set'].tolist()
            Face_list_inport = Observations_table['Face'].tolist()
            Point_ID_list_inport = Observations_table['Point ID'].tolist()
            String_list_inport = Observations_table['String'].tolist()
            String_red = String_list_inport[::2]
            String_red2 = String_red[::2]
            HzBearing_list_inport = Observations_table['Horizontal Bearing (DMS)'].tolist()
            VzBearing_list_inport = Observations_table['Vertical Bearing (DMS)'].tolist()
            HzDist_list_inport = Observations_table['Horizontal distance (m)'].tolist()
            SlDist_list_inport = Observations_table['Slope distance (m)'].tolist()
            Hi = Observations_table['Height of Instrument (m)'].tolist()
            Hi_red = Hi[::2]
            Hi_red2 = Hi_red[::2]
            Ht = Observations_table['Height of Target (m)'].tolist()
            Ht_red = Ht[::2]
            Ht_red2 = Ht_red[::2]
            #+++++++++++++++++++Importing all the data+++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++++Converting data++++++++++++++++++++++++++++++++++++++++++++++
            #Convert the horizontal bearings from decimal degrees to DMS
            Hz_bearing_list = []
            for i in HzBearing_list_inport: 
                Hz_bearing_list.append(dms_to_dd(i))
            #Convert the vertical bearings from decimal degrees to DMS
            Vz_bearing_list = []
            for i in VzBearing_list_inport: 
                Vz_bearing_list.append(dms_to_dd(i))
            #+++++++++++++++++++Converting data++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++++Averaging all the data+++++++++++++++++++++++++++++++++++++++
            first_value = 0
            second_value = 1
            FLred_Hz = []
            for i in Hz_bearing_list:
                if second_value <= len(Hz_bearing_list):
                    FL = Hz_bearing_list[first_value]
                    FR = Hz_bearing_list[second_value]
                    temp = Horizontal_FLred(FL, FR)
                    FLred_Hz.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            FLred_Vz = []
            for i in Vz_bearing_list:
                if second_value <= len(Vz_bearing_list):
                    FL = Vz_bearing_list[first_value]
                    FR = Vz_bearing_list[second_value]
                    temp = Vertical_FLred(FL, FR)
                    FLred_Vz.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            Hz_dist = []
            for i in HzDist_list_inport:
                if second_value <= len(HzDist_list_inport):
                    dist1 = HzDist_list_inport[first_value]
                    dist2 = HzDist_list_inport[second_value]
                    temp = (dist1 + dist2)/2
                    Hz_dist.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            Sl_dist = []
            for i in SlDist_list_inport:
                if second_value <= len(SlDist_list_inport):
                    dist1 = SlDist_list_inport[first_value]
                    dist2 = SlDist_list_inport[second_value]
                    temp = (dist1 + dist2)/2
                    Sl_dist.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            #+++++++++++++++++++Reducing the data to a new table+++++++++++++++++++++++++++++++++++++++
            new_Stn = Point_at_Stn_list_inport[::2]
            Stn = new_Stn[::2]
            new_Set_list_inport = Set_list_inport[::2]
            new_Face_list = Face_list_inport[::2]
            for i in new_Face_list:
                i = new_Face_list.index('FL')
                new_Face_list = new_Face_list[:i]+['FLred']+new_Face_list[i+1:]
            new_Point_ID_list = Point_ID_list_inport[::2]
            Point_ID = []
            for i in new_Point_ID_list:
                if i not in Point_ID:
                    Point_ID.append(i)
            new_String_list = String_list_inport[::2]
            #+++++++++++++++++++Reducing the data to a new table+++++++++++++++++++++++++++++++++++++++
            #   Now that we have reduced the table we need to find any point ID 
            #   duplicates. From here we can Identifiy where they are 
            #   and average the bearings and distances baced on where they are located
            #   in their respective lists.
            # Calling DataFrame constructor after zipping
            # both lists, with columns specified
            Hz_DMS_Str = []
            for i in FLred_Hz:
                i = DDtoDMS_Str(i)
                Hz_DMS_Str.append(i)
            V_DMS_Str = []
            for i in FLred_Vz:
                i = DDtoDMS_Str(i)
                V_DMS_Str.append(i)
            reduced_table1 = pd.DataFrame(list(zip(new_Stn, new_Set_list_inport,  new_Face_list, new_Point_ID_list, String_red, Hz_DMS_Str, V_DMS_Str, Hz_dist, Sl_dist, Hi_red, Ht_red)),
                            columns =['Station', 'Set', 'Face', 'Point ID', 'string', 'Horizontal Bearing (DMS)', 'Vertical Bearing (DMS)', 'Horizontal distance (m)', 'Slope distance (m)', 'Height of Instrument (m)', 'Height of Target (m)'])
            #+++++++++++++++++++Reducing the data to a new table+++++++++++++++++++++++++++++++++++++++
            #++++++++++++++++++Identifying where the repective point IDs are++++++++++++++++++++++++++
            points = []
            for i in Point_ID_list_inport:
                if i not in points:
                    points.append(i)
            counter = new_Point_ID_list.count(new_Point_ID_list[0])
            position_list = []
            list_of_positions = []
            temp_list = new_Point_ID_list
            for i in points:
                point = i
                position = temp_list.index(point)
                for k in temp_list:
                    if point in temp_list:
                        index = temp_list.index(point)
                        position_list.append(index)
                        temp_list[index] = ''
                        list_of_positions.append(position_list)
                    else:
                        break
            list_of_positions = list_of_positions[0]
            def divide_chunks(l, n):
                # looping till length l
                for i in range(0, len(l), n):
                    yield l[i:i + n]
            #Definition sourced from https://www.geeksforgeeks.org/break-list-chunks-size-n-python/
            list_of_positions = list(divide_chunks(list_of_positions, counter))
            #++++++++++++++++++FINAL REDUCTION FOR THE VALUES++++++++++++++++
            def final_reduction(a):
                reduced_values = []
                locations = []
                average_point = []
                bounds = len(list_of_positions[0])
                for i in list_of_positions:
                    locations = i
                    for i in locations: 
                        average_point.append(a[i])
                average_values = list(divide_chunks(average_point, bounds))     
                for i in average_values:
                    point_average = sum(i)/len(i)
                    reduced_values.append(point_average)
                return reduced_values
            Hz_reduced = list(final_reduction(FLred_Hz))
            Vz_reduced = list(final_reduction(FLred_Vz))
            Hz_dist_reduced = list(final_reduction(Hz_dist))
            Sl_dist_reduced = list(final_reduction(Sl_dist))
            Stn = new_Stn[::2]
            Face_list = new_Face_list[::2]
            String_list = new_String_list[::2]
            Hz_DMS_Str2 = []
            for i in Hz_reduced:
                i = DDtoDMS_Str(i)
                Hz_DMS_Str2.append(i)
            V_DMS_Str2 = []
            for i in Hz_reduced:
                i = DDtoDMS_Str(i)
                V_DMS_Str2.append(i)
            reduced_table2 = pd.DataFrame(list(zip(Stn, Face_list, Point_ID, String_red2, Hz_DMS_Str2, V_DMS_Str2, Hz_dist_reduced, Sl_dist_reduced, Hi_red2, Ht_red2)),
                            columns =['Station', 'Face', 'Point ID', 'String', 'Horizontal Bearing (DMS)', 'Vertical Bearing (DMS)', 'Horizontal distance (m)', 'Slope distance (m)', 'Height of Instrument (m)', 'Height of Target (m)'])
            print(reduced_table2)
            file_name = Export_path + '/' + ExportFolder + '/Reduced table.xlsx'
            reduced_table2.to_excel(file_name)
            def dd_to_rad(bearing):
                radian = bearing * math.pi()/180
                return radian
            Face = Face_list
            String = String_list
            Hz_Bearing_deg = Hz_reduced
            Hz_Bearing_rad = []
            for i in Hz_Bearing_deg:
                y = math.radians(i)
                Hz_Bearing_rad.append(y)
            V_Bearing_deg = Vz_reduced
            V_Bearing_rad = []
            for i in V_Bearing_deg:
                y = math.radians(i)
                V_Bearing_rad.append(y)
            Hz_Distance = Hz_dist_reduced
            Sl_Distance = Sl_dist_reduced
            ΔE_list = [x*math.sin(y) for x,y in zip(Hz_Distance,Hz_Bearing_rad)]
            ΔN_list = [x*math.cos(y) for x,y in zip(Hz_Distance,Hz_Bearing_rad)]
            Hi_red2.pop(-1)
            Ht_red2.pop(-1)
            ΔH_list = [x*math.cos(y) + (a - b) for x,y,a,b in zip(Sl_Distance,V_Bearing_rad, Hi_red2, Ht_red2)]
            Stations_fixed = Coordinates_table['Coordinated Points'].tolist()
            Easting_fixed = Coordinates_table['Easting (m)'].tolist()
            Northing_fixed = Coordinates_table['Northing (m)'].tolist()
            Height_fixed = Coordinates_table['Height (m)'].tolist()
            Stations_count_dict = {i:Stn.count(i) for i in Stn}
            dict = Stations_count_dict
            res_list = dict.values()
            Stations_count = list(res_list)
            print(Stations_count)
            temp_df = pd.DataFrame({
                'Station' : Stations_fixed,
                'Easting' : Easting_fixed,
                'Northing': Northing_fixed,
                'Height' :  Height_fixed,
                'StationCount' : Stations_count
                })
            coord_to_obs = (temp_df.loc[temp_df.index.repeat(temp_df.StationCount)].reset_index(drop=True))
            Easting_station = coord_to_obs['Easting'].tolist()
            Northing_station = coord_to_obs['Northing'].tolist()
            Height_station = coord_to_obs['Height'].tolist()
            Easting_unrounded = [x+y for x,y in zip(Easting_station,ΔE_list)]
            Easting = [round(num, 3) for num in Easting_unrounded]
            Northing_unrounded = [x+y for x,y in zip(Northing_station,ΔN_list)]
            Northing = [round(num, 3) for num in Northing_unrounded]
            Height_unrounded = [x+y for x,y in zip(Height_station,ΔH_list)]
            Height = [round(num, 3) for num in Height_unrounded]
            plt.scatter(Easting, Northing)
            plt.show()
            coords = pd.DataFrame(list(zip(Point_ID, Easting, Northing, Height)),
                            columns =['Point ID', 'Easting (m)', 'Northing (m)', 'Height.ortho (m)'])
            print(coords)
            file_name = Export_path + '/' + ExportFolder + '/Coordinates.xlsx'
            coords.to_excel(file_name)
            #+++++++++++++++++sending the data++++++++++++++++++++++++++++++++++++++++
            zip_path = directory + '/Exports'
            email_path = directory + "/Exports.zip"
            def get_all_file_paths(directory):
                # initializing empty file paths list
                file_paths = []
                # crawling through directory and subdirectories
                for root, directories, files in os.walk(directory):
                    for filename in files:
                        # join the two strings in order to form the full filepath.
                        filepath = os.path.join(root, filename)
                        file_paths.append(filepath)
                # returning all file paths
                return file_paths        
            def main(): ##CODE USED from user: nkmk.me found at: https://note.nkmk.me/en/python-zipfile/#:~:text=To%20compress%20individual%20files%20into,'w'%20(write).
                # path to folder which needs to be zipped
                directory = zip_path
                # calling function to get all file paths in the directory
                file_paths = get_all_file_paths(directory)
                # printing the list of all files to be zipped
                print('Following files will be zipped:')
                for file_name in file_paths:
                    print(file_name)
                # writing files to a zipfile
                with ZipFile(email_path,'w') as zip:
                    # writing each file one by one
                    for file in file_paths:
                        zip.write(file)
                print('All files zipped successfully!')        
            if __name__ == "__main__":
                main()
            # #For this to work you need to set up an app sepecific password
            # #Do this by going to https://myaccount.google.com/security > Signing in to Google > App passwords > 
            # #Select app > Other (Custom name) > "Python" > Generate
            email = Email[-1]
            password = Password[-1]
            msg=EmailMessage() #CODE USED FROM Joska de Langen found at: https://realpython.com/python-send-email/
            msg['Subject'] = 'Detail Survey Reductions'
            msg['From']='Xavier Ryan (with love)'
            msg['To']=email
            email_path = directory + "/Exports.zip"
            with open(email_path, "rb") as f:
                file_data=f.read()
                #print("File data in binary", file_data)
                file_name=f.name
                #print("File name is", file_name)
                msg.add_attachment(file_data,maintype="application",subtype="zip",filename=file_name)
            with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
                server.login(email, password)
                server.send_message(msg)
            print('Email Sent!!!')
            #+++++++++++++++++sending the data++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++++++++Deleting the temporery exports folder++++++++++++++
            Export_folder_location = Export_path + '/' + ExportFolder 
            Reduced_table_data_file_location = Export_path + '/' + ExportFolder + '/Reduced Table.xlsx'
            Coord_data_file_location = Export_path + '/' + ExportFolder + '/Coordinates.xlsx'
            Zip_file_location = Export_path + '/Exports.zip'
            os.unlink(Reduced_table_data_file_location)
            os.unlink(Coord_data_file_location)
            os.unlink(Zip_file_location)
            os.rmdir(Export_folder_location)     
        self.Folder_line = QLineEdit(self)
        self.tab3.layout.addRow(self.Folder_line)
        self.tab3.setLayout(self.tab1.layout)
        self.File_name = QLabel('What is the excel file called?: ', self)
        self.tab3.layout.addRow(self.File_name)
        self.tab3.setLayout(self.tab3.layout)
        self.File_line = QLineEdit(self)
        self.tab3.layout.addRow(self.File_line)
        self.tab3.setLayout(self.tab1.layout)
        self.Submit_all_email = QPushButton("Process")
        self.tab3.layout.addRow(self.Submit_all_email)
        self.tab3.setLayout(self.tab1.layout)
        self.Submit_all_email.clicked.connect(Submit_Folder_location_detail)
        #tab4
        # Create first tab
        self.tab4.layout = QFormLayout(self)
        self.Station_label = QLabel("Station: ", self)
        self.tab4.layout.addRow(self.Station_label)
        self.tab4.setLayout(self.tab4.layout)
        def Submit_Station():
            a = self.Station_line.text()
            Station.append(a)
            a = self.Station_line.clear()
            x = Station[-1]
            print('Station: ', x)
        self.Station_line = QLineEdit(self)
        self.tab4.layout.addRow(self.Station_line)
        self.tab4.setLayout(self.tab4.layout)
        self.Easting_label = QLabel("Easting (m): ", self)
        self.tab4.layout.addRow(self.Easting_label)
        self.tab4.setLayout(self.tab4.layout)
        def Submit_Easting():
            a = self.Easting_line.text()
            a = float(a)
            Easting.append(a)
            a = self.Easting_line.clear()
        self.Easting_line = QLineEdit(self)
        self.tab4.layout.addRow(self.Easting_line)
        self.tab4.setLayout(self.tab4.layout)
        self.Northing_label = QLabel("Northing (m): ", self)
        self.tab4.layout.addRow(self.Northing_label)
        self.tab4.setLayout(self.tab4.layout)
        def Submit_Northing():
            a = self.Northing_line.text()
            a = float(a)
            Northing.append(a)
            a = self.Northing_line.clear()
        self.Northing_line = QLineEdit(self)
        self.tab4.layout.addRow(self.Northing_line)
        self.tab4.setLayout(self.tab4.layout)
        self.Height_label = QLabel("Height (m): ", self)
        self.tab4.layout.addRow(self.Height_label)
        self.tab4.setLayout(self.tab4.layout)
        def Submit_Height():
            a = self.Height_line.text()
            a = float(a)
            Height.append(a)
            a = self.Height_line.clear()
        self.Height_line = QLineEdit(self)
        self.tab4.layout.addRow(self.Height_line)
        self.tab4.setLayout(self.tab4.layout)
        self.Submit_all_Coordinates = QPushButton("Submit all")
        self.tab4.layout.addRow(self.Submit_all_Coordinates)
        self.tab4.setLayout(self.tab4.layout)
        self.Submit_all_Coordinates.clicked.connect(Submit_Station)
        self.Submit_all_Coordinates.clicked.connect(Submit_Easting)
        self.Submit_all_Coordinates.clicked.connect(Submit_Northing)
        self.Submit_all_Coordinates.clicked.connect(Submit_Height)
        #tab5
        # user input
        self.tab5.layout = QFormLayout(self)
        self.Stn_label = QLabel("Station: ", self)
        self.tab5.layout.addRow(self.Stn_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_Stn():
            a = self.Stn_line.text()
            Point_at_Stn_list_inport.append(a)
            x = Point_at_Stn_list_inport[-1]
            print('Station: ', x)
        self.Stn_line = QLineEdit(self)
        self.tab5.layout.addRow(self.Stn_line)
        self.tab5.setLayout(self.tab5.layout)
        self.Set_label = QLabel("Set: ", self)
        self.tab5.layout.addRow(self.Set_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_Set():
            a = self.Set_line.text()
            a = int(a)
            Set_list_inport.append(a)
        self.Set_line = QLineEdit(self)
        self.tab5.layout.addRow(self.Set_line)
        self.tab5.setLayout(self.tab5.layout)
        self.Face_label = QLabel("Face: ", self)
        self.tab5.layout.addRow(self.Face_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_Face():
            a = self.Face_line.text()
            a = a.upper()
            Face_list_inport.append(a)
            a = self.Face_line.clear()
            x = Face_list_inport[-1]
            print('Face: ', x)
        self.Face_line = QLineEdit(self)
        self.tab5.layout.addRow(self.Face_line)
        self.tab5.setLayout(self.tab5.layout)
        self.PointID_label = QLabel("PointID: ", self)
        self.tab5.layout.addRow(self.PointID_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_PointID():
            a = self.PointID_line.text()
            Point_ID_list_inport.append(a)
            a = self.PointID_line.clear()
            x = Point_ID_list_inport[-1]
            print('PointID: ', x)
        self.PointID_line = QLineEdit(self)
        self.tab5.layout.addRow(self.PointID_line)
        self.tab5.setLayout(self.tab5.layout)
        self.String_label = QLabel("String: ", self)
        self.tab5.layout.addRow(self.String_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_String():
            a = self.String_line.text()
            String_list_inport.append(a)
            a = self.String_line.clear()
        self.String_line = QLineEdit(self)
        self.tab5.layout.addRow(self.String_line)
        self.tab5.setLayout(self.tab5.layout)
        self.Hz_label = QLabel("Hz (dd.mmss): ", self)
        self.tab5.layout.addRow(self.Hz_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_Hz():
            a = self.Hz_line.text()
            a = float(a)
            a = format(a, ".4f")
            HzBearing_list_inport.append(a)
            a = self.Hz_line.clear()
        self.Hz_line = QLineEdit(self)
        self.tab5.layout.addRow(self.Hz_line)
        self.tab5.setLayout(self.tab5.layout)
        self.V_label = QLabel("V (dd.mmss): ", self)
        self.tab5.layout.addRow(self.V_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_V():
            a = self.V_line.text()
            a = float(a)
            a = format(a, ".4f")
            VzBearing_list_inport.append(a)
            a = self.V_line.clear()
        self.V_line = QLineEdit(self)
        self.tab5.layout.addRow(self.V_line)
        self.tab5.setLayout(self.tab5.layout)
        self.HD_label = QLabel("HD (m): ", self)
        self.tab5.layout.addRow(self.HD_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_HD():
            a = self.HD_line.text()
            a = float(a)
            HzDist_list_inport.append(a)
            a = self.HD_line.clear()
        self.HD_line = QLineEdit(self)
        self.tab5.layout.addRow(self.HD_line)
        self.tab5.setLayout(self.tab5.layout)    
        self.SD_label = QLabel("SD (m): ", self)
        self.tab5.layout.addRow(self.SD_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_SD():
            a = self.SD_line.text()
            a = float(a)
            SlDist_list_inport.append(a)
            a = self.SD_line.clear()
        self.SD_line = QLineEdit(self)
        self.tab5.layout.addRow(self.SD_line)
        self.tab5.setLayout(self.tab5.layout)
        self.Hi_label = QLabel("Hi (m): ", self)
        self.tab5.layout.addRow(self.Hi_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_Hi():
            a = self.Hi_line.text()
            a = float(a)
            Hi_xyz.append(a)
        self.Hi_line = QLineEdit(self)
        self.tab5.layout.addRow(self.Hi_line)
        self.tab5.setLayout(self.tab5.layout)
        self.Ht_label = QLabel("Ht (m): ", self)
        self.tab5.layout.addRow(self.Ht_label)
        self.tab5.setLayout(self.tab5.layout)
        def Submit_Ht():
            a = self.Ht_line.text()
            a = float(a)
            Ht_xyz.append(a)
        self.Ht_line = QLineEdit(self)
        self.tab5.layout.addRow(self.Ht_line)
        self.tab5.setLayout(self.tab5.layout)
        self.Submit_all_Observations = QPushButton("Submit Observation")
        self.tab5.layout.addRow(self.Submit_all_Observations)
        self.tab5.setLayout(self.tab5.layout)
        self.Submit_all_Observations.clicked.connect(Submit_Stn)
        self.Submit_all_Observations.clicked.connect(Submit_Set)
        self.Submit_all_Observations.clicked.connect(Submit_Face)
        self.Submit_all_Observations.clicked.connect(Submit_PointID)
        self.Submit_all_Observations.clicked.connect(Submit_String)
        self.Submit_all_Observations.clicked.connect(Submit_Hz)
        self.Submit_all_Observations.clicked.connect(Submit_V)
        self.Submit_all_Observations.clicked.connect(Submit_HD)
        self.Submit_all_Observations.clicked.connect(Submit_SD)
        self.Submit_all_Observations.clicked.connect(Submit_Hi)
        self.Submit_all_Observations.clicked.connect(Submit_Ht)
        #TAB 6
        self.tab6.layout = QFormLayout(self)
        def ClearData():
            #Email Tab
            Email = []
            print(Email)
            Password = []
            #Stn coodinates tab
            Station = []
            Easting = []
            Northing = []
            Height = []
            #Observations tab
            Point_at_Stn_list_inport = []
            Set_list_inport = []
            Face_list_inport = []
            Point_ID_list_inport = []
            String_list_inport = []
            HzBearing_list_inport = []
            VzBearing_list_inport = []
            HzDist_list_inport = []
            SlDist_list_inport = []
            Hi = []
            print('data has been cleared')
        self.ClearData = QPushButton("Clear Data")
        self.tab6.layout.addRow(self.ClearData)
        self.tab6.setLayout(self.tab6.layout)
        self.ClearData.clicked.connect(ClearData)
        self.Warning = QLabel("Clearing all data will get rid of all observatins made")
        self.tab6.layout.addRow(self.Warning)
        self.tab6.setLayout(self.tab6.layout)
        def Process():
            import pandas as pd
            import numpy as np
            import math
            import matplotlib.pyplot as plt
            from mpl_toolkits import mplot3d
            import matplotlib.pyplot as plt
            import os
            from zipfile import ZipFile
            import smtplib
            from email.message import EmailMessage
            #+++++++++++++++++++++++ALL DEFINITIONS ARE HERE+++++++++++++++++++++++++++++
            #This function is used to convert decimal degrees (format: ddd.mmss) into degrees minutes seconds 
            #This function will be called upon when converting all of the bearing in the table to DMS
            def dms_to_dd(bearing):
                bearing_deg_str = bearing[:-5]
                bearing_min_str = bearing[-4:-2]
                bearing_sec_str = bearing[-2:]
                bearing_deg = float(bearing_deg_str)
                bearing_min = float(bearing_min_str)
                bearing_sec = float(bearing_sec_str)
                DMS = bearing_deg + (float(bearing_min)/60) + (float(bearing_sec)/3600)
                return DMS
            def Horizontal_FLred(FL, FR):
                temp_face = FL - FR
                if temp_face < 0:
                    temp_face = temp_face + 180
                else: 
                    temp_face = temp_face - 180
                Hz_FLred = FL - 0.5*temp_face
                return Hz_FLred #
            def Vertical_FLred(FL, FR):
                if FL > 180.0000:
                    print("FL was likey observed in FR orientation. Please review the data")
                temp_face = FL + FR - 360
                V_FLred = FL - 0.5*temp_face
                return V_FLred
            def dd_to_rad(bearing):
                radian = bearing * math.pi()/180
                return radian
            #+++++++++++++++++++++++ALL DEFINITIONS ARE HERE+++++++++++++++++++++++++++++
            import os
            directory = os.getcwd()
            print('Directory: ', directory)
            #++++++++++++++++++++++determining where in import and export the files too++++++++++++++++++++++++
            #Export location
            Export_path = directory
            Export_path = Export_path.replace("\\", "/")
            os.chdir(Export_path)
            ExportFolder = 'Exports'
            os.makedirs(ExportFolder)
            RAW_data_file_location = Export_path + '/' + ExportFolder + '/RAW DATA.xlsx'
            COORD_data_file_location = Export_path + '/' + ExportFolder + '/COORD DATA.xlsx'
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++Appending the coordinates+++++++++++
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            RAWCOORD = pd.DataFrame(list(zip(Station, Easting, Northing, Height)),
                            columns =['Station', 'Easting (m)', 'Northing (m)', 'Height (m)'])
            file_name = COORD_data_file_location
            RAWCOORD.to_excel(file_name)
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++Appending the coordinates+++++++++++
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++Appending the observations++++++++++
            #+++++++++++++++++++++++++++++++++++++++++++++++++

            frame = pd.DataFrame(list(zip(Point_at_Stn_list_inport, Set_list_inport, Face_list_inport, Point_ID_list_inport, String_list_inport, HzBearing_list_inport, VzBearing_list_inport, HzDist_list_inport, SlDist_list_inport, Hi_xyz, Ht_xyz)),
                            columns =['Station', 'Set', 'Face', 'Point ID', 'String', 'Hz', 'V', 'HD (m)', 'SD (m)', 'Hi (m)', 'Hi (m)'])
            Hi_red = Hi_xyz[::2]
            Ht_red = Ht_xyz[::2]
            file_name = RAW_data_file_location
            frame.to_excel(file_name)
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++Appending the observations++++++++++
            #+++++++++++++++++++++++++++++++++++++++++++++++++
            #Import file location
            Import_path = Export_path + '/Exports' 
            Import_path = Import_path.replace("\\", "/")          
            Import_COORD_DATA_file = '/COORD DATA.xlsx'
            Import_RAW_DATA_file = '/RAW DATA.xlsx'
            #++++++++++++++++++++++determining where in import and export the files too++++++++++++++++++++++++
            #Import_file
            Import_COORD_DATA = Import_path + Import_COORD_DATA_file
            Import_RAW_DATA = Import_path + Import_RAW_DATA_file
            #NEED TO CREAT A NEW TABLE FOR COORDINATES
            Coordinates_table = pd.read_excel (Import_COORD_DATA, skiprows = 1) #place "r" before the path string to address special character, such as '\'. Don't forget to put the file name at the end of the path + '.xlsx'
            Observations_table = pd.read_excel (Import_RAW_DATA, skiprows = 1)
            #+++++++++++++++++++Converting data++++++++++++++++++++++++++++++++++++++++++++++
            #Convert the horizontal bearings from decimal degrees to DMS
            Hz_bearing_list = []
            for i in HzBearing_list_inport: 
                Hz_bearing_list.append(dms_to_dd(i))
            #Convert the vertical bearings from decimal degrees to DMS
            Vz_bearing_list = []
            for i in VzBearing_list_inport: 
                Vz_bearing_list.append(dms_to_dd(i))
            #+++++++++++++++++++Converting data++++++++++++++++++++++++++++++++++++++++++++++
            #+++++++++++++++++++Averaging all the data+++++++++++++++++++++++++++++++++++++++
            first_value = 0
            second_value = 1
            FLred_Hz = []
            for i in Hz_bearing_list:
                if second_value <= len(Hz_bearing_list):
                    FL = Hz_bearing_list[first_value]
                    FR = Hz_bearing_list[second_value]
                    temp = Horizontal_FLred(FL, FR)
                    FLred_Hz.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            FLred_Vz = []
            for i in Vz_bearing_list:
                if second_value <= len(Vz_bearing_list):
                    FL = Vz_bearing_list[first_value]
                    FR = Vz_bearing_list[second_value]
                    temp = Vertical_FLred(FL, FR)
                    FLred_Vz.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            Hz_dist = []
            for i in HzDist_list_inport:
                if second_value <= len(HzDist_list_inport):
                    dist1 = HzDist_list_inport[first_value]
                    dist2 = HzDist_list_inport[second_value]
                    dist1 = float(dist1)
                    dist2 = float(dist2)
                    temp = (dist1 + dist2)/2
                    Hz_dist.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            first_value = 0
            second_value = 1
            Sl_dist = []
            for i in SlDist_list_inport:
                if second_value <= len(SlDist_list_inport):
                    dist1 = SlDist_list_inport[first_value]
                    dist2 = SlDist_list_inport[second_value]
                    dist1 = float(dist1)
                    dist2 = float(dist2)
                    temp = (dist1 + dist2)/2
                    Sl_dist.append(temp)
                    first_value +=2
                    second_value +=2
                else:
                    break
            #+++++++++++++++++++Reducing the data to a new table+++++++++++++++++++++++++++++++++++++++
            new_Stn = Point_at_Stn_list_inport[::2]
            Stn = new_Stn[::2]
            new_Set_list_inport = Set_list_inport[::2]
            new_Face_list = Face_list_inport[::2]
            for i in new_Face_list:
                i = new_Face_list.index('FL')
                # replace Rahul with Shikhar
                new_Face_list = new_Face_list[:i]+['FLred']+new_Face_list[i+1:]
            new_Point_ID_list = Point_ID_list_inport[::2]
            Point_ID = []
            for i in new_Point_ID_list:
                if i not in Point_ID:
                    Point_ID.append(i)
            new_String_list = String_list_inport[::2]
            #+++++++++++++++++++Reducing the data to a new table+++++++++++++++++++++++++++++++++++++++
            reduced_table1 = pd.DataFrame(list(zip(new_Stn, new_Set_list_inport,  new_Face_list, new_Point_ID_list, new_String_list, FLred_Hz, FLred_Vz, Hz_dist, Sl_dist, Hi_red, Ht_red)),
                            columns =['Station', 'Set', 'Face', 'Point ID', 'String', 'Horizontal Bearing (DMS)', 'Vertical Bearing (DMS)', 'Horizontal distance (m)', 'Slope distance (m)', 'Height of Instrument (m)', 'Height of Target (m)'])
            points = []

            for i in Point_ID_list_inport:
                if i not in points:
                    points.append(i)
            counter = new_Point_ID_list.count(new_Point_ID_list[0])
            position_list = []
            list_of_positions = []
            temp_list = new_Point_ID_list
            for i in points:
                point = i
                position = temp_list.index(point)
                for k in temp_list:
                    if point in temp_list:
                        index = temp_list.index(point)
                        position_list.append(index)
                        temp_list[index] = ''
                        list_of_positions.append(position_list)
                    else:
                        break
            list_of_positions = list_of_positions[0]
            def divide_chunks(l, n):
                 
                # looping till length l
                for i in range(0, len(l), n):
                    yield l[i:i + n]
            #Definition sourced from https://www.geeksforgeeks.org/break-list-chunks-size-n-python/
            list_of_positions = list(divide_chunks(list_of_positions, counter))
            #++++++++++++++++++FINAL REDUCTION FOR THE VALUES++++++++++++++++
            def final_reduction(a): 
                reduced_values = []
                locations = []
                average_point = []
                bounds = len(list_of_positions[0])
                for i in list_of_positions:
                    locations = i
                    for i in locations: 
                        average_point.append(a[i])
                average_values = list(divide_chunks(average_point, bounds))     
                for i in average_values:
                    point_average = sum(i)/len(i)
                    reduced_values.append(point_average)
                return reduced_values
            Hz_reduced = list(final_reduction(FLred_Hz))
            Vz_reduced = list(final_reduction(FLred_Vz))
            Hz_dist_reduced = list(final_reduction(Hz_dist))
            Sl_dist_reduced = list(final_reduction(Sl_dist))
            Stn = new_Stn[::2]
            Face_list = new_Face_list[::2]
            String_list = new_String_list[::2]
            Hi_red2 = Hi_red[::2]
            Ht_red2 = Ht_red[::2]
            reduced_table2 = pd.DataFrame(list(zip(Stn, Face_list, Point_ID, String_list, Hz_reduced, Vz_reduced, Hz_dist_reduced, Sl_dist_reduced, Hi_red2, Ht_red2)),
                            columns =['Station', 'Face', 'Point ID', 'String', 'Horizontal Bearing (DMS)', 'Vertical Bearing (DMS)', 'Horizontal distance (m)', 'Slope distance (m)', 'Height of Instrument (m)', 'Height of Target (m)'])
            station_count = set(Point_at_Stn_list_inport)
            station_count = set(station_count)
            how_many_stations = len(station_count)
            set_count = set(new_Set_list_inport)
            how_many_sets = len(set_count)
            if how_many_stations >= 1 and how_many_sets > 1:
                reduced_table = reduced_table2
            else:
                reduced_table = reduced_table1
            file_name = Export_path + '/Exports/Reduced table.xlsx'
            reduced_table.to_excel(file_name)
            Red_Observations_table = pd.read_excel (file_name)
            Stn = Red_Observations_table['Station'].tolist()
            Face = Red_Observations_table['Face'].tolist()
            Point_ID = Red_Observations_table['Point ID'].tolist()
            String = Red_Observations_table['String'].tolist()
            Hz_Bearing_deg = Red_Observations_table['Horizontal Bearing (DMS)'].tolist()
            V_Bearing_deg =  Red_Observations_table['Vertical Bearing (DMS)'].tolist()
            Hz_Distance = Red_Observations_table['Horizontal distance (m)'].tolist()
            Sl_Distance = Red_Observations_table['Slope distance (m)'].tolist()
            Hi = Red_Observations_table['Height of Instrument (m)'].tolist()
            Ht = Red_Observations_table['Height of Target (m)'].tolist()
            Hz_Bearing_rad = []
            for i in Hz_Bearing_deg:
                y = i * (3.141592653589793/180)
                Hz_Bearing_rad.append(y)
            V_Bearing_rad = []
            for i in V_Bearing_deg:
                y = i * (3.141592653589793/180)
                V_Bearing_rad.append(y)
            import math
            sin_part_deltaE = []
            for i in Hz_Bearing_rad:
                x = math.sin(i)
                sin_part_deltaE.append(x)
            deltaE_list = [x*y for x,y in zip(Hz_Distance,sin_part_deltaE)]
            cos_part_deltaN = []
            for i in Hz_Bearing_rad:
                x = math.cos(i)
                cos_part_deltaN.append(x)
            deltaN_list = [x*y for x,y in zip(Hz_Distance,cos_part_deltaN)]
            cos_part_deltaH = []
            for i in V_Bearing_rad:
                x = math.cos(i)
                cos_part_deltaH.append(x)
            deltaH_list = [x*y + (a - b) for x,y,a,b in zip(Sl_Distance,cos_part_deltaH,Hi,Ht)]
            for i in range(len(deltaH_list)):
                print('delta H = ' + str(Sl_Distance[i]) + ' x ' + str(cos_part_deltaH) + ' + ' + str(Hi[i]))
            Easting_fixed = RAWCOORD['Easting (m)'].tolist()
            Northing_fixed = RAWCOORD['Northing (m)'].tolist()
            Height_fixed = RAWCOORD['Height (m)'].tolist()
            Easting_unrounded = [x+y for x,y in zip(Easting_fixed,deltaE_list)]
            Easting_fin = [round(num, 3) for num in Easting_unrounded]
            Northing_unrounded = [x+y for x,y in zip(Northing_fixed,deltaN_list)]
            Northing_fin = [round(num, 3) for num in Northing_unrounded]
            Height_unrounded = [x+y for x,y in zip(Height_fixed,deltaH_list)]
            Height_fin = [round(num, 3) for num in Height_unrounded]
            coords = pd.DataFrame(list(zip(Point_ID, Easting_fin, Northing_fin, Height_fin)),
                            columns =['Point ID', 'Easting (m)', 'Northing (m)', 'Height.ortho (m)'])
            file_name = Export_path + '/Exports/Coordinates.xlsx'
            coords.to_excel(file_name)
            zipPath = directory + '/Exports'
            email_path = directory + "/Exports.zip"
            def get_all_file_paths(directory):
                file_paths = []
                for root, directories, files in os.walk(directory):
                    for filename in files:
                        # join the two strings in order to form the full filepath.
                        filepath = os.path.join(root, filename)
                        file_paths.append(filepath)
                return file_paths        
            def main():
                # path to folder which needs to be zipped
                directory = zipPath
                # calling function to get all file paths in the directory
                file_paths = get_all_file_paths(directory)
                # printing the list of all files to be zipped
                print('Following files will be zipped:')
                for file_name in file_paths:
                    print(file_name)
                # writing files to a zipfile
                with ZipFile(email_path,'w') as zip:
                    # writing each file one by one
                    for file in file_paths:
                        zip.write(file)
                print('All files zipped successfully!')        
            if __name__ == "__main__":
                main()
            email = Email[-1]
            password = Password[-1]
            msg=EmailMessage()
            msg['Subject'] = 'Detail Survey Reductions'
            msg['From']='Xavier Ryan (with love)'
            msg['To']=email
            email_path = directory + "/Exports.zip"
            with open(email_path, "rb") as f:
                file_data=f.read()
                #print("File data in binary", file_data)
                file_name=f.name
                #print("File name is", file_name)
                msg.add_attachment(file_data,maintype="application",subtype="zip",filename=file_name)
            with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
                server.login(email, password)
                server.send_message(msg)
            print('Email Sent!!!')
            Export_folder_location = Export_path + '/' + ExportFolder 
            Coordinates_data_file_location = Export_path + '/' + ExportFolder + '/Coordinates.xlsx'
            Reduced_data_file_location = Export_path + '/' + ExportFolder + '/Reduced table.xlsx'
            Export_zip_data_file_location = Export_path + '/Exports.zip'
            os.unlink(RAW_data_file_location)
            os.unlink(COORD_data_file_location)
            os.unlink(Coordinates_data_file_location)
            os.unlink(Reduced_data_file_location)
            os.unlink(Export_zip_data_file_location)
            os.rmdir(Export_folder_location)
        self.Process = QPushButton("Process Data")
        self.tab6.layout.addRow(self.Process)
        self.tab6.setLayout(self.tab6.layout)
        self.Process.clicked.connect(Process)
        # Add tabs to widget
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
