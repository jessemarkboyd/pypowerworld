#####################################################################
#                   POWERWORLD INTERFACE FILE                       #
# This file defines a class object for interfacing with PW. It is   #
# instantized by a path to a Power World Binary (pwb) file. The     #
# instance methods will be performed on that pwb file. The file may #
# be changed ad hoc.                                                #
#####################################################################

import os.path
import win32com.client
from win32com.client import VARIANT
import pythoncom
import pandas as pd
import numpy as np


def __init__(self,fullfile_path=None):
    self.pw_com = win32com.client.Dispatch('pwrworld.SimulatorAuto') 
    if fullfile_path == None:
        fullfile_path = input('Please enter full pwb file path and name:')
    self.file_path = os.path.split(fullfile_path)[0]
    self.file_name = os.path.splitext(os.path.split(fullfile_path)[1])[0]
    self.auxfile_path = self.file_path + '/' + self.file_name + '.aux'
    self.output = ''
    self.error = False
    self.error_message = ''
    self.open_case()

def __pw_err__(self, pw_output):
    if pw_output is None:
        self.output = None
        self.error = False
        self.error_message = ''
    elif pw_output[0] != '':
        self.output = None
        self.error = False
        self.error_message = ''
    elif 'No data returned' in pw_output[0]:
        self.output = None
        self.error = False
        self.error_message = pw_output[0]
    else:
        self.output = pw_output
        self.error = True
        self.error_message = pw_output[0]
    return self.error            

def open_case(self, file_name=None, file_path=None):
    # Opens case defined by file_name and file_path; if these are undefined, opens by previous file path
    if file_path is not None:
        self.file_path = file_path
    if file_name is not None:
        self.file_name = file_name      
    self.pw_com.OpenCase(self.file_path + '/' + self.file_name + '.pwb')
    if self.__pw_err__():
        print('Error opening case:/n/n%s', self.error_message)
        print('/nPlease check the file name and path and try again (using the open_case method)')

def save_case(self):
    # Saves case with changes to existing file name and path
    self.pw_com.SaveCase(self.file_path + '/' + self.file_name + '.pwb','PWB', 1)
    if self.__pw_err__():
        print('Error saving case:/n/n%s', self.error_message)
        print('/nCASE NOT SAVED!')

def save_case_as(self, file_name=None, file_path=None):
    # If file name and path are specified, saves case as a new file. Overwrites any existing file with the same name and path
    if file_path is not None:
        self.file_path = file_path
    if file_name is not None:
        self.file_name = file_name
    self.auxfile_path = self.file_path + '/' + self.file_name + '.aux'
    self.save_case()
        
def close_case(self):
    # Closes case without saving changes
    self.pw_com.CloseCase()
    if self.__pw_err__():
        print('Error closing case:/n/n%s', self.error_message)

def run_script(self,scriptcommand):
    # Input a script command as in an Auxiliary file SCRIPT{} statement or the PowerWorld Script command prompt
    self.pw_com.RunScriptCommand(scriptcommand)
    if self.__pw_err__():
        print('Error encountered with script:/n/n%s', self.error_message)
        print('Script which was attempted:/n/n%s', scriptcommand)

def load_auxiliary_file_text(self,aux_text):
    # Creates and loads an Auxiliary file under the file_path and 
    # file_name.aux with the text specified in aux_text parameter
    auxfile_obj = open(self.auxfile_path, 'w')
    auxfile_obj.writelines(aux_text)
    auxfile_obj.close()
    self.pw_com.ProcessAuxFile(self.auxfile_path)
    if self.__pw_err__():
        print('Error running auxiliary text:/n/n%s', self.error_message)

def get_parameters_single_element(self, element_type = 'BUS', field_list = ['BusName', 'BusNum'], value_list = [0, 1]):
    # Retrieves parameter data accourding to the fields specified in field_list. 
    # value_list consists of identifying parameter values and zeroes and should be the same length as field_list
    field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
    value_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, value_list) 
    self.pw_com.GetParametersSingleElement(element_type, field_array, value_array)
    if self.__pw_err__():
        print('Error retrieving single element parameters:/n/n%s', self.error_message)
    elif self.error_message != '':
        print(self.error_message)
    else:
        output_df = pd.DataFrame(np.array(self.pwo.output[1]).transpose(),columns=field_list)
        output_df = output_df.replace('',np.nan,regex=True)
        return output_df
    return None

def get_parameters_multiple_element(self, element_type = 'BUS', filter_name = '', field_list = ['BusNum','BusName']):
    field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
    self.pw_com.GetParametersMultipleElement(element_type, field_array, filter_name)
    if self.__pw_err__():
        print('Error retrieving single element parameters:/n/n%s', self.error_message)
    elif self.error_message != '':
        print(self.error_message)
    else:
        output_df = pd.DataFrame(np.array(self.pwo.output[1]).transpose(),columns=field_list)
        output_df = output_df.replace('',np.nan,regex=True)
        return output_df
    return None

def get_3PB_fault_current(self, bus_num):
    # Calculates the three phase fault; this can be done even with cases which 
    # only contain positive sequence impedances
    script_cmd = ('Fault([BUS {}], 3PB);\n'.format(bus_num))
    self.run_script(script_cmd)
    if self.__pw_err__():
        print('Error running 3PB fault:/n/n%s', self.error_message)
    field_list = ['BusNum','FaultCurMag']
    self.get_parameters_single_element('BUS', field_list, [bus_num, 0])
    if self.__pw_err__():
        print('Error retrieving fault current data:/n/n%s', self.error_message)
    else:
        output_df = pd.DataFrame(np.array(self.pwo.output[1]).transpose(),columns=field_list)
        output_df = output_df.replace('',np.nan,regex=True)
        return output_df
    return None
    
def create_filter(self, objecttype='BUS', filtername='', filterlogic='AND', filterpre='NO', enabled='YES', condition):
    # this function creates a filter in PowerWorld. The attempt is to reduce the clunkiness of 
    # creating a filter in the API, which entails creating an aux data file
    auxtext = """
        DATA (FILTER, [ObjectType,FilterName,FilterLogic,FilterPre,Enabled])
        {
        "{}" "{}" "{}" "{}" "YES"
            <SUBDATA Condition>
                {}
            </SUBDATA>
        }""".format(objecttype, filtername, filterlogic, filterpre, enabled, condition)
    self.pwo.load_aux(auxtext)
    if self.pwo.error:
        print('Error creating filter %s:/n/n' % (filtername,self.pwo.error_message))
    return None

def exit(self):
    # Clean up for the PowerWorld COM object
    self.close_case()
    del self.pw_com
    self.pw_com = None

def __del__(self):
    self.exit()