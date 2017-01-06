#####################################################################
#                   POWERWORLD INTERFACE FILE                       #
# This file defines a class object for interfacing with PW. It is   #
# instantized by a path to a Power World Binary (pwb) file. The     #
# instance methods will be performed on that pwb file. The file may #
# be changed ad hoc.                                                #
#####################################################################

import sys, os.path
import win32com.client
from win32com.client import VARIANT
import pythoncom
import logging
import time

class PowerWorld(object):

    # Initilized by a connection to the PowerWorld COM object and a PowerWorld file
    def __init__(self,fullfile_path=None):
        
        # Instantize a COM object for the PowerWorld application
        self.pw_com = win32com.client.Dispatch('pwrworld.SimulatorAuto') 
        
        # Check to make sure file path exists
        if fullfile_path == None:
            fullfile_path = input("Please enter full pwb file path and name:")
        self.file_path = os.path.split(fullfile_path)[0]
        self.file_name = os.path.splitext(os.path.split(fullfile_path)[1])[0]
        self.auxfile_path = self.file_path + '/' + self.file_name + '.aux'
        self.output = ''
        self.error = False
	self.error_message = ‘’
	self.open_case()
	if self.error:
		print(“Error encountered opening file %s:/n/t%s” %(self.file_name,self.error_message))
		print(“Please check the file name and path and try again (using the open_case method)”)


    # Open *.pwb case
    def open_case(self):
        temp_status = self.set_output(self.pw_com.OpenCase(self.file_path + '/' + self.file_name + '.pwb'))
        if not temp_status:
            print("Error opening case")
        return temp_status

    # this method should be used to change output, otherwise errors may not be caught
    def set_output(self, temp_output):
        if temp_output[0] != '':
            print(temp_output)
            self.output = None
            return False
        elif "No data returned" in temp_output[0]:
            self.output = None
            return True
        else:
            self.output = temp_output
            return True

    def save_case(self):
        time.sleep(5)
        return self.set_output(self.pw_com.SaveCase(self.file_path + '/' + self.file_name + '.pwb','PWB', 1))

    def close_case(self):
        temp_status = self.set_output(self.pw_com.CloseCase())
        if not temp_status:
            print("Error closing case")
        return temp_status
        
    def change_file_name(self,new_file_name):
        self.file_name = new_file_name 
        self.auxfile_path = self.file_path + '/' + self.file_name + '.aux'
        self.save_case() 

    def run_script(self,scriptCommand):
        return self.set_output(self.pw_com.RunScriptCommand(scriptCommand))

    def load_aux(self,auxText):
        auxfile_obj = open(self.auxfile_path, 'w')
        auxfile_obj.writelines(auxText)
        auxfile_obj.close()
        return self.set_output(self.pw_com.ProcessAuxFile(self.auxfile_path))

    def get_parameters_single_element(self,element_type,field_list,valueList):
        field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
        value_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, valueList) 
        output = self.pw_com.GetParametersSingleElement(element_type, field_array, value_array)
        return self.set_output(output)

    def get_parameters_multiple_element(self,element_type,filter_name,field_list):
        field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
        self.set_output(self.pw_com.GetParametersMultipleElement(element_type, field_array, filter_name))
        return self.output

    def get_parameters_multiple_element_into_dict(self,element_type,field_list,filter_name=' ',field_key_cols = None,optional_key_function = None):
        element_dict = dict()
        if self.get_parameters_multiple_element(element_type,filter_name,field_list):
            for n in range(0,len(self.output[1][0])):
                d = dict()
                for i in range(0,len(self.output[1])):
                    d[field_list[i]] = self.output[1][i][n] 
                if field_key_cols == None:
                    key = n
                elif type(field_key_cols) is int:
                    if optional_key_function == None:
                        value = self.output[1][field_key_cols][n]
                        if value == '':
                            continue
                        else:
                            key = value
                    else:
                        key = optional_key_function(self.output[1][field_key_cols][n])
                elif type(field_key_cols) is tuple:
                    key = ''
                    for x in range(0,len(field_key_cols)):
                        if x > 0:
                            key += ' -- '
                        key += self.output[1][int(field_key_cols[x])][n]
                else:
                    key = n

                # Error handle the case of non-unique keys (keys are used as an ID in dictionaries and must be unique)
                if key in element_dict.keys():
                    print('Attempting to get elements into dict: the key values supplied were not unique to the elements')
                    print('Duplicate key in elements: %s' %str(key))

                else:
                    # The element dictionary object has a key as defined above and the element as defined in the d object
                    element_dict[key] = d

                # Clean up
                del d
                d = None

        # Assume any PW error indicates no objects were returned in the requested data
        else: 
            logging.warning("No elements obtained for dictionary")
        
        return element_dict

    def get_3PB_fault(self,bus_num):
        script_cmd = ('Fault([BUS {}], 3PB);\n'.format(bus_num))
        self.run_script(script_cmd)
        field_list = ['BusNum','FaultCurMag']
        if self.get_parameters_single_element('BUS',field_list,[bus_num,0]):
            return float(self.output[1][1])
        else:
            return None

    def delete_ctgs(self):
        script_cmd = 'DELETE(CONTINGENCY);'
        self.run_script(script_cmd)
        return None
        
    def reopen_case(self):
        self.close_case()
        self.open_case()
        return

    def exit(self):
        self.close_case()
        del self.pw_com
        self.pw_com = None

    def __del__(self):
        self.exit()