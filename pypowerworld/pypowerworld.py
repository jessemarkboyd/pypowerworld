""""#####################################################################
#                   POWERWORLD INTERFACE FILE                       #
# This file defines a class object for interfacing with PW. It is   #
# instantized by a path to a Power World Binary (pwb) file. The     #
# instance methods will be performed on that pwb file. The file may #
# be changed ad hoc by the opencase method.                         #
#####################################################################"""

import pandas as pd
import numpy as np
import os
import win32com
from win32com.client import VARIANT
import pythoncom


class PowerWorld(object):

    # Initilized by a connection to the PowerWorld COM object and a PowerWorld file
    def __init__(self, fullfile_path=None, originalFilePath=None):

        # Instantize a COM object for the PowerWorld application
        self.pw_com = win32com.client.Dispatch('pwrworld.SimulatorAuto')

        # Check to make sure file path exists
        if fullfile_path == None:
            fullfile_path = input("Please enter full pwb file path and name:")

        # The file name and path attributes are used in the methods
        self.file_path = os.path.split(fullfile_path)[0]
        self.file_name = os.path.splitext(os.path.split(fullfile_path)[1])[0]
        self.original_file_path = originalFilePath
        self.original_file_name = os.path.splitext(os.path.split(fullfile_path)[1])[0]

        # Create an aux file path in case an aux file is needed later
        self.auxfile_path = self.file_path + '/' + self.file_name + '.aux'

        # all data returned from PW is contained in the output array
        self.output = ''

        # the error attribute is used to indicate PW error message feedback
        self.error = False

    # Open *.pwb case
    def open_case(self):
        # time.sleep(2)
        temp_status = self.set_output(self.pw_com.OpenCase(self.file_path + '/' + self.file_name + '.pwb'))
        if not temp_status:
            print("Error opening case")
        return temp_status

    # Open *.pwb case
    def open_original_case(self):
        # time.sleep(2)
        temp_status = self.set_output(self.pw_com.OpenCase(self.original_file_path + '/' + self.file_name + '.pwb'))
        if not temp_status:
            print("Error opening case")
        return temp_status

    # this method should be used to change output, otherwise errors may not be caught
    def set_output(self, temp_output):
        if temp_output[0] != '':
            logging.debug(temp_output)
            self.output = None
            return False
        elif "No data returned" in temp_output[0]:
            self.output = None
            return False
        else:
            self.output = temp_output
            return True

    # Save *.pwb case with changes
    def save_case(self):
        # self.output = self.pw_com.SaveCase(self.file_path + '\\' + self.file_name + '.pwb','PWB', 1)
        time.sleep(5)
        return self.set_output(self.pw_com.SaveCase(self.file_path + '\\' + self.file_name + '.pwb', 'PWB', 1))

    # Close PW case, retain the PW interface object
    def close_case(self):
        # time.sleep(5)
        temp_status = self.set_output(self.pw_com.CloseCase())
        if not temp_status:
            print("Error closing case")
        # time.sleep(5)
        return temp_status

    # The file name is change so that the original is not changed during the analysis
    def change_file_name(self, new_file_name):

        # Change the filename in the object attributes
        self.file_name = new_file_name

        # Update the auxiliary file name
        self.auxfile_path = self.file_path + '/' + self.file_name + '.aux'

        # Save case automatically saves as the filename in object attributes
        self.save_case()

        # Scripts can be run through the pw_com object. Scripts are essentially PowerWorld methods

    def run_script(self, scriptCommand):
        # self.output = self.pw_com.RunScriptCommand(scriptCommand)
        return self.set_output(self.pw_com.RunScriptCommand(scriptCommand))

    # Auxiliary files are used to change data in PowerWorld (sometimes data *must* be changed through aux files)
    def load_aux(self, auxText):

        # Save a *.aux text file with the auxText as content
        auxfile_obj = open(self.auxfile_path, 'w')
        auxfile_obj.writelines(auxText)
        auxfile_obj.close()

        # Open the aux file in PW to modify the case
        return self.set_output(self.pw_com.ProcessAuxFile(self.auxfile_path))

    # Get PW data for a single element (e.g. a single bus, branch or zone)
    def get_parameters_single_element(self, element_type, field_list, valueList):

        # The field array has the names of the key fields and the fields of the requested data
        field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)

        # The value list contains values in the key fields and zeros for all other fields
        # The key fields define exactly which element data is requested for
        # The zero values are place-holders for requested data
        value_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, valueList)

        # The data on the single element is stored in output[1] (in the same order as the field_array)
        output = self.pw_com.GetParametersSingleElement(element_type, field_array, value_array)

        return self.set_output(output)

    # Get PW data for multiple elements (e.g. multiple buses, branches or zones)
    def get_parameters_multiple_element(self, element_type, filter_name, field_list):

        # The field array has the names of the fields of the requested data
        field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)

        # The data on the single element is stored in output[1]
        # The element_type is branch or bus or zone, etc.
        # The filter can be defined as '' if no filter is desired
        return self.set_output(self.pw_com.GetParametersMultipleElement(element_type, field_array, filter_name))

    # Get multiple elements and return the values as a dictionary object
    def get_parameters_multiple_element_into_dict(self, element_type, filter_name, field_list, field_key_cols=None,
                                                  optional_key_function=None):

        # Create a dictionary to hold all the elements returned. This will actually be a dictionary of dictionaries.
        element_dict = dict()

        # Get the multiple elements into an array
        # if the data transfer was successful, convert the array into a dictionary
        if self.get_parameters_multiple_element(element_type, filter_name, field_list):

            # Loop through each element
            for n in range(0, len(self.output[1][0])):

                # Loop through each column (i.e. attribute) of each element
                d = dict()
                for i in range(0, len(self.output[1])):
                    # each element attribute is defined by the field name in PowerWorld
                    d[field_list[i]] = self.output[1][i][n]

                    # define a default key for the element dictionary
                if field_key_cols == None:
                    key = n

                # if the field_key_cols is defined, the element_dict keys are determined by them
                # The field_key_cols is an integer or tuple of integers of which element col(s) are used to create the dictionary keys
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
                    for x in range(0, len(field_key_cols)):
                        if x > 0:
                            key += ' -- '
                        key += self.output[1][int(field_key_cols[x])][n]
                else:
                    key = n

                # Error handle the case of non-unique keys (keys are used as an ID in dictionaries and must be unique)
                if key in element_dict.keys():
                    logging.error(
                        'Attempting to get elements into dict: the key values supplied were not unique to the elements')
                    logging.error('Duplicate key in elements: %s' % str(key))

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

    def get_3PB_fault(self, bus_num):

        # Calculate 3 phase fault at the bus
        script_cmd = ('Fault([BUS %d], 3PB);\n' % bus_num)
        self.run_script(script_cmd)

        # Retrieve fault analysis results
        field_list = ['BusNum', 'FaultCurMag']
        if self.get_parameters_single_element('BUS', field_list, [bus_num, 0]):
            return float(self.output[1][1])
        else:
            return None

    # PowerWorld has the ability to send data to excel. This method is currently unused
    def send_to_excel(self, element_type, filter_name, field_list):
        field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
        return set_output(self.pw_com.SendToExcel(element_type, filter_name, field_array))

    # Reopen the same case to start fresh
    def reopen_case(self):
        self.close_case()
        self.open_case()
        return

    # The case should be closed upon exit. If not, the PW instance will continue to exist
    def exit(self):
        self.close_case()
        del self.pw_com
        self.pw_com = None

    # Clean up -- the __del__ functions is automatically entered when the class object is deleted
    def __del__(self):
        self.exit()
