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


class PyPowerWorld(object):
    """Class object designed for easy interface with PowerWorld."""
    def __init__(self, pwb_file_path=None):
        try:
            self.__pwcom__ = win32com.client.Dispatch('pwrworld.SimulatorAuto')
        except Exception as e:
            print(str(e))
            print("Unable to launch SimAuto.",
                  "Please confirm that your PowerWorld license includes the SimAuto add-on ",
                  "and that SimAuto has been successfuly installed.")
        self.pwb_file_path = pwb_file_path
        self.__setfilenames__()
        self.output = ''
        self.error = False
        self.error_message = ''
        self.COMout = ''
        self.opencase()

    def __setfilenames__(self):
        self.file_folder = os.path.split(self.pwb_file_path)[0]
        self.file_name = os.path.splitext(os.path.split(self.pwb_file_path)[1])[0]
        self.aux_file_path = self.file_folder + '/' + self.file_name + '.aux'  # some operations require an aux file
        self.save_file_path = os.path.splitext(os.path.split(self.pwb_file_path)[1])[0]

    def __pwerr__(self):
        if self.COMout is None:
            self.output = None
            self.error = False
            self.error_message = ''
        elif self.COMout[0] == '':
            self.output = None
            self.error = False
            self.error_message = ''
        elif 'No data' in self.COMout[0]:
            self.output = None
            self.error = False
            self.error_message = self.COMout[0]
        else:
            self.output = self.COMout[-1]
            self.error = True
            self.error_message = self.COMout[0]
        return self.error            
    
    def opencase(self, pwb_file_path=None):
        """Opens case defined by the full file path; if this is undefined, opens by previous file path"""
        if pwb_file_path is None and self.pwb_file_path is None:
            pwb_file_path = input('Enter full pwb file path > ')
        if pwb_file_path:
            self.pwb_file_path = os.path.splitext(pwb_file_path)[0] + '.pwb'
        else:
            self.COMout = self.__pwcom__.OpenCase(self.file_folder + '/' + self.file_name + '.pwb')
            if self.__pwerr__():
                print('Error opening case:\n\n%s\n\n', self.error_message)
                print('Please check the file name and path and try again (using the opencase method)\n')
                return False
        return True
    
    def savecase(self):
        """Saves case with changes to existing file name and path."""
        self.COMout = self.__pwcom__.SaveCase(self.pwb_file_path, 'PWB', 1)
        if self.__pwerr__():
            print('Error saving case:\n\n%s\n\n', self.error_message)
            print('******CASE NOT SAVED!******\n\n')
            return False
        return True
    
    def savecaseas(self, pwb_file_path=None):
        """If file name and path are specified, saves case as a new file.
        Overwrites any existing file with the same name and path."""
        if pwb_file_path is not None:
            self.pwb_file_path = os.path.splitext(pwb_file_path)[1] + '.pwb'
            self.__setfilenames__()
        return self.savecase()
    
    def savecaseasaux(self, file_name=None, FilterName='', ObjectType=None, ToAppend=True, FieldList='all'):
        """If file name and path are specified, saves case as a new aux file.
        Overwrites any existing file with the same name and path."""
        if file_name is None:
            file_name = self.file_folder + '/' + self.file_name + '.aux'
        self.file_folder = os.path.split(file_name)[0]
        self.save_file_path = os.path.splitext(os.path.split(file_name)[1])[0]
        self.aux_file_path = self.file_folder + '/' + self.save_file_path + '.aux'
        self.COMout = self.__pwcom__.WriteAuxFile(self.aux_file_path,FilterName,ObjectType,ToAppend,FieldList)
        if self.__pwerr__():
            print('Error saving case:\n\n%s\n\n', self.error_message)
            print('******CASE NOT SAVED!******\n\n')
            return False
        return True
            
    def closecase(self):
        """Closes case without saving changes."""
        self.COMout = self.__pwcom__.CloseCase()
        if self.__pwerr__():
            print('Error closing case:\n\n%s\n\n', self.error_message)
            return False
        return True
    
    def runscriptcommand(self,script_command):
        """Input a script command as in an Auxiliary file SCRIPT{} statement or the PowerWorld Script command prompt."""
        self.COMout = self.__pwcom__.RunScriptCommand(script_command)
        if self.__pwerr__():
            print('Error encountered with script:\n\n%s\n\n', self.error_message)
            print('Script command which was attempted:\n\n%s\n\n', script_command)
            return False
        return True
    
    def loadauxfiletext(self,auxtext):
        """Creates and loads an Auxiliary file with the text specified in auxtext parameter."""
        f = open(self.aux_file_path, 'w')
        f.writelines(auxtext)
        f.close()
        self.COMout = self.__pwcom__.ProcessAuxFile(self.aux_file_path)
        if self.__pwerr__():
            print('Error running auxiliary text:\n\n%s\n', self.error_message)
            return False
        return True
    
    def getparameterssingleelement(self, element_type = 'BUS', field_list = ['BusName', 'BusNum'], value_list = [0, 1]):
        """Retrieves parameter data according to the fields specified in field_list.
        value_list consists of identifying parameter values and zeroes and should be
        the same length as field_list"""
        assert len(field_list) == len(value_list)
        field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
        value_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, value_list) 
        self.COMout = self.__pwcom__.GetParametersSingleElement(element_type, field_array, value_array)
        if self.__pwerr__():
            print('Error retrieving single element parameters:\n\n%s', self.error_message)
        elif self.error_message != '':
            print(self.error_message)
        elif self.__pwcom__.output is not None:
            df = pd.DataFrame(np.array(self.__pwcom__.output[1]).transpose(),columns=field_list)
            df = df.replace('',np.nan,regex=True)
            return df
        return None

    def getparametersmultipleelement(self, elementtype, fieldlist, filtername = ''):
        fieldarray = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, fieldlist)
        self.COMout = self.__pwcom__.GetParametersMultipleElement(elementtype, fieldarray, filtername)
        if self.__pwerr__():
            print('Error retrieving single element parameters:\n\n%s\n\n', self.error_message)
        elif self.error_message != '':
            print(self.error_message)
        elif self.__pwcom__.output is not None:
            df = pd.DataFrame(np.array(self.__pwcom__.output[1]).transpose(), columns=fieldlist)
            df = df.replace('', np.nan, regex=True)
            return df
        return None

    def get3PBfaultcurrent(self, busnum):
        """Calculates the three phase fault; this can be done even with cases which
        only contain positive sequence impedances"""
        scriptcmd = f'Fault([BUS {busnum}], 3PB);\n'
        self.COMout = self.run_script(scriptcmd)
        if self.__pwerr__():
            print('Error running 3PB fault:\n\n%s\n\n', self.error_message)
            return None
        fieldlist = ['BusNum', 'FaultCurMag']
        return self.getparameterssingleelement('BUS', fieldlist, [busnum, 0])
        
    def createfilter(self, condition, objecttype, filtername, filterlogic='AND', filterpre='NO', enabled='YES'):
        """Creates a filter in PowerWorld. The attempt is to reduce the clunkiness of
        # creating a filter in the API, which entails creating an aux data file"""
        auxtext = '''
            DATA (FILTER, [ObjectType,FilterName,FilterLogic,FilterPre,Enabled])
            {
            "{objecttype}" "{filtername}" "{filterlogic}" "{filterpre}" "{enabled]"
                <SUBDATA Condition>
                    {condition}
                </SUBDATA>
            }'''.format(condition=condition, objecttype=objecttype, filtername=filtername, filterlogic=filterlogic,
                        filterpre=filterpre, enabled=enabled)
        self.COMout = self.__pwcom__.load_aux(auxtext)
        if self.__pwcom__.error:
            print('Error creating filter %s:\n\n%s' % (filtername,self.__pwcom__.error_message))
            return False
        return True
    
    def exit(self):
        """Clean up for the PowerWorld COM object"""
        self.closecase()
        del self.__pwcom__
        self.__pwcom__ = None
        return None
    
    def __del__(self):
        self.exit()