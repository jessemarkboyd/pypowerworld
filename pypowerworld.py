#####################################################################
#                   POWERWORLD INTERFACE FILE                       #
# This file defines a class object for interfacing with PW. It is   #
# instantized by a path to a Power World Binary (pwb) file. The     #
# instance methods will be performed on that pwb file. The file may #
# be changed ad hoc.                                                #
#####################################################################


import pandas as pd
import numpy as np
import os.path
import win32com
from win32com.client import VARIANT
import pythoncom

def __init__(self,fullfilepath=None):
    self.__pwcom__ = win32com.client.Dispatch('pwrworld.SimulatorAuto') 
    if fullfilepath == None:
        fullfilepath = input('Please enter full pwb file path:')
    self.filefolder = os.path.split(fullfilepath)[0]
    self.filename = os.path.splitext(os.path.split(fullfilepath)[1])[0]
    self.auxfilepath = self.filefolder + '/' + self.filename + '.aux'
    self.output = ''
    self.error = False
    self.errormessage = ''
    self.opencase()

def __pwerr__(self):
    if self.__pwcom__.output is None:
        self.output = None
        self.error = False
        self.errormessage = ''
    elif self.__pwcom__.output[0] != '':
        self.output = None
        self.error = False
        self.errormessage = ''
    elif 'No data' in self.__pwcom__.output[0]:
        self.output = None
        self.error = False
        self.errormessage = self.__pwcom__.output[0]
    else:
        self.output = self.__pwcom__.output[1]
        self.error = True
        self.errormessage = self.__pwcom__.output[0]
    return self.error            

def opencase(self, fullfilepath=None):
    # Opens case defined by the full file path; if this are undefined, opens by previous file path
    if fullfilepath is not None:
        self.__init__(fullfilepath)
    else:
        self.__pwcom__.OpenCase(self.filefolder + '/' + self.filename + '.pwb')
        if self.__pwerr__():
            print('Error opening case:\n\n%s\n\n', self.errormessage)
            print('Please check the file name and path and try again (using the opencase method)\n')

def savecase(self):
    # Saves case with changes to existing file name and path
    self.__pwcom__.SaveCase(self.file_path + '/' + self.file_name + '.pwb','PWB', 1)
    if self.__pwerr__():
        print('Error saving case:\n\n%s\n\n', self.errormessage)
        print('******CASE NOT SAVED!******\n\n')

def savecaseas(self, fullfilepath=None):
    # If file name and path are specified, saves case as a new file. Overwrites any existing file with the same name and path
    if fullfilepath is not None:
        self.filefolder = os.path.split(fullfilepath)[0]
        self.filename = os.path.splitext(os.path.split(fullfilepath)[1])[0]
        self.auxfilepath = self.filefolder + '/' + self.filename + '.aux'
    self.savecase()
        
def closecase(self):
    # Closes case without saving changes
    self.__pwcom__.CloseCase()
    if self.__pwerr__():
        print('Error closing case:\n\n%s\n\n', self.errormessage)

def runscriptcommand(self,scriptcommand):
    # Input a script command as in an Auxiliary file SCRIPT{} statement or the PowerWorld Script command prompt
    self.__pwcom__.RunScriptCommand(scriptcommand)
    if self.__pwerr__():
        print('Error encountered with script:\n\n%s\n\n', self.errormessage)
        print('Script command which was attempted:\n\n%s\n\n', scriptcommand)

def loadauxfiletext(self,auxtext):
    # Creates and loads an Auxiliary file with the text specified in auxtext parameter
    f = open(self.auxfilepath, 'w')
    f.writelines(auxtext)
    f.close()
    self.__pwcom__.ProcessAuxFile(self.auxfilepath)
    if self.__pwerr__():
        print('Error running auxiliary text:\n\n%s\n', self.errormessage)

def getparameterssingleelement(self, element_type = 'BUS', field_list = ['BusName', 'BusNum'], value_list = [0, 1]):
    # Retrieves parameter data accourding to the fields specified in field_list. 
    # value_list consists of identifying parameter values and zeroes and should be the same length as field_list
    field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, field_list)
    value_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, value_list) 
    self.__pwcom__.GetParametersSingleElement(element_type, field_array, value_array)
    if self.__pwerr__():
        print('Error retrieving single element parameters:\n\n%s', self.errormessage)
    elif self.errormessage != '':
        print(self.errormessage)
    elif self.__pwcom__.output is not None:
        df = pd.DataFrame(np.array(self.__pwcom__.output[1]).transpose(),columns=field_list)
        df = df.replace('',np.nan,regex=True)
        return df
    return None

def getparametersmultipleelement(self, elementtype, fieldlist, filtername = ''):
    fieldarray = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, fieldlist)
    self.__pwcom__.GetParametersMultipleElement(elementtype, fieldarray, filtername)
    if self.__pwerr__():
        print('Error retrieving single element parameters:\n\n%s\n\n', self.errormessage)
    elif self.errormessage != '':
        print(self.errormessage)
    elif self.__pwcom__.output is not None:
        df = pd.DataFrame(np.array(self.__pwcom__.output[1]).transpose(),columns=fieldlist)
        df = df.replace('',np.nan,regex=True)
        return df
    return None

def get3PBfaultcurrent(self, busnum):
    # Calculates the three phase fault; this can be done even with cases which 
    # only contain positive sequence impedances
    scriptcmd = ('Fault([BUS {}], 3PB);\n'.format(busnum))
    self.run_script(scriptcmd)
    if self.__pwerr__():
        print('Error running 3PB fault:\n\n%s\n\n', self.errormessage)
        return None
    fieldlist = ['BusNum','FaultCurMag']
    return self.getparameterssingleelement('BUS', fieldlist, [busnum, 0])
    
def createfilter(self, condition, objecttype, filtername, filterlogic='AND', filterpre='NO', enabled='YES'):
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
    self.__pwcom__.load_aux(auxtext)
    if self.__pwcom__.error:
        print('Error creating filter %s:\n\n' % (filtername,self.__pwcom__.errormessage))
    return None

def exit(self):
    # Clean up for the PowerWorld COM object
    self.closecase()
    del self.__pwcom__
    self.__pwcom__ = None

def __del__(self):
    self.exit()