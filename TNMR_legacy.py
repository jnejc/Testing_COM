# Communication with windows API devices
import comtypes
from comtypes.client import CreateObject
from comtypes.client import GetActiveObject

#import os # For compiling file paths
import time

from T1 import Geometric_list
from T1 import Arithmetic_list
from T1 import ToN
from T1 import ToStr
from NMR7 import Legacy_export
from NMR7 import Legacy_G

# Global variables
MEAS_LOOP_TIME = 2 # time in seconds how often python interacts with tnmr


class Tecmag():
    '''Creates application class for TNMR communication'''

    def __init__(self):
        self.Open_TNMR()

    
    def Open_TNMR(self):
        '''Finds TNMR window or creates new one'''
        try:
            self.app = GetActiveObject('NTNMR.Application')
            print('Found existing TNMR window')
        except OSError:
            self.app = CreateObject('NTNMR.Application')
            print('Opening new TNMR window')
            # Close the empty document
            if self.app.CloseFile(''): pass
            else: print('Failed to close starting file')


    def Get_FID(self, params):
        ''' Measure a single FID'''
        # File path for pulse programs
        pulse_path = 'C:\\TNMR\\sequences\\'
        pulse_file = 'two_pulse.tps'

        # Adds paths to params
        params['pulse_path'] = pulse_path
        params['pulse_file'] = pulse_file

        file_path = params['file_path']
        file_name = params['file_name']

        # Define dictionary of parameters for this two pulse experiment
        nmrparam = {
            # Pulse parameters
            'a1': 'a123',
            'a2': 'a123',
            'atn1': 'atn1',
            'atn2': 'atn23',
            'd1': 'd123',
            'd2': 'd123',
            'ad': 'ad',

            'tau': 'TAU',

            # Experiment parameters
            'Last Delay': 'D9',
            'Scans 1D': 'NS',
            'Observe Freq.': 'FR'
        }

        # Default parameters
        default_params = {
            'Dwell Time': 'DW',
            'Acq. Points': 'TD',
            'Filter': 'Filter',
            'Receiver Gain': 'RecGain',
            'Receiver Phase': 'RecPh',
            # Pulse safety
            'pre': 'pre',
            'rd': 'rd',
            'Nucleus': 'Nuc',
            # Visual
            'Data Type': 'Data Type',
            'Vertical Scale': 'Vertical Scale',
            'Zoom End': 'TD'
        }
        # Add to main params
        nmrparam.update(default_params)

        #return data
        return self.Run_measurement(nmrparam, params)


    def Get_INV(self, params):
        ''' Measure a single FID'''
        # File path for pulse programs
        pulse_path = 'C:\\TNMR\\sequences\\'
        pulse_file = 'three_pulse.tps'

        # Adds paths to params
        params['pulse_path'] = pulse_path
        params['pulse_file'] = pulse_file

        file_path = params['file_path']
        file_name = params['file_name']

        # Define dictionary of parameters for this two pulse experiment
        nmrparam = {
            # Pulse parameters
            'a1': 'a123',
            'a2': 'a123',
            'a3': 'a123',
            'atn1': 'atn1',
            'atn2': 'atn23',
            'atn3': 'atn23',
            'd1': 'd123',
            'd2': 'd123',
            'd3': 'd123',
            'ad': 'ad',

            'tau': 'TAU',
            'd5': 'D5meas',

            # Experiment parameters
            'Last Delay': 'D9',
            'Scans 1D': 'NS',
            'Observe Freq.': 'FR'
        }

        # Default parameters
        default_params = {
            'Dwell Time': 'DW',
            'Acq. Points': 'TD',
            'Filter': 'Filter',
            'Receiver Gain': 'RecGain',
            'Receiver Phase': 'RecPh',
            # Pulse safety
            'pre': 'pre',
            'rd': 'rd',
            'Nucleus': 'Nuc',
            # Visual
            'Data Type': 'Data Type',
            'Vertical Scale': 'Vertical Scale',
            'Zoom End': 'TD'
        }
        # Add to main params
        nmrparam.update(default_params)

        #return data
        return self.Run_measurement(nmrparam, params)


    def Get_T1(self, params):
        '''Takes Params and makes T1 measurement'''
        # Generate D5 list
        d5_list = Geometric_list(params['D5min'],
            params['D5max'], params['D5N'], shuffle=True)

        # Remember file name
        file_name = params['file_name']

        # Make a fake .G file
        Legacy_G(d5_list, params)

        # Do all measurements
        for i, d5 in enumerate(d5_list.split(' ')):
            # Assign D5 and file name
            params['D5'] = d5
            # Reduce to accomodate for prepulse
            params['D5meas'] = ToStr(ToN(d5) - ToN(params['pre']))
            params['file_name'] = file_name + '-' + (str(1000+i+1)[-3:])

            # Measure FID (three pulse inversion) and legacy export
            data = self.Get_INV(params)
            Legacy_export(data, params)


    def Get_T2(self, params):
        '''Takes Params and makes T1 measurement'''
        # Generate D5 list
        tau_list = Arithmetic_list(params['dTAU'], params['TAUN'],
            start=params['TAUmin'], shuffle=True)

        # Remember file name
        file_name = params['file_name']

        # Make a fake .G file
        Legacy_G(tau_list, params)

        # Do all measurements
        for i, tau in enumerate(tau_list.split(' ')):
            # Assign D5 and file name
            params['TAU'] = tau
            params['ad'] = ToStr(ToN(tau) - ToN(params['TAUmin'])
                 + ToN(params['adT2']))
            params['file_name'] = file_name + '-' + (str(1000+i+1)[-3:])

            # Measure FID (three pulse inversion) and legacy export
            data = self.Get_FID(params)
            Legacy_export(data, params)


    def Run_measurement(self, nmrparam, params):
        '''Runs measurement on doc, waits till finished 
            Possibly return data once done'''
        # with makes sure that communication aborts correctly
        with TNMR_document(params['file_path'], params['file_name'], self.app) as doc:
            # Loads file program
            doc.LoadSequence(params['pulse_path'] + params['pulse_file'])

            # Set the parameters in file
            for key, value in nmrparam.items():
                doc.SetNMRParameter(key, params[value])

            # Measure
            print('Starting measurement')
            # Zero and go
            doc.ZG()

            # Checks if the measurement is finished
            while not self.app.CheckAcquisition: # End loop when finished
                print('Measurement in progress')
                time.sleep(MEAS_LOOP_TIME)
        
            print('Measurement finished')
            
            # Read parameters from file
            self.Get_parameters(doc, params)

            # Return Re Im data
            return doc.GetData 

        # Exit on with saves file and exports          


    def Get_parameters(self, doc, params):
        '''Gets the desired parameters from TNMR doc
            and saves them into params'''
        # Get start finish time in HH:MM:SS
        start = doc.GetNMRParameter('Exp. Start Time')[-8:]
        finish = doc.GetNMRParameter('Exp. Finish Time')[-8:]

        # Get date and reformat into DD.MM.YYYY format
        date = doc.GetNMRParameter('Date')
        date = '.'.join(reversed(date.split(' ')[0].split('/')))

        # Get actual NS
        ns = doc.GetNMRParameter('Actual Scans 1D')

        # Write to params
        params['DATESTA'] = date
        params['TIMESTA'] = start
        params['TIMEEND'] = finish
        params['NS'] = ns


    def Test_params(self):
        '''Creates dictionary of params for BCAO NMR'''
        params = dict()
        # Use string values for communication with Tecmag

        # File and directory parameters
        params['file_path'] = 'C:\\TNMR\\data\\Nejc\\20210217_As_BCAO\\'
        params['file_key'] = 'test-2p1T-10K'
        #params['file_name'] = 'T2-As-x1-1p3T-18K'

        # Default parameters
        params['DW'] = '100n'       # Dwell time
        params['TD'] = '2048'       # Time domain / acq. points
        params['RecGain'] = '25'     # Reciever gain
        params['RecPh'] = '155'     # Reciever phase
        params['pre'] = '3u'        # Prepare for pulse
        params['rd'] = '3u'         # Ringdown
        params['Filter'] = '100000' # Signal filter?

        # Pulse definition
        params['a123'] = '28'       # Pulse amplitude
        params['atn1'] = '19'       # Pi/2 pulse attenuation
        params['atn23'] = '13'      # Pi pulse attenuation
        params['d123'] = '3u'       # Pulse duration

        # Sample property
        params['FR'] = '24.88'      # Central frequency
        params['NS'] = '256'        # Number of scans
        params['D9'] = '1s'        # Last Delay
        params['TAU'] = '35u'       # Default tau
        params['ad'] = '30u'         # Acquisition delay
        params['Nuc'] = 'Ba'    # Nucleus

        # T1
        params['D5N'] = '20'        # Number of points in T1
        params['D5min'] = '1u'
        params['D5max'] = '1s'
        params['adT1'] = '15u'       # Acquisiton delay
        params['D5'] = '4u'         # D5 var value, adds to list :S

        # T2
        params['TAUN'] = '60'        # Number of points in T2
        params['TAUmin'] = '35u'
        params['dTAU'] = '50u'
        params['adT2'] = '15u'        # Acquisiton delay

        # Environment               # Set these by hand for now
        params['Tset'] = '10.0'    # Set temperature
        params['Tsens'] = '10.0'   # Cryostat temperature
        params['Tprobe'] = '10.05'  # Probe temperature
        params['Fpers'] = '2.1'     # Persistent field
        params['Pcryo'] = '660m'    # Pressure on manometer

        # Visual
        params['Vertical Scale'] = '10000000'
        params['Data Type'] = 'Real, Imaginary, Magnitude'

        self.params = params



class TNMR_document():
    '''Class that opens a TNMR document, returns reference to it
    and makes sure to save and close when finished using with'''

    def __init__(self, file_path, file_name, app):
        '''Initialization when class is called, reference to TNMR app'''
        self.file_path = file_path
        self.file_name = file_name
        self.app = app


    def __enter__(self):
        '''Entering functions when class is started up with with'''
        print('Creating:', self.file_path, self.file_name)
        self.doc = CreateObject('NTNMR.Document')
        return self.doc


    def __exit__(self, e_type, e_value, e_traceback):
        '''Closing up functions when code in with is finished'''
        if e_value != None:
            print('Found error: ', e_type, e_value)
            if e_type == KeyboardInterrupt:
                print('Interrupting measurement', self.file_name)
                if not self.app.CheckAcquisition:
                    print('Aborting')
                    self.app.Abort
        else:
            print('Closing measurement', self.file_name)

        # Save document and close it
        self.doc.SaveAs(self.file_path + self.file_name + '.tnt')
        #self.doc.Export(self.file_path + self.file_name + '.txt', 0)

        if not self.app.CloseActiveFile: print('Failed to close file')
    
        #return True # Supresses errors in code


    

if __name__ == "__main__":
    A = Tecmag()
    print(A.app.GetDocumentList)

    # Load parameters
    A.Test_params()

    # Do FID, T1 and T2
    #A.params['file_name'] = 'FID-' + A.params['file_key']
    #data = A.Get_FID(A.params)
    #Legacy_export(data, A.params)

    A.params['file_name'] = 'T1-' + A.params['file_key']
    A.Get_T1(A.params)

    #A.params['file_name'] = 'T2-' + A.params['file_key']
    #A.Get_T2(A.params)


    # Automate
"""
    key = A.params['file_name'].split('-')[0]
    if key == 'T1':
        A.Get_T1(A.params)
    elif key == 'T2':
        A.Get_T2(A.params)
    elif key == 'FID':
        data = A.Get_FID(A.params)
        Legacy_export(data, A.params)
    elif key == 'INV':
        data = A.Get_INV(A.params)
        Legacy_export(data, A.params)
    else:
        print('Cannot recognize measurement type:', key)
"""


