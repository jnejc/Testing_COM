# Communication with windows API devices
import comtypes
from comtypes.client import CreateObject
from comtypes.client import GetActiveObject

#import os # For compiling file paths
import time

from T1 import Geometric_list
from T1 import Arithmetic_list



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

    # These experiments should correspond to specific pulse files
    def T1(self, params):
        '''Does a new T1 measurement in a new file and saves results'''
        # File path for pulse programs
        pulse_path = 'C:\\TNMR\\sequences\\'
        pulse_file = 'T1.tps'

        # Save path for measurements
        file_path = 'C:\\TNMR\\data\\Nejc\\210202_As_BCAO_NQR\\'
        file_name = 'T1_As_0T_200K'

        # Define dictionary of parameters for this experiment
        self.nmrparam_dict = {
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
            'ad': 'adT1',

            'tau': 'TAU',
            'd5': 'D5off',

            # Experiment parameters
            'Last Delay': 'D9',
            'Points 2D': 'D5N',
            'Scans 1D': 'NS',
            'Observe Freq.': 'FR'
        }

        # Generate D5 table
        d5_list = Geometric_list(params['D5min'],
            params['D5max'], params['D5N'], shuffle=False)

        # with makes sure that communication correctly aborts in the end
        with TNMR_document(file_path, file_name, self.app) as doc:
            doc.LoadSequence(pulse_path + pulse_file)

            # Set regular parameters
            self.Default_params(doc)

            # Set the parameters
            for param in self.nmrparam_dict:
                doc.SetNMRParameter(param,
                    params[self.nmrparam_dict[param]])

            # Set d5 table
            doc.SetTable('d5:2', d5_list)

            # Display all three lines on graph
            doc.SetNMRParameter('Data Type',
                'Real, Imaginary, Magnitude')

            # Start measurement and write data
            self.data_T1 = self.Run_measurement(doc)

            # Save file and export data
            doc.SaveAs(file_path + file_name + '.tnt')
            doc.Export(file_path + file_name + '.txt', 0)


    def T2(self, params):
        '''Does a new T2 measurement in a new file and saves results'''
        # File path for pulse programs
        pulse_path = 'C:\\TNMR\\sequences\\'
        pulse_file = 'T2.tps'

        # Save path for measurements
        file_path = 'C:\\TNMR\\data\\Nejc\\210202_As_BCAO_NQR\\'
        file_name = 'T2_As_0T_200K'

        # Define dictionary of parameters for this experiment
        self.nmrparam_dict = {
            # Pulse parameters
            'a1': 'a123',
            'a2': 'a123',
            'atn1': 'atn1',
            'atn2': 'atn23',
            'd1': 'd123',
            'd2': 'd123',
            'ad': 'adT2',

            'tau': 'TAUmin',

            # Experiment parameters
            'Last Delay': 'D9',
            'Points 2D': 'TAUN',
            'Scans 1D': 'NS',
            'Observe Freq.': 'FR',

            # Tau increment
            'dtau': 'dTAU'
        }

        # Generate D5 table
        #dtau_list = Arithmetic_list(params['dTAU'], params['TAUN'])
        #print(dtau_list)

        # with makes sure that communication correctly aborts in the end
        with TNMR_document(file_path, file_name, self.app) as doc:
            doc.LoadSequence(pulse_path + pulse_file)

            # Set regular parameters
            self.Default_params(doc)

            # Set the parameters
            for param in self.nmrparam_dict:
                print(param)
                doc.SetNMRParameter(param,
                    params[self.nmrparam_dict[param]])

            # Set tau list (automatic increment)
            #doc.SetTable('dtau:2', dtau_list)
            #doc.SetNMRParameter('deltatau', params['dTAU'])

            # Start measurement and write data
            self.data_T2 = self.Run_measurement(doc)

            # Save file and export data
            doc.SaveAs(file_path + file_name + '.tnt')
            doc.Export(file_path + file_name + '.txt', 0)


    def Get_FID(self, params):
        ''' Measure a single FID'''
# File path for pulse programs
        pulse_path = 'C:\\TNMR\\data\\Nejc test\\Al Ref\\'
        pulse_file = 'two-pulse_27Al.tps'

        # Save path for measurements
        file_path = 'C:\\TNMR\\data\\Nejc test\\Al Ref\\'
        file_name = 'test_FID'

        # Define dictionary of parameters for this experiment
        self.nmrparam_dict = {
            # Pulse parameters
            'a1': 'a123',
            'atn1': 'atn1',
            'atn2': 'atn23',
            'd1': 'd123',
            'ad': 'ad',

            'tau': 'TAU',

            # Experiment parameters
            'Last Delay': 'D9',
            'Scans 1D': 'NS',
            'Observe Freq.': 'FRal'
        }

        # with makes sure that communication correctly aborts in the end
        with TNMR_document(file_path, file_name, self.app) as doc:
            doc.LoadSequence(pulse_path + pulse_file)

            # Set regular parameters
            self.Default_params(doc)

            # Set the parameters
            for param in self.nmrparam_dict:
                doc.SetNMRParameter(param,
                    params[self.nmrparam_dict[param]])

            # Display all three lines on graph
            doc.SetNMRParameter('Data Type',
                'Real, Imaginary, Magnitude')

            # Start measurement and write data
            self.data_FID = self.Run_measurement(doc)

            # Save file and export data
            doc.SaveAs(file_path + file_name + '.tnt')
            doc.Export(file_path + file_name + '.txt', 0)

    
    def Run_measurement(self, doc):
        '''Runs measurement on doc, waits till finished 
            Possibly return data once done'''
        print('Starting measurement')
        # Zero and go
        doc.ZG()

        # Cehcks if the measurement is finished
        while not self.app.CheckAcquisition: # End loop when finished
            print('Measurement in progress')
            time.sleep(10)
        
        print('Measurement finished')

        return doc.GetData


    def Default_params(self, doc):
        '''Writes the common measurement parameters to doc'''
        # Parameters common to most experiments
        doc.SetNMRParameter('Dwell Time', '200n')
        doc.SetNMRParameter('Acq. Points', '2048')
        doc.SetNMRParameter('Receiver Gain', '100')  # Might vary
        doc.SetNMRParameter('Receiver Phase', '122')# Unnecessary?
        # Pulse safety
        doc.SetNMRParameter('pre', '3u')    # Prepare for pulse
        doc.SetNMRParameter('rd', '3u')     # Ringdown before acq.


    def Test_params2(self):
        '''Creates dictionary of testing params for Cu NQR'''
        params = dict()
        # Use string values for communication with Tecmag

        # Pulse definition
        params['a123'] = '55'       # Pulse amplitude
        params['atn1'] = '13'       # Pi/2 pulse attenuation
        params['atn23'] = '7'       # Pi pulse attenuation
        params['d123'] = '3u'       # Pulse duration

        # Sample property
        params['FR'] = '25.98MHz'   # Central frequency
        params['NS'] = '128'         # Number of scans
        params['D9'] = '400m'       # Last Delay
        params['TAU'] = '35u'       # Default tau

        # T1
        params['D5N'] = '20'        # Number of points in T1
        params['D5min'] = '20u'
        params['D5max'] = '10s'
        params['adT1'] = '20u'        # Acquisiton delay
        params['D5off'] = '1u'      # D5 var value, adds to list :S

        # T2
        params['TAUN'] = '2'        # Number of points in T2
        params['TAUmin'] = '10u'
        params['dTAU'] = '10u'
        params['adT2'] = '10u'        # Acquisiton delay

        # FID
        params['ad'] = '10u'        # Acquisition delay
        params['FRal'] = '77.693MHz'

        return params

    def Test_params(self):
        '''Creates dictionary of params for BCAO NQR'''
        params = dict()
        # Use string values for communication with Tecmag

        # Pulse definition
        params['a123'] = '45'       # Pulse amplitude
        params['atn1'] = '19'       # Pi/2 pulse attenuation
        params['atn23'] = '13'       # Pi pulse attenuation
        params['d123'] = '4u'       # Pulse duration

        # Sample property
        params['FR'] = '27.7MHz'   # Central frequency
        params['NS'] = '1024'         # Number of scans
        params['D9'] = '30m'       # Last Delay
        params['TAU'] = '10u'       # Default tau

        # T1
        params['D5N'] = '20'        # Number of points in T1
        params['D5min'] = '5u'
        params['D5max'] = '10m'
        params['adT1'] = '5u'        # Acquisiton delay
        params['D5off'] = '1u'      # D5 var value, adds to list :S

        # T2
        params['TAUN'] = '20'        # Number of points in T2
        params['TAUmin'] = '10u'
        params['dTAU'] = '5u'
        params['adT2'] = '5u'        # Acquisiton delay

        # FID
        params['ad'] = '5u'        # Acquisition delay
        params['FRal'] = '27.7MHz'

        return params

    

class TNMR_document():
    '''Class that opens a TNMR document, returns reference to it
    and makes sure to save and close when finished using with'''

    def __init__(self, file_path, file_name, app):
        '''Initialization when class is called, reference to TNMR app'''
        print('init!!')
        self.file_path = file_path
        self.file_name = file_name
        self.app = app
        print(file_path, file_name, app)


    def __enter__(self):
        '''Entering functions when class is started up with with'''
        print('entering')
        self.doc = CreateObject('NTNMR.Document')
        return self.doc


    def __exit__(self, e_type, e_value, e_traceback):
        '''Closing up functions when code in with is finished'''
        if e_value != None:
            print('Found error: ', e_type, e_value)
            if e_type == KeyboardInterrupt:
                print('Interrupting measurement')
                if not self.app.CheckAcquisition:
                    print('Aborting')
                    self.app.Abort
        else:
            print('Exiting')

        # Save document and close it
        self.doc.SaveAs(self.file_path + self.file_name + '.tnt')
        self.doc.Export(self.file_path + self.file_name + '.txt', 0)

        if not self.app.CloseActiveFile: print('Failed to close file')
    
        #return True # Supresses errors in code



if __name__ == "__main__":
    A = Tecmag()
    print(A.app.GetDocumentList)

    A.T2(A.Test_params())
    #A.T1(A.Test_params())

    #print(type(A.data_T2), len(A.data_T2), A.data_T2)
    #print('ok\n', type(A.data_T2[0]), A.data_T2[0])
    

    #file_path = 'C:\\TNMR\\data\\Nejc test\\Cu NQR\\'
    #file_name = 'test_T1'

    file_path = 'C:\\TNMR\\data\\Nejc\\210202_As_BCAO_NQR\\'
    file_name = 'T1_As_0T_200K'

    #with TNMR_document(file_path, file_name, A.app) as x:
    #    time.sleep(2)