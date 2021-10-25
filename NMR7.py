# Methods for exporting into Legacy 7NMR format

# Imports
import os

from T1 import ToN



def Legacy_export(data, params):
    '''Takes the data and parameters and exports in 7NMR format'''
    file_path = params['file_path'] + 'export\\'
    file_name = params['file_name'] + '.DAT'

    print('Writing to file: ', os.path.join(file_path, file_name))

    L = Legacy_params()

    # Open writeable file
    with open(os.path.join(file_path, file_name),'w') as f:
        f.write('[PARAMETERS]\n')
        for key, value in L.PARAMETERS.items():
            # Write in scientific format
            num = "{:.3E}".format(ToN(params[value]))
            f.write('='.join([key, num])+'\n')

        f.write('[ADDITIONAL]\n')
        for key, value in L.ADDITIONAL.items():
            f.write('='.join([key, params[value]])+'\n')

        f.write('[VARIABLES]\n')
        for key, value in L.VARIABLES.items():
            f.write('='.join([key, params[value]])+'\n')

        f.write('[PPFILE]\n')

        f.write('[DATA]\n')
        for i in range(int(params['TD'])):
            # Add space for positive numbers
            # Normalize by number of scans
            f.write(' '.join([
                '{: f}'.format(data[2*i]/int(params['NS'])),
                '{: f}'.format(data[2*i+1]/int(params['NS']))
                ]) + '\n')

    print('File writing complete')



def Legacy_G(d5_list, params):
    '''Generates a .G file. ADD columns later!'''
    file_path = params['file_path'] + 'export\\'
    file_name = params['file_name'] + '-G.DAT'

    with open(os.path.join(file_path, file_name),'w') as f:
        f.write('D5\n')

        for d5 in d5_list.split(' '):
            f.write('{: E}'.format(ToN(d5)) + '\n')



class Legacy_params():
    '''Class holding the rules on how to transcribe
        params dictionary to 7NMR paramteter naming'''
    def __init__(self):
        #[PARAMETERS]
        self.PARAMETERS = {
            'DW': 'DW', 
            'TAU': 'TAU',
            'TD': 'TD',    # Time domain
            #'BS': '2048',           # Unknown
            'NS': 'NS',             # Number of scans
            'NSC': 'NS',            # Unknown
            'FR': 'FR',             # Frequency
            #'GAIN': '1',           # Gain, obsolete
            'D0': 'D9',             # Last delay older name
            'D1': 'd123',           # Echo pi/2 pulse
            'D2': 'd123',           # Echo pi pulse
            'D3': 'd123',           # Inversion pi pulse
            #'D4': 'd123',           # Fillers
            'D5': 'D5',             # Delay after inversion pulse
            #'D6': 'd123',           # Fillers
            #'D7': 'd123',           # Fillers
            #'D8': 'd123',           # Fillers
            'D9': 'D9'              # Last delay
            #'C0'
            #'SEGTD'
            #'SEGOFF'
            #'SEGNUM'
        }

        #[ADDITIONAL]
        self.ADDITIONAL = {
            'PPFILE': 'pulse_file',
            #'DATEWR'
            #'TIMEWR'
            'DATESTA': 'DATESTA',
            'TIMESTA': 'TIMESTA',
            #'DATEEND'
            'TIMEEND': 'TIMEEND',
            #'CURR_Z0'
            #'CURR_Z1'
            #'TE'
            ## New ones
            'a123': 'a123',
            'atn1': 'atn1',
            'atn23': 'atn23',
            'RecGain': 'RecGain',
            'RecPh': 'RecPh',
            'Filter': 'Filter',
            'Pcryo': 'Pcryo',
            'ad': 'ad',
            'rd': 'rd',
            'pre': 'pre'
        }

        #[VARIABLES]
        self.VARIABLES = {
            #'_USER'
            #'_EMAIL_F'
            '_ITC_R0': 'Tset',   # Set temperature
            '_ITC_R1': 'Tsens',   # Cryostat
            '_ITC_R2': 'Tprobe',  # Probe temp
            #'_ITC_R3'
            #'_ITC_R4'
            #'_ITC_V'
            #'_ITC_X'
            #'_FW'
            #'_ORIENT'
            #'_THETA'
            #'_TITLE'
            #'_DESCR1'
            #'_DESCR2'
            #'_FR_LARM'
            #'_FR_SYNT'
            #'_TRIMLEV'
            #'_IPS_R0': 'Fset'   # Set field
            '_FIELD': 'Fpers'  # Current persistent field
        }


