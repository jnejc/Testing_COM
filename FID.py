''' Routines for importing an FID from .txt or direct transfer '''
# Imports
import os # For compiling directory paths
import numpy as np  # Numpy for numerical operations

import matplotlib   # Plotting
import matplotlib.pyplot as plt     # Short notation for plots

# Global variables



class FID():
    '''Class for holding FID adress and data manipulation'''
    def __init__(self, file_name, file_dir, params=None):
        '''Initialization routine'''
        # Remembers name and directory
        self.file_name = file_name
        self.file_dir = file_dir
        # Try to accept given params
        if not params:
            self.params = params
        else:
            self.params = dict()

    
    def Import_file(self):
        '''Imports data from a TNMR text file assumes 1D for FID'''
        with open(os.path.join(self.file_dir, self.file_name)) as f:
            lines = list(f)

        # strip /n from end of line
        lines = [line.strip() for line in lines]

        # get the head strings using the first empty line
        heads = list()
        while len(lines) > 0:
            line = lines.pop(0)
            if line == '': break # Stop after empty line
            heads.append(line)
        print('Heads: ', heads)

        # Find the 1D 2D 3D 4D depths
        self.depths = list()
        for value in heads[1:-1]: # Skip first and last row
            self.depths.append(int(value.split('\t')[-1]))
        dimension = len(self.depths)
        print('Dimension: ', dimension)
        if not dimension == 1: # Fix the 1D case
            print('Was not given a 1D file! Taking only first case')

        # Assume 1D
        self.x1_list = []
        self.x2_list = []

        # Extract the data
        for i in range(depths[0]):
            split = lines[i].split('\t')
            self.x1_list.append(float(split[0]))
            self.x2_list.append(float(split[1]))
            
        # Make a complex numpy array
        self.x_list = ( np.array(self.x1_list[i]) 
            + 1j*np.array(self.x2_list[i]) )

        # Dont have the params....


    # Dont use for FID
"""
    def Import_file_nD(self):
        '''Imports data from a TNMR text file'''
        with open(os.path.join(self.file_dir, self.file_name)) as f:
            lines = list(f)

        # strip /n from end of line
        lines = [line.strip() for line in lines]

        # get the head strings using the first empty line
        heads = list()
        while len(lines) > 0:
            line = lines.pop(0)
            if line == '': break # Stop after empty line
            heads.append(line)
        print('Heads: ', heads)

        # Find the 1D 2D 3D 4D depths
        self.depths = list()
        for value in heads[1:-1]: # Skip first and last row
            self.depths.append(int(value.split('\t')[-1]))
        dimension = len(self.depths)
        print('Dimension: ', dimension)
        if dimension == 1: # Fix the 1D case
            self.depths.append(1)

        # Assume 2D
        self.x1_list = [[] for i in range(self.depths[1])]
        self.x2_list = [[] for i in range(self.depths[1])]

        # Extract the data
        for i,line in enumerate(lines):
            if line == '': break # Stop after empty line
            split = line.split('\t')
            self.x1_list[i//self.depths[0]].append(float(split[0]))
            self.x2_list[i//self.depths[0]].append(float(split[1]))
            
        # Make a complex numpy array
        self.x_list = [
            np.array(self.x1_list[i]) + 1j*np.array(self.x2_list[i])
            for i in range(self.depths[1])
            ]

        # Dont have the params....
"""


    def Import_data(self, data, params=None):
        '''Imports FID from techmag directly
        Assumes single FID'''
        # Update parameter dictionary with given values
        if not params:
            self.params.update(params)

        # data format: tuple of floats (R, I, R, I, R, I, ...)

        self.x1_list = []
        self.x2_list = []

        # Split real and imaginary
        for i in range(len(data)//2):
            self.x1_list.append(data[2*i])
            self.x2_list.append(data[2*i+1])

        # Make a complex numpy array
        self.x_list = ( np.array(self.x1_list[i]) 
            + 1j*np.array(self.x2_list[i]) )


       
if __name__ == "__main__":
    '''Executes if executed directly'''
    # testing of script
    file_dir = 'C:\\TNMR\\data\\Nejc test\\Cu NQR\\Pulsing'
    file_name = '201001_FID.txt'

    A = FID(file_name, file_dir)
    A.Import_file()

    # Higher dimension
    file_dir2 = 'C:\\TNMR\\data\\Nejc test\\Cu NQR'
    file_name2 = 'test_T1.txt'

    B = FID(file_name2, file_dir2)
    B.Import_file()


