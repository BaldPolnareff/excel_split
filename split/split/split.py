import pandas as pd 
import os
from utils import *

#filepath = '../excel_in/example_merged.xlsx'
filepath = '../../excel_in/AU5000-GDAT-ASSY.xls'
output_path = '../../excel_out'

# split_excel(filepath, output_path)
split_excel_xlwings(filepath, output_path)




