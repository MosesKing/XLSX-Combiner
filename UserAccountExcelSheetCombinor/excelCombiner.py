# Author: Moeses King
# Date: 01/31/2019

# Central Contention:
#     This tool was created by Moeses King. In order to be a tool that could be used to combine
#     multiple excel sheets into one large one.It is done by parsing all of the excel files that are located
#     in InputFiles directory here in the project folder, and writing it all out into one excel file

# Let's Import Our Needed Libraries

import pandas as pd
import numpy as np
import xlsxwriter
import xlrd
import glob


# Let's gather and parse all of our excel files from the input folder directory into one place
all_data = pd.DataFrame()
for f in glob.glob('inputFiles/*.xlsx'):
   df = pd.read_excel(f)
   all_data = all_data.append(df, ignore_index=True)

#Finally, let's write everything out into one file.
writer = pd.ExcelWriter('ExcelCombined.xlsx', engine='xlsxwriter')
all_data.to_excel(writer, sheet_name='Sheet1')
writer.save()