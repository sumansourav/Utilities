import os
import argparse
try:
    import pandas as pd
except:
    print("pandas is missing. Install using: pip install pandas")
    exit()

__author__ = 'sumansourav'

parser = argparse.ArgumentParser(description='Merge all excel files in current folder')
parser.add_argument('sheet_name', help='Enter the common sheet name to be read and merged from all files')
parser.add_argument('-o', '--output_file', default='output.xlsx',
                    help='Enter the output excel file where all data has to be merged into. Ex: output.xlsx')
args = parser.parse_args()

# This script can merge excel files of the same format and write them to an output xlsx file.
# Get the script and excel files to be merged into same folder
# output file is created in an 'output' sub-folder

path = os.getcwd()
files = os.listdir(path)
files_xls = [f for f in files if f[-4:] == 'xlsx']  # Change this if the format of files are different

print("Files to be merged:\n", files_xls)

df = pd.DataFrame()
for f in files_xls:
    data = pd.read_excel(f, args.sheet_name)
    df = df.append(data)

# Create the output in the sub-folder,
# so that re-running the code with the same params does not count in the output file.

merged_output_folder = 'merged_output'
if not os.path.exists(merged_output_folder):
    os.makedirs(merged_output_folder)

writer = pd.ExcelWriter('{0}/{1}'.format(merged_output_folder, args.output_file), engine='xlsxwriter')  # change engine if file format is different
df.to_excel(writer, sheet_name=args.sheet_name)
writer.save()
print("Output file created: ", args.output_file)
