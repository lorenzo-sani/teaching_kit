from pptx_extract import pptx_extract
import os

path = 'files/ppt_presentations'

folder = os.fsencode(path)
filenames = []

for file in os.listdir(folder):
    
    filename = os.fsdecode(file)
    if filename == 'ppt_figure.pptx' or filename == 'ppt_table.pptx':
        print("skip "+filename)
        continue
    else:
        pptx_extract(path,filename)

print("converted all presentations in "+path+" to markdown presentations")
    
