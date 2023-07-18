import sys
import os,os.path
import comtypes.client

wdFormatPDF = 17

input_dir = 'C:\sankar\Python\Codes\Rtf2pdf\1843.rtf'
output_dir = 'C:\sankar\Python\Codes\Rtf2pdf'

for subdir, dirs, files in os.walk(output_dir):
    for file in files:
        in_file = os.path.join(subdir, file)
        output_file = file.split('.')[0]
        out_file = output_dir+output_file+'.pdf'
        word = comtypes.client.CreateObject('Word.Application')

        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        
print("newly added")        
print("added again")