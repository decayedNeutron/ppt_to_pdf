import win32com.client
import os
from pywintypes import com_error
import time

PARENT_DIR = os.path.abspath(os.curdir)




# fil = input('Enter the Name of PPTX  File: ')
#num = int(input('No. of Students in Excel File: '))

PPT_EXT = r'.ppt'



PDF_EXT = r'.pdf'


mode = 0o666
NEW_DIR = os.path.join(PARENT_DIR,r'PDFs')
files = os.listdir(PARENT_DIR+"\\")
err = ""

pwp = win32com.client.Dispatch("Powerpoint.Application")
#pwp.Visible = 1
count = 0

try:
    print('Printing PDFs')
    
    for fil in files:
        if not os.path.isdir(NEW_DIR):
            os.mkdir(NEW_DIR)
        if not fil.endswith((".ppt",".pptx")):
            continue
        PATH_PPT = os.path.join(PARENT_DIR,fil)
        name1 = os.path.splitext(fil)[0]
        PATH_PDF = os.path.join(NEW_DIR,name1)
        print(PATH_PDF)
        try:
            slides = pwp.Presentations.Open(PATH_PPT,WithWindow=0)
            slides.SaveAs(PATH_PDF+PDF_EXT,32)
            print(PATH_PDF+PDF_EXT)
        except Exception as e:
            err = str(e)
        finally: 
            slides.Close()
        count=count+1
            
            
except Exception as e:
    print('Failed to save PDFs')
    log_file = open("log.txt","a")
    log_file.write("\n")
    log_file.write(f"[{time.ctime(time.time())}] -- ")
    log_file.write(str(e)+"---"+err)
    
else:
    print(f'{count} PDFs succesfully saved')
finally:
    pwp.Quit()
