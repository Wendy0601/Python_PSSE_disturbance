import os
import sys
import numpy as np 
PYTHONPATH = r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27'
MODELFOLDER = r'C:\Program Files (x86)\PTI\PSSEXplore34\MYMODEL'

sys.path.append(r"C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN") # http://wangwei007.blog.51cto.com/68019/1104940 talk about the os.path usage
os.environ['PATH']=(r"C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN;" +os.environ['PATH'])
sys.path.append(PYTHONPATH)
os.environ['PATH'] += ';' + PYTHONPATH

import excelpy
import dyntools
 
OutNames ="C:\python27\linetrip_no_fault"#[Out1, Out2]
Ruta = os.getcwd()


outFile = dyntools.CHNF(OutNames + ".out")
#Extrae información del archivo *.out
short_title, chanid_dict, chandata_dict = outFile.get_data()  
a=zip([chanid_dict])
print a
print len(a)
report = Ruta + '\\' + 'my.xlsx' 
xl = excelpy.workbook(xlsfile=report, sheet="Voltage",
                      overwritesheet=True, mode='w')
#xl.show()
#xl.worksheet_add_end(sheet="Voltage")
#address=(1,'b')
#xl.autofit_columns(address, sheet=chandata_dict)
#xl.set_range(1, 'a', zip([short_title]))
#xl.set_range(1, 'b', zip([chanid_dict]))
#xl.set_range(1, 'c', zip([chandata_dict]) )
#xl.save()
#xl.close()
#xl.show()