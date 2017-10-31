import os
import sys
import numpy as np 
PYTHONPATH = r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
MODELFOLDER = r'C:\Program Files (x86)\PTI\PSSEXplore34\MYMODEL'

sys.path.append(r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN") # http://wangwei007.blog.51cto.com/68019/1104940 talk about the os.path usage
os.environ['PATH']=(r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN;" +os.environ['PATH'])
sys.path.append(PYTHONPATH)
os.environ['PATH'] += ';' + PYTHONPATH

import excelpy
import dyntools
 
"""
Comienza a extraer la información para cada uno de los *.out
"""
OutNames ="NETS-NYPS 68 Bus System_1"#[Out1, Out2]
Ruta = "C:\\Program Files (x86)\\PTI\\PSSE34\\EXAMPLE\\New England 68-Bus Test System\\PSSE"

rootDir = Ruta + "\\" 
logFile = file(rootDir + "Reporte_" + OutNames + "_Din.txt", "w")
logFile1 = file(rootDir + "Errores_" + OutNames + "_Din.txt", "w")
sys.stdout = logFile  # Las salidas
sys.stderr = logFile1  # Los errores
logFile.close()
logFile1.close()

outFile = dyntools.CHNF("C:\\Program Files (x86)\\PTI\\PSSE34\\EXAMPLE\\New England 68-Bus Test System\\PSSE\\NETS-NYPS 68 Bus System_1.out")
#Extrae información del archivo *.out
#short_title, chanid_dict, chandata_dict = outFile.get_data()
outFile.xlsout( channels='', show=True, overwritesheet=True);
#report=Ruta+"\\"+"ex1.xlsx"
#xl=excelpy.workbook(xlsfile=report, sheet="", overwritesheet=True, mode='w')
#xl.show()
#xl.set_range(1, 'a', zip([short_title]))
#xl.set_range(1, 'b', zip([chanid_dict]))
#xl.set_range(1, 'c', [chandata_dict])
#xl.save() 
 