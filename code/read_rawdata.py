def branch_bus():
    filename='C:\\Python27\\Ex_psse_python\\datafile\\Network_data_68bus.raw'
    with open(filename,"r") as rawf:
            line = ""
            #first three lines are bogus
            for i in range(1,3):
                rawf.readline();
            #bus_data
            _nbus = -1
            while True:
                line = rawf.readline()
                if "END OF BUS DATA" in line:
                    break
                _nbus += 1             
                
            #skip everything until you get to the load data
            while "BEGIN LOAD DATA" not in line:
                line = rawf.readline()
            #load and pevs data
            _loads = []
            while True:
                line = rawf.readline()
                if "END OF LOAD DATA" in line:
                    break
                #parse the data
                sline = str.split(line,',')
                bus = int(sline[0])
                id = sline[1].replace('\'','')
                P = float(sline[9])+float(sline[7])+float(sline[5])
                Q = float(sline[10])+float(sline[8])+float(sline[6])
                
            #skip everything until you get to the generator data
            while "BEGIN GENERATOR DATA" not in line:
                line = rawf.readline()
            #generator data
            _generators = []
            _machine_loads = []
            while True:
                    line = rawf.readline() 
                    if "END OF GENERATOR DATA" in line:
                        break
                    #parse raw data
                    sline = str.split(line,',')
                    bus = int(sline[0])
                    id = int(sline[1].replace('\'',''))
                    P = float(sline[2])
                    Q = float(sline[3]) 
            #skip everything until you get to the branch data 
            while "BEGIN BRANCH DATA" not in line:
                line = rawf.readline()
            #branch data
            _branches = []
            n_line=-1
            ibus=[]
            jbus=[]
            idbus=[]
            R=[]
            X=[]
            while True:
                line = rawf.readline()
                if "END OF BRANCH DATA" in line:
                    break
                #parse raw data
                sline = str.split(line,',')
                ibus.append(int(sline[0])) 
                jbus .append(int(sline[1]))
                idbus.append(int(sline[2].replace('\'','')))
                R.append(float(sline[3]))
                X.append(float(sline[4]))
                n_line+=1
 return ibus,jbus,idbus

