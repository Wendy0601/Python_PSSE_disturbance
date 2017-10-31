import os, sys, collections
import read_rawdata
# =====================================================================================================

def check_psse_example_folder():
    # if called from PSSE's Example Folder, create report in subfolder 'Output_Pyscript'

    outdir = os.getcwd()
    cwd = outdir.lower()
    i = cwd.find('pti')
    j = cwd.find('psse')
    k = cwd.find('example')
    if i>0 and j>i and k>j:     # called from Example folder
        outdir = os.path.join(outdir, 'Output_Pyscript')
        if not os.path.exists(outdir): os.mkdir(outdir)

    return outdir

# =============================================================================================

def get_demotest_file_names(outpath,n):

    if outpath:
        outdir = outpath
    else:
        outdir = check_psse_example_folder()
    outfile=[]
    for i in range(n):
        outfile.append(os.path.join(outdir,'Line_fault'+str(i+1)+'.out')) 
    prgfile  = os.path.join(outdir,'Line_trip_progress.txt')

    return outfile, prgfile

# =============================================================================================
# Run Dynamic simulation on SAVNW to generate .out files

def run_savnw_simulation(datapath, outfile, prgfile):

    import psspy
    psspy.psseinit()

    savfile = 'Converted_NETS-NYPS 68 Bus System_C.sav'
    snpfile = 'NETS-NYPS 68 Bus System.snp'

    if datapath:
        savfile = os.path.join(datapath, savfile)
        snpfile = os.path.join(datapath, snpfile)

    psspy.lines_per_page_one_device(1,90)
    psspy.progress_output(2,prgfile,[0,0]) # directly output to file 
    

    ierr = psspy.case(savfile)
    if ierr:
        psspy.progress_output(1,"",[0,0])
        print(" psspy.case Error")
        return
    ierr = psspy.rstr(snpfile)
    if ierr:
        psspy.progress_output(1,"",[0,0])
        print(" psspy.rstr Error")
        return

    # branches
    ibus,jbus,id=read_rawdata.branch_bus()
    for i,gener in enumerate(all_gener):
        psspy.case(savfile)
        psspy.rstr(snpfile)
        psspy.strt(0,outfile[i]) 
        psspy.run(0, 1.0,1000,1,0)
        dist_branch_fault(ibus[i], jbus[i], id[i])
        psspy.run(0, 1.2,1000,1,0)
        psspy.dist_clear_fault(1)
        psspy.run(0, 5.0,1000,1,0)

    psspy.lines_per_page_one_device(2,10000000)#Integer DEVICE Indicates which of the four output devices is to be processed (input;
    #1 for disk files.
    #2 for the report window.
    #3 for the first primary hard copy output device.
    #4 for the second primary hard copy output device.
    psspy.progress_output(1,"",[0,0])
    return outfile,prgfile

# =============================================================================================
# 0. Run savnw dynamics simulation to create .out files

def test0_run_simulation(datapath=None, outpath=None):
    n_gener=63
    outfile, prgfile = get_demotest_file_names(outpath,n_linetrip)
    outfile,prgfile=run_savnw_simulation(datapath, outfile, prgfile)
    print(" Done SAVNW dynamics simulation")
    return outfile, prgfile

# =============================================================================================
# 1. Data extraction/information

def test1_data_extraction(outpath=None, show=True,outfile=None):

    import dyntools

    #outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath,3)

    # create object
    for i in range(len(outfile)):
        chnfobj = dyntools.CHNF(outfile[i])
        sh_ttl1, ch_id1, ch_data = chnfobj.get_data()
        sh_ttl2, ch_id2 = chnfobj.get_id()
        ch_range = chnfobj.get_range()
        ch_scale = chnfobj.get_scale()
        chnfobj.xlsout(show=show) 
 
# =============================================================================================

if __name__ == '__main__':
    import os
    import sys
    sys.path.append(r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27');
    os.environ['PATH']+=';'+r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27';  #or where else you find the psspy.pyc
    sys.path.append(r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN');
    os.environ['PATH']+=';'+r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN';
    import psspy
    import redirect 
    #(a) Run one test a time
    # Need to run "test0_run_simulation(..)" before running other tests.
    # After running "test0_run_simulation(..)", run other tests one at a time.
    datapath = r'C:\Program Files (x86)\PTI\PSSEXplore34\EXAMPLE'#None
    outpath  = r'C:\Python27\Ex_psse_python\outfile\generator_trip'
    show     = False # True  --> create, save and show Excel spreadsheets and Plots when done
                        # False --> create, save but do not show Excel spreadsheets and Plots when done   

    outfile,prgfile=test0_run_simulation(datapath, outpath)

    test1_data_extraction(outpath, show,outfile)  
# =============================================================================================
