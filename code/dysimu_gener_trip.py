#[dyntools_demo.py]  Demo for using functions from dyntools module
# ====================================================================================================
'''
'dyntools' module provide access to data in PSS(R)E Dynamic Simulation Channel Output file.
This module has functions:
- to get channel data in Python scripts for further processing
- to get channel information and their min/max range
- to export data to text files, excel spreadsheets
- to open multiple channel output files and post process their data using Python scripts
- to plot selected channels
- to plot and insert plots in word document

This is an example file showing how to use various functions available in dyntools module.

Other Python modules 'matplotlib', 'numpy' and 'python win32 extension' are required to be
able to use 'dyntools' module.
Self installation EXE files for these modules are available at:
   PSSE User Support Web Page and follow link 'Python Modules used by PSSE Python Utilities'.

- The dyntools is developed and tested using these versions of with matplotlib and numpy.
  When using Python 2.5
  Python 2.5 matplotlib-1.1.1
  Python 2.5 numpy-1.7.0
  Python 2.5 pywin32-218

  When using Python 2.7
  Python 2.7 matplotlib-1.2.0
  Python 2.7 numpy-1.7.0
  Python 2.7 pywin32-218

  Versions later than these may work.

---------------------------------------------------------------------------------
How to use this file?
- Open Python IDLE (or any Python Interpreter shell)
- Open this file
- run (F5)

Note: Do NOT run this file from PSS(R)E GUI. The 'xyplots' function from dyntools can
save plots to eps, png, pdf or ps files. However, creating only 'eps' files from inside
PSS(R)E GUI works. This is because different backends matplotlib uses to create different
plot types.
When run from any Python interpreter (outside PSS(R)E GUI) plots can be saved to any of
these four (eps, png, pdf or ps) file types.

Get information on functions available in dyntools as:
import dyntools
help(dyntools)

'''

import os, sys, collections

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
        outfile.append(os.path.join(outdir,'generator_trip'+str(i)+'.out')) 
    prgfile  = os.path.join(outdir,'dyntools_demo_progress.txt')

    return outfile, prgfile

# =============================================================================================
# Run Dynamic simulation on SAVNW to generate .out files

def run_savnw_simulation(datapath, outfile, prgfile):

    import psspy
    psspy.psseinit()

    savfile = 'savcnv.sav'
    snpfile = 'savnw.snp'

    if datapath:
        savfile = os.path.join(datapath, savfile)
        snpfile = os.path.join(datapath, snpfile)

    psspy.lines_per_page_one_device(1,90)
    psspy.progress_output(2,prgfile,[0,0]) # directly output to file
    psspy.chsb(0,1,[-1,-1,-1,1,13,0])
    

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

    # run generator trip automatically 
    all_gener=[206,211,102,101,3018,3011]
    for i,gener in enumerate(all_gener):
        psspy.case(savfile)
        psspy.rstr(snpfile)
        psspy.strt(0,outfile[i])
        #psspy.chsb(0,1,[-1,-1,-1,1,13,0]) 
        psspy.run(0, 1.0,1000,1,0)
        psspy.dist_machine_trip(gener,'1')
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
    n_gener=6
    outfile, prgfile = get_demotest_file_names(outpath,n_gener)
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
