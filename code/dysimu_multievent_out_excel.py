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

import os, sys, collections#?

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

def get_demotest_file_names(outpath):

    if outpath:
        outdir = outpath
    else:
        outdir = check_psse_example_folder()

    outfile1 = os.path.join(outdir,'dyntools_demo_bus154_fault.out')
    outfile2 = os.path.join(outdir,'dyntools_demo_bus3018_gentrip.out')
    outfile3 = os.path.join(outdir,'dyntools_demo_brn3005_3007_trip.out')
    prgfile  = os.path.join(outdir,'dyntools_demo_progress.txt')

    return outfile1, outfile2, outfile3, prgfile

# =============================================================================================
# Run Dynamic simulation on SAVNW to generate .out files

def run_savnw_simulation(datapath, outfile1, outfile2, outfile3, prgfile):

    import psspy
    psspy.psseinit()

    savfile = 'Converted_NETS-NYPS 68 Bus System_C.sav'
    snpfile = 'NETS-NYPS 68 Bus System.snp'

    if datapath:
        savfile = os.path.join(datapath, savfile)
        snpfile = os.path.join(datapath, snpfile) #why produce these two kinds of files?

    psspy.lines_per_page_one_device(1,90)
    psspy.progress_output(2,prgfile,[0,0])

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

# fault + line trip
    psspy.strt(0,outfile1)
    psspy.run(0, 1.0,1000,1,0)
    psspy.dist_bus_fault(52,1, 138.0,[0.0,-0.2E+10])
    psspy.run(0, 1.1,1000,1,0)
    psspy.dist_clear_fault(1) 
    psspy.dist_branch_trip(52,55,'1')
    psspy.run(0,1.2,1000,1,0)
    psspy.dist_machine_trip(1,'1')
    psspy.run(0, 5.0,1000,1,0)

# line trip (with faults) + generator trip   
    psspy.case(savfile)
    psspy.rstr(snpfile)
    psspy.strt(0,outfile2)
    psspy.run(0, 1.0,1000,1,0)
    psspy.dist_bus_fault(52,1,138.0,[0.0,-0.2E+10])
    psspy.run(0,1.1,1000,1,0)
    psspy.dist_clear_fault(1)
    psspy.run(0,1.2,1000,1,0)
    psspy.dist_machine_trip(8,'1')
    psspy.run(0, 5.0,1000,1,0)

    psspy.case(savfile)
    psspy.rstr(snpfile)
    psspy.strt(0,outfile3)
    psspy.run(0, 1.0,1000,1,0)
    psspy.dist_branch_trip(32,33,'1')
    psspy.run(0, 5.0,1000,1,0)

    psspy.lines_per_page_one_device(2,10000000)
    psspy.progress_output(1,"",[0,0])

# =============================================================================================
# 0. Run savnw dynamics simulation to create .out files

def test0_run_simulation(datapath=None, outpath=None):

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    run_savnw_simulation(datapath, outfile1, outfile2, outfile3, prgfile)

    print(" Done SAVNW dynamics simulation")

# =============================================================================================
# 1. Data extraction/information

def test1_data_extraction(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    # create object
    chnfobj = dyntools.CHNF(outfile1)# this is used to read the .out file to excel file by dyntools.CHNF

    print '\n Testing call to get_data'
    sh_ttl, ch_id, ch_data = chnfobj.get_data()
    print sh_ttl
    print ch_id

    print '\n Testing call to get_id'
    sh_ttl, ch_id = chnfobj.get_id()
    print sh_ttl
    print ch_id

    print '\n Testing call to get_range'
    ch_range = chnfobj.get_range()
    print ch_range

    print '\n Testing call to get_scale'
    ch_scale = chnfobj.get_scale()
    print ch_scale

    print '\n Testing call to print_scale'
    chnfobj.print_scale()

    print '\n Testing call to txtout'
    chnfobj.txtout(channels=[1,4])

    print '\n Testing call to xlsout'
    chnfobj.xlsout(show=show)

# =============================================================================================
# 2. Multiple subplots in a figure, but one trace in each subplot
#    Channels specified with normal dictionary

# See how "set_plot_legend_options" method can be used to place and format legends

def test2_subplots_one_trace(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2)

    chnfobj.set_plot_page_options(size='letter', orientation='portrait')
    chnfobj.set_plot_markers('square', 'triangle_up', 'thin_diamond', 'plus', 'x',
                             'circle', 'star', 'hexagon1')
    chnfobj.set_plot_line_styles('solid', 'dashed', 'dashdot', 'dotted')
    chnfobj.set_plot_line_colors('blue', 'red', 'black', 'green', 'cyan', 'magenta', 'pink', 'purple')

    optnfmt  = {'rows':3,'columns':2,'dpi':300,'showttl':True, 'showoutfnam':True, 'showlogo':True,
                'legendtype':1, 'addmarker':True}

     #optnchn1 = {1:{'chns':[1]},2:{'chns':[2]},3:{'chns':[3]},4:{'chns':[4]},5:{'chns':[5]}}
    optnchn1 = {1:{'chns':1,  'title':'Ch#1,bus154_fault'}, 2:{'chns':6,  'title':'Ch#6,bus154_fault'}, 3:{'chns':11, 'title':'Ch#11,bus154_fault'},4:{'chns':16, 'title':'Ch#16,bus154_fault'},5:{'chns':26, 'title':'Ch#26,bus154_fault'},6:{'chns':40, 'title':'Ch#40,bus154_fault'}, 
                }
    pn,x     = os.path.splitext(outfile1)
    pltfile1 = pn+'.png'

    optnchn2 = {1:{'chns':{outfile2:1}, 'title':'Channel 1 from bus3018_gentrip'},
                2:{'chns':{outfile2:6}, 'title':'Channel 6 from bus3018_gentrip'},
                3:{'chns':{outfile2:11}},
                4:{'chns':{outfile2:16}},
                5:{'chns':{outfile2:26}},
                6:{'chns':{outfile2:40}},
                }
    pn,x     = os.path.splitext(outfile2)
    pltfile2 = pn+'.png'

    figfiles1 = chnfobj.xyplots(optnchn1,optnfmt,pltfile1)
    figfiles2 = chnfobj.xyplots(optnchn2,optnfmt,pltfile2)
    chnfobj.set_plot_legend_options(loc='lower center', borderpad=0.2, labelspacing=0.5,
                                    handlelength=1.5, handletextpad=0.5, fontsize=8, frame=False)

    optnfmt  = {'rows':3,'columns':1,'dpi':300,'showttl':False, 'showoutfnam':True, 'showlogo':False,
                'legendtype':2, 'addmarker':False}

    

    if figfiles1 or figfiles2:
        print 'Plot fils saved:'
        if figfiles1: print '   ', figfiles1[0]
        if figfiles2: print '   ', figfiles2[0]

    if show:
        chnfobj.plots_show()
    else:
        chnfobj.plots_close()

# =============================================================================================
# 3. Multiple subplots in a figure and more than one trace in each subplot
#    Channels specified with normal dictionary

def test3_subplots_mult_trace(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2, outfile3)

    chnfobj.set_plot_page_options(size='letter', orientation='portrait')
    chnfobj.set_plot_markers('square', 'triangle_up', 'thin_diamond', 'plus', 'x',
                             'circle', 'star', 'hexagon1')
    chnfobj.set_plot_line_styles('solid', 'dashed', 'dashdot', 'dotted')
    chnfobj.set_plot_line_colors('blue', 'red', 'black', 'green', 'cyan', 'magenta', 'pink', 'purple')

    optnfmt  = {'rows':2,'columns':2,'dpi':300,'showttl':False, 'showoutfnam':True, 'showlogo':False,
                'legendtype':2, 'addmarker':True}

    optnchn1 = {1:{'chns':[1]},2:{'chns':[2]},3:{'chns':[3]},4:{'chns':[4]},5:{'chns':[5]}}
    pn,x     = os.path.splitext(outfile1)
    pltfile1 = pn+'.png'

    optnchn2 = {1:{'chns':{outfile2:1}},
                2:{'chns':{'v82_test1_bus_fault.out':3}},
                3:{'chns':4},
                4:{'chns':[5]}
               }
    pn,x     = os.path.splitext(outfile2)
    pltfile2 = pn+'.pdf'

    optnchn3 = {1:{'chns':{'v80_test1_bus_fault.out':1}},
                2:{'chns':{'v80_test2_complex_wind.out':[1,5]}},
                3:{'chns':{'v82_test1_bus_fault.out':3}},
                5:{'chns':[4,5]},
               }
    pn,x     = os.path.splitext(outfile3)
    pltfile3 = pn+'.eps'

    figfiles1 = chnfobj.xyplots(optnchn1,optnfmt,pltfile1)
    figfiles2 = chnfobj.xyplots(optnchn2,optnfmt,pltfile2)
    figfiles3 = chnfobj.xyplots(optnchn3,optnfmt,pltfile3)

    figfiles = figfiles1[:]
    figfiles.extend(figfiles2)
    figfiles.extend(figfiles3)
    if figfiles:
        print 'Plot fils saved:'
        for f in figfiles:
            print '    ', f

    if show:
        chnfobj.plots_show()
    else:
        chnfobj.plots_close()

# =============================================================================================
# 4. Multiple subplots in a figure, but one trace in each subplot
#    Channels specified with Ordered dictionary

def test4_subplots_mult_trace_OrderedDict(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2, outfile3)

    chnfobj.set_plot_page_options(size='letter', orientation='portrait')
    chnfobj.set_plot_markers('square', 'triangle_up', 'thin_diamond', 'plus', 'x',
                             'circle', 'star', 'hexagon1')
    chnfobj.set_plot_line_styles('solid', 'dashed', 'dashdot', 'dotted')
    chnfobj.set_plot_line_colors('blue', 'red', 'black', 'green', 'cyan', 'magenta', 'pink', 'purple')

    optnfmt  = {'rows':1,'columns':1,'dpi':300,'showttl':False, 'showoutfnam':True, 'showlogo':False,
                'legendtype':2, 'addmarker':True}

    optnchn  = {1:{'chns':collections.OrderedDict([(outfile1,3), (outfile2,3), (outfile3,3)]),
                  }
               }
    p,nx     = os.path.split(outfile1)
    pltfile  = os.path.join(p, 'plot_chns_ordereddict.png')

    figfiles = chnfobj.xyplots(optnchn,optnfmt,pltfile)

    if show:
        chnfobj.plots_show()
    else:
        chnfobj.plots_close()

# =============================================================================================
# 5. Do XY plots and insert them into word file
# Does not work because win32 API to Word does not work.

def test5_plots2word(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2, outfile3)

    p,nx       = os.path.split(outfile1)
    docfile    = os.path.join(p,'savnw_response')
    overwrite  = True
    caption    = True
    align      = 'center'
    captionpos = 'below'
    height     = 0.0
    width      = 0.0
    rotate     = 0.0

    optnfmt  = {'rows':3,'columns':1,'dpi':300,'showttl':True}

    optnchn  = {1:{'chns':{outfile1:1,  outfile2:1,  outfile3:1} },
                2:{'chns':{outfile1:7,  outfile2:7,  outfile3:7} },
                3:{'chns':{outfile1:17, outfile2:17, outfile3:17} },
                4:{'chns':[1,2,3,4,5]},
                5:{'chns':{outfile2:[1,2,3,4,5]} },
                6:{'chns':{outfile3:[1,2,3,4,5]} },
               }
    ierr, docfile = chnfobj.xyplots2doc(optnchn,optnfmt,docfile,show,overwrite,caption,align,
                        captionpos,height,width,rotate)

    if not ierr:
        print 'Plots saved to file:'
        print '    ', docfile

# =============================================================================================
# Run all tests and save plot and report files.

def run_all_tests(datapath=None, outpath=None):

    show = False        # This must be false in this test.

    test0_run_simulation(datapath, outpath)

    test1_data_extraction(outpath, show)

    test2_subplots_one_trace(outpath, show)

    test3_subplots_mult_trace(outpath, show)

    test4_subplots_mult_trace_OrderedDict(outpath, show)

    test5_plots2word(outpath, show)

# =============================================================================================

if __name__ == '__main__':
    import os
    import sys
    sys.path.append(r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27');
    os.environ['PATH']+=';'+r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27';  #or where else you find the psspy.pyc
    sys.path.append(r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN');
    os.environ['PATH']+=';'+r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN';
    import psspy
    import redirect#psse34
    #import psse34

    #(a) Run one test a time
    # Need to run "test0_run_simulation(..)" before running other tests.
    # After running "test0_run_simulation(..)", run other tests one at a time.
    datapath = r'C:\Program Files (x86)\PTI\PSSE34\EXAMPLE\New England 68-Bus Test System\PSSE'#None
    outpath  = r'C:\Program Files (x86)\PTI\PSSE34\EXAMPLE\New England 68-Bus Test System\PSSE'
    show     = True     # True  --> create, save and show Excel spreadsheets and Plots when done
                        # False --> create, save but do not show Excel spreadsheets and Plots when done

    test0_run_simulation(datapath, outpath)

    test1_data_extraction(outpath, show)

    #test2_subplots_one_trace(outpath, show)

    #test3_subplots_mult_trace(outpath, show)

    #test4_subplots_mult_trace_OrderedDict(outpath, show)

    #test5_plots2word(outpath, show)

    #(b) Run all tests

    #run_all_tests(datapath, outpath)

# =============================================================================================
