#[pout_excel.py]  POWER FLOW RESULTS Exported to Excel Spreadsheet
# ====================================================================================================
'''
This is an example file showing how to use "subsystem data retrieval APIs (API Manual, Chapter 8")
from Python to export power flow results to excel spreadsheet.
    Input : Solved PSS(R)E saved case file name
    Output: Excel file name to save
    When 'savfile' is not provided, it uses Network Data from PSS(R)E memory.
    When 'xlsfile' is provided and exists, power flow results are saved in 'next sheet#' of 'xlsfile'.
    When 'xlsfile' is provided and does not exists, power flow results are saved in 'Sheet1' of 'xlsfile'.
    When 'xlsfile' is not provided, power flow results are saved in 'Sheet1' of 'Book#.xls' file.

The subsystem data retrieval APIs return values as List of Lists. For example:
When "abusint" API is called with "istrings" as defined below:
    istrings = ['number','type','area','zone','owner','dummy']
    ierr, idata = psspy.abusint(sid, flag_bus, istrings)
The returned list will have format:
    idata=[[list of 'number'],[list of 'type'],[],[],[],[list of 'dummy']]

This example is written such that, such returned lists are converted into dictionary with
keys as strings specified in "istrings". This makes it easier to refer and use these lists.
    ibuses = array2dict(istrings, idata)

So ibuses['number'] gives the bus numbers returned by "abusint".

You need to have Win32 extensions for Python installed.
(http://sourceforge.net/projects/pywin32)

Refer to
Microsoft Excel Visual Basic Reference --> Reference -->Enumerations --> Microsoft Excel Constants
in MSDN library for various constants used in this example, viz., xlHAlignCenter = -4108

Notes:
(1)
This command enables whether to make Excel workbook visible while exporting data
    - Set to True if want see spredsheet
    - Set to False if just want to save data to spredsheet
xlApp.Visible = True

(2)
These two commands close the Excel workbook and quits Excel application opened when exporting data
xlBook.Close()
xlApp.Quit()

---------------------------------------------------------------------------------
How to use this file?

As showed in __main__ (end of this file)
- Enable PSSE version specific environment, as an example:
    import psse34

- call funtion
    pout_excel()
    # OR
    #pout_excel(savfile='savnw.sav', outpath=None, show=True)

'''
# ----------------------------------------------------------------------------------------------------
import sys, os
import win32com.client


# ----------------------------------------------------------------------------------------------------
def array2dict(dict_keys, dict_values):
    '''Convert array to dictionary of arrays.
    Returns dictionary as {dict_keys:dict_values}
    '''
    tmpdict = {}
    for i in range(len(dict_keys)):
        tmpdict[dict_keys[i].lower()] = dict_values[i]
    return tmpdict

# ----------------------------------------------------------------------------------------------------
def busindexes(busnum, busnumlist):
    '''Find indexes of a bus in list of buses.
    Returns list with indexes of 'busnum' in 'busnumlist'.
    '''
    busidxes = []
    startidx = 0
    buscounts = busnumlist.count(busnum)
    if buscounts:
        for i in range(buscounts):
            tmpidx = busnumlist.index(busnum,startidx)
            busidxes.append(tmpidx)
            startidx = tmpidx+1
    return busidxes

# ----------------------------------------------------------------------------------------------------
def exportedvalues():

    rowvars = ['DESC','BUS','BUSNAME','CKT','MW','MVAR','MVA','%I','VOLTAGE','MWLOSS','MVARLOSS','AREA','ZONE']
    xlsclns = [   'A',  'B',      'C',  'D', 'E',   'F',  'G', 'H',      'I',     'J',       'K',   'L',   'M']
    xlsclnsdict = {}
    for i in range(len(rowvars)):
        xlsclnsdict[rowvars[i]] = xlsclns[i]

    nclns = len(rowvars)

    return nclns, rowvars, xlsclnsdict

# ----------------------------------------------------------------------------------------------------
def initdict(rowvars):
    nullrowvarsdict = {}
    for each in rowvars:
        nullrowvarsdict[each] = ''
    return nullrowvarsdict

# ----------------------------------------------------------------------------------------------------
def splitstring_commaspace(tmpstr):
    '''Split string first at comma and then by space. Example:
    Input  tmpstr = a1       a2,  ,a4 a5 ,,,a8,a9
    Output strlst = ['a1', 'a2', ' ', 'a4', 'a5', ' ', ' ', 'a8', 'a9']
    '''
    strlst = []
    commalst = tmpstr.split(',')
    for each in commalst:
        eachlst = each.split()
        if eachlst:
            strlst.extend(eachlst)
        else:
            strlst.extend(' ')

    return strlst

# -----------------------------------------------------------------------------------------------------

def get_output_dir(outpath):
    # if called from PSSE's Example Folder, create report in subfolder 'Output_Pyscript'

    if outpath:
        outdir = outpath
        if not os.path.exists(outdir): os.mkdir(outdir)
    else:
        outdir = os.getcwd()
        cwd = outdir.lower()
        i = cwd.find('pti')
        j = cwd.find('psse')
        k = cwd.find('example')
        if i>0 and j>i and k>j:     # called from Example folder
            outdir = os.path.join(outdir, 'Output_Pyscript')
            if not os.path.exists(outdir): os.mkdir(outdir)

    return outdir

# -----------------------------------------------------------------------------------------------------

def get_output_filename(outpath, fnam):

    p, nx = os.path.split(fnam)
    if p:
        retvfile = fnam
    else:
        outdir = get_output_dir(outpath)
        retvfile = os.path.join(outdir, fnam)

    return retvfile

# ----------------------------------------------------------------------------------------------------

def pout_excel(savfile='savnw.sav', outpath=None, show=True):
    '''Exports power flow results to Excel Spreadsheet.
    When 'savfile' is not provided, it uses Network Data from PSS(R)E memory.
    When 'xlsfile' is provided and exists, power flow results are saved in 'next sheet#' of 'xlsfile'.
    When 'xlsfile' is provided and does not exists, power flow results are saved in 'Sheet1' of 'xlsfile'.
    When 'xlsfile' is not provided, power flow results are saved in 'Sheet1' of 'Book#.xls' file.
    '''

    import psspy

    psspy.psseinit()

    if savfile:
        ierr = psspy.case(savfile)
        if ierr != 0: return
        fpath, fext = os.path.splitext(savfile)
        if not fext: savfile = fpath + '.sav'
        #ierr = psspy.fnsl([0,0,0,1,1,0,0,0])
        #if ierr != 0: return
    else:   # saved case file not provided, check if working case is in memory
        ierr, nbuses = psspy.abuscount(-1,2)
        if ierr != 0:
            print '\n No working case in memory.'
            print ' Either provide a Saved case file name or open Saved case in PSS(R)E.'
            return
        savfile, snapfile = psspy.sfiles()

    # ================================================================================================
    # PART 1: Get the required results data
    # ================================================================================================

    # Select what to report
    if psspy.bsysisdef(0):
        sid = 0
    else:   # Select subsytem with all buses
        sid = -1

    flag_bus     = 1    # in-service
    flag_plant   = 1    # in-service
    flag_load    = 1    # in-service
    flag_swsh    = 1    # in-service
    flag_brflow  = 1    # in-service
    owner_brflow = 1    # bus, ignored if sid is -ve
    ties_brflow  = 5    # ignored if sid is -ve

    # ------------------------------------------------------------------------------------------------
    # Case Title
    titleline1, titleline2 = psspy.titldt()

    # ------------------------------------------------------------------------------------------------
    # Bus Data
    # Bus Data - Integer
    istrings = ['number','type','area','zone','owner','dummy']
    ierr, idata = psspy.abusint(sid, flag_bus, istrings)
    if ierr:
        print '(1) psspy.abusint error = %d' % ierr
        return
    ibuses = array2dict(istrings, idata)
    # Bus Data - Real
    rstrings = ['base','pu','kv','angle','angled','mismatch','o_mismatch']
    ierr, rdata = psspy.abusreal(sid, flag_bus, rstrings)
    if ierr:
        print '(1) psspy.abusreal error = %d' % ierr
        return
    rbuses = array2dict(rstrings, rdata)
    # Bus Data - Complex
    xstrings = ['voltage','shuntact','o_shuntact','shuntnom','o_shuntnom','mismatch','o_mismatch']
    ierr, xdata = psspy.abuscplx(sid, flag_bus, xstrings)
    if ierr:
        print '(1) psspy.abuscplx error = %d' % ierr
        return
    xbuses = array2dict(xstrings, xdata)
    # Bus Data - Character
    cstrings = ['name','exname']
    ierr, cdata = psspy.abuschar(sid, flag_bus, cstrings)
    if ierr:
        print '(1) psspy.abuschar error = %d' % ierr
        return
    cbuses = array2dict(cstrings, cdata)

    # Store bus data for all buses
    ibusesall={};rbusesall={};xbusesall={};cbusesall={};
    if sid == -1:
        ibusesall=ibuses
        rbusesall=rbuses
        xbusesall=xbuses
        cbusesall=cbuses
    else:
        ierr, idata = psspy.abusint(-1, flag_bus, istrings)
        if ierr:
            print '(2) psspy.abusint error = %d' % ierr
            return
        ibusesall = array2dict(istrings, idata)

        ierr, rdata = psspy.abusreal(-1, flag_bus, rstrings)
        if ierr:
            print '(2) psspy.abusreal error = %d' % ierr
            return
        rbusesall = array2dict(rstrings, rdata)

        ierr, xdata = psspy.abuscplx(-1, flag_bus, xstrings)
        if ierr:
            print '(2) psspy.abuscplx error = %d' % ierr
            return
        xbusesall = array2dict(xstrings, xdata)

        ierr, cdata = psspy.abuschar(-1, flag_bus, cstrings)
        if ierr:
            print '(2) psspy.abuschar error = %d' % ierr
            return
        cbusesall = array2dict(cstrings, cdata)

    # ------------------------------------------------------------------------------------------------
    # Plant Bus Data
    # Plant Bus Data - Integer
    istrings = ['number','type','area','zone','owner','dummy', 'status','ireg']
    ierr, idata = psspy.agenbusint(sid, flag_plant, istrings)
    if ierr:
        print 'psspy.agenbusint error = %d' % ierr
        return
    iplants = array2dict(istrings, idata)
    # Plant Bus Data - Real
    rstrings = ['base','pu','kv','angle','angled','iregbase','iregpu','iregkv','vspu','vskv','rmpct',
                'pgen',  'qgen',  'mva', 'percent', 'pmax',  'pmin',  'qmax',  'qmin',  'mismatch',
                'o_pgen','o_qgen','o_mva','o_pmax','o_pmin','o_qmax','o_qmin','o_mismatch']
    ierr, rdata = psspy.agenbusreal(sid, flag_plant, rstrings)
    if ierr:
        print 'psspy.agenbusreal error = %d' % ierr
        return
    rplants = array2dict(rstrings, rdata)
    # Plant Bus Data - Complex
    xstrings = ['voltage','pqgen','mismatch','o_pqgen','o_mismatch']
    ierr, xdata = psspy.agenbuscplx(sid, flag_plant, xstrings)
    if ierr:
        print 'psspy.agenbusreal error = %d' % ierr
        return
    xplants = array2dict(xstrings, xdata)
    # Plant Bus Data - Character
    cstrings = ['name','exname','iregname','iregexname']
    ierr, cdata = psspy.agenbuschar(sid, flag_plant, cstrings)
    if ierr:
        print 'psspy.agenbuschar error = %d' % ierr
        return
    cplants = array2dict(cstrings, cdata)

    # ------------------------------------------------------------------------------------------------
    # Load Data - based on Individual Loads Zone/Area/Owner subsystem
    # Load Data - Integer
    istrings = ['number','area','zone','owner','status']
    ierr, idata = psspy.aloadint(sid, flag_load, istrings)
    if ierr:
        print 'psspy.aloadint error = %d' % ierr
        return
    iloads = array2dict(istrings, idata)
    # Load Data - Real
    rstrings = ['mvaact','mvanom','ilact','ilnom','ylact','ylnom','totalact','totalnom','o_mvaact',
                'o_mvanom','o_ilact','o_ilnom','o_ylact','o_ylnom','o_totalact','o_totalnom']
    ierr, rdata = psspy.aloadreal(sid, flag_load, rstrings)
    if ierr:
        print 'psspy.aloadreal error = %d' % ierr
        return
    rloads = array2dict(rstrings, rdata)
    # Load Data - Complex
    xstrings = rstrings
    ierr, xdata = psspy.aloadcplx(sid, flag_load, xstrings)
    if ierr:
        print 'psspy.aloadcplx error = %d' % ierr
        return
    xloads = array2dict(xstrings, xdata)
    # Load Data - Character
    cstrings = ['id','name','exname']
    ierr, cdata = psspy.aloadchar(sid, flag_load, cstrings)
    if ierr:
        print 'psspy.aloadchar error = %d' % ierr
        return
    cloads = array2dict(cstrings, cdata)

    # ------------------------------------------------------------------------------------------------
    # Total load on a bus
    totalmva={}; totalil={}; totalyl={}; totalys={}; totalysw={}; totalload={}; busmsm={}
    for b in ibuses['number']:
        ierr, ctmva = psspy.busdt2(b,'MVA','ACT')
        if ierr==0: totalmva[b]=ctmva

        ierr, ctil = psspy.busdt2(b,'IL','ACT')
        if ierr==0: totalil[b]=ctil

        ierr, ctyl = psspy.busdt2(b,'YL','ACT')
        if ierr==0: totalyl[b]=ctyl

        ierr, ctys = psspy.busdt2(b,'YS','ACT')
        if ierr==0: totalys[b]=ctys

        ierr, ctysw = psspy.busdt2(b,'YSW','ACT')
        if ierr==0: totalysw[b]=ctysw

        ierr, ctld = psspy.busdt2(b,'TOTAL','ACT')
        if ierr==0: totalload[b]=ctld

        #Bus mismstch
        ierr, msm = psspy.busmsm(b)
        if ierr != 1: busmsm[b]=msm

    # ------------------------------------------------------------------------------------------------
    # Switched Shunt Data
    # Switched Shunt Data - Integer
    istrings = ['number','type','area','zone','owner','dummy','mode','ireg','blocks',
                'stepsblock1','stepsblock2','stepsblock3','stepsblock4','stepsblock5',
                'stepsblock6','stepsblock7','stepsblock8']
    ierr, idata = psspy.aswshint(sid, flag_swsh, istrings)
    if ierr:
        print 'psspy.aswshint error = %d' % ierr
        return
    iswsh = array2dict(istrings, idata)
    # Switched Shunt Data - Real (Note: Maximum allowed NSTR are 50. So they are split into 2)
    rstrings = ['base','pu','kv','angle','angled','vswhi','vswlo','rmpct','bswnom','bswmax',
                'bswmin','bswact','bstpblock1','bstpblock2','bstpblock3','bstpblock4','bstpblock5',
                'bstpblock6','bstpblock7','bstpblock8','mismatch']
    rstrings1 = ['o_bswnom','o_bswmax','o_bswmin','o_bswact','o_bstpblock1',
                 'o_bstpblock2','o_bstpblock3','o_bstpblock4','o_bstpblock5','o_bstpblock6',
                 'o_bstpblock7','o_bstpblock8','o_mismatch']
    ierr, rdata = psspy.aswshreal(sid, flag_swsh, rstrings)
    if ierr:
        print '(1) psspy.aswshreal error = %d' % ierr
        return
    rswsh = array2dict(rstrings, rdata)
    ierr, rdata1 = psspy.aswshreal(sid, flag_swsh, rstrings1)
    if ierr:
        print '(2) psspy.aswshreal error = %d' % ierr
        return
    rswsh1 = array2dict(rstrings1, rdata1)
    for k, v in rswsh1.iteritems():
        rswsh[k]=v
    # Switched Shunt Data - Complex
    xstrings = ['voltage','yswact','mismatch','o_yswact','o_mismatch']
    ierr, xdata = psspy.aswshcplx(sid, flag_swsh, xstrings)
    if ierr:
        print 'psspy.aswshcplx error = %d' % ierr
        return
    xswsh = array2dict(xstrings, xdata)
    # Switched Shunt Data - Character
    cstrings = ['vscname','name','exname','iregname','iregexname']
    ierr, cdata = psspy.aswshchar(sid, flag_swsh, cstrings)
    if ierr:
        print 'psspy.aswshchar error = %d' % ierr
        return
    cswsh = array2dict(cstrings, cdata)

    # ------------------------------------------------------------------------------------------------
    # Branch Flow Data
    # Branch Flow Data - Integer
    istrings = ['fromnumber','tonumber','status','nmeternumber','owners','own1','own2','own3','own4']
    ierr, idata = psspy.aflowint(sid, owner_brflow, ties_brflow, flag_brflow, istrings)
    if ierr:
        print 'psspy.aflowint error = %d' % ierr
        return
    iflow = array2dict(istrings, idata)
    # Branch Flow Data - Real
    rstrings = ['amps','pucur','pctrate','pctratea','pctrateb','pctratec','pctmvarate',
                'pctmvaratea','pctmvarateb',#'pctmvaratec','fract1','fract2','fract3',
                'fract4','rate','ratea','rateb','ratec',
                'p','q','mva','ploss','qloss',
                'o_p','o_q','o_mva','o_ploss','o_qloss'
                ]
    ierr, rdata = psspy.aflowreal(sid, owner_brflow, ties_brflow, flag_brflow, rstrings)
    if ierr:
        print 'psspy.aflowreal error = %d' % ierr
        return
    rflow = array2dict(rstrings, rdata)
    # Branch Flow Data - Complex
    xstrings = ['pq','pqloss','o_pq','o_pqloss']
    ierr, xdata = psspy.aflowcplx(sid, owner_brflow, ties_brflow, flag_brflow, xstrings)
    if ierr:
        print 'psspy.aflowcplx error = %d' % ierr
        return
    xflow = array2dict(xstrings, xdata)
    # Branch Flow Data - Character
    cstrings = ['id','fromname','fromexname','toname','toexname','nmetername','nmeterexname']
    ierr, cdata = psspy.aflowchar(sid, owner_brflow, ties_brflow, flag_brflow, cstrings)
    if ierr:
        print 'psspy.aflowchar error = %d' % ierr
        return
    cflow = array2dict(cstrings, cdata)

    # ================================================================================================
    # PART 2: Export acquired results to Excel
    # ================================================================================================
    p, nx = os.path.split(savfile)
    n, x  = os.path.splitext(nx)
    # Require path otherwise Excel stores file in My Documents directory
    xlsfile = get_output_filename(outpath, 'pout_'+n+'.xlsx')

    if os.path.exists(xlsfile):
        xlsfileExists = True
    else:
        xlsfileExists = False

    # Excel Specifications, Worksheet Size: 65,536 rows by 256 columns
    # Limit maximum data that can be exported to meet above Worksheet Size.
    maxrows, maxcols = 65530, 256
        # if required, validate number of rows and columns against these values

    # Start Excel, add a new workbook, fill it with acquired data
    xlApp = win32com.client.Dispatch("Excel.Application")

    # DisplayAlerts = True is important in order to save changed data.
    # DisplayAlerts = False suppresses all POP-UP windows, like File Overwrite Yes/No/Cancel.
    xlApp.DisplayAlerts = False

    # set this to True if want see Excel file, False if just want to save
    xlApp.Visible = show

    if xlsfileExists:   # file exist, open it and add worksheet
        xlApp.Workbooks.Open(xlsfile)
        xlBook = xlApp.ActiveWorkbook
        xlSheet = xlBook.Worksheets.Add()
    else: # file does not exist, add workbook and select sheet (=1, default)
        xlApp.Workbooks.Add()
        xlBook = xlApp.ActiveWorkbook
        xlSheet = xlBook.ActiveSheet
        try:
            xlBook.Sheets("Sheet2").Delete()
            xlBook.Sheets("Sheet3").Delete()
        except:
            pass

    # Format Excel Sheet
    xlSheet.Columns.WrapText  = False
    xlSheet.Columns.Font.Name = 'Courier New'
    xlSheet.Columns.Font.Size = 10

    nclns, rowvars, xlsclnsdict = exportedvalues()

    xlSheet.Columns(eval('"'+xlsclnsdict['DESC']    +':'+xlsclnsdict['BUS']     +'"')).ColumnWidth = 6
    xlSheet.Columns(eval('"'+xlsclnsdict['BUSNAME'] +':'+xlsclnsdict['BUSNAME'] +'"')).ColumnWidth = 18
    xlSheet.Columns(eval('"'+xlsclnsdict['CKT']     +':'+xlsclnsdict['CKT']     +'"')).ColumnWidth = 3
    xlSheet.Columns(eval('"'+xlsclnsdict['MW']      +':'+xlsclnsdict['MVA']     +'"')).ColumnWidth = 10
    xlSheet.Columns(eval('"'+xlsclnsdict['%I']      +':'+xlsclnsdict['%I']      +'"')).ColumnWidth = 6
    xlSheet.Columns(eval('"'+xlsclnsdict['VOLTAGE'] +':'+xlsclnsdict['MVARLOSS']+'"')).ColumnWidth = 10
    xlSheet.Columns(eval('"'+xlsclnsdict['AREA']    +':'+xlsclnsdict['ZONE']    +'"')).ColumnWidth = 4

    xlSheet.Columns(eval('"'+xlsclnsdict['MW']      +':'+xlsclnsdict['MVA']     +'"')).NumberFormat = "0.00"
    xlSheet.Columns(eval('"'+xlsclnsdict['%I']      +':'+xlsclnsdict['%I']      +'"')).NumberFormat = "0.00"
    xlSheet.Columns(eval('"'+xlsclnsdict['VOLTAGE'] +':'+xlsclnsdict['MVARLOSS']+'"')).NumberFormat = "0.00"

    xlSheet.Columns(eval('"'+xlsclnsdict['CKT']     +':'+xlsclnsdict['CKT']     +'"')).HorizontalAlignment = -4108
    xlSheet.Columns(eval('"'+xlsclnsdict['AREA']    +':'+xlsclnsdict['ZONE']    +'"')).HorizontalAlignment = -4108
        # Integer value -4108 is for setting alignment to "center"

    # Page steup
    xlSheet.PageSetup.Orientation  = 2      #1: Portrait, 2:landscape
    xlSheet.PageSetup.LeftMargin   = xlApp.InchesToPoints(0.5)
    xlSheet.PageSetup.RightMargin  = xlApp.InchesToPoints(0.5)
    xlSheet.PageSetup.TopMargin    = xlApp.InchesToPoints(0.25)
    xlSheet.PageSetup.BottomMargin = xlApp.InchesToPoints(0.5)
    xlSheet.PageSetup.HeaderMargin = xlApp.InchesToPoints(0.25)
    xlSheet.PageSetup.FooterMargin = xlApp.InchesToPoints(0.25)

    # ColorIndex Constants
    # BLACK       --> ColorIndex = 1
    # WHITE       --> ColorIndex = 2
    # RED         --> ColorIndex = 3
    # GREEN       --> ColorIndex = 4
    # BLUE        --> ColorIndex = 5
    # PURPLE      --> ColorIndex = 7
    # LIGHT GREEN --> ColorIndex = 43

    # ------------------------------------------------------------------------------------------------
    # Report Title
    colstart = 1
    row = 1
    col = colstart
    xlSheet.Cells(row,col).Value = "POWER FLOW OUTPUT REPORT"
    xlSheet.Cells(row,col).Font.Bold  = True
    xlSheet.Cells(row,col).Font.Size  = 14
    xlSheet.Cells(row,col).Font.ColorIndex = 7

    row += 1
    xlSheet.Cells(row,col).Value = savfile

    row += 1
    xlSheet.Cells(row,col).Value = titleline1

    row += 1
    xlSheet.Cells(row,col).Value = titleline2

    row += 2
    tr, lc, br, rc = row, 1, row, nclns #toprow, leftcolumn, bottomrow, rightcolumn
    xlSheet.Range(xlSheet.Cells(tr,lc+1),xlSheet.Cells(br,rc)).Value = rowvars[1:]
    xlSheet.Range(xlSheet.Cells(tr,lc),xlSheet.Cells(br,rc)).Font.Bold  = True
    xlSheet.Range(xlSheet.Cells(tr,lc),xlSheet.Cells(br,rc)).Font.ColorIndex = 3
    xlSheet.Range(xlSheet.Cells(tr,lc),xlSheet.Cells(br,rc)).VerticalAlignment = -4108
    xlSheet.Range(xlSheet.Cells(tr,lc),xlSheet.Cells(br,rc)).HorizontalAlignment = -4108

    clnlabelrow = row
    row += 1    # add blank row after lables

    # Worksheet Headers and Footer
    # Put Title and ColumnHeads on top of each page
    rows2repeat="$"+str(1)+":$"+str(row)
    xlSheet.PageSetup.PrintTitleRows = rows2repeat

    xlSheet.PageSetup.LeftFooter = "PF Results: " + savfile
    xlSheet.PageSetup.RightFooter = "&P of &N"

    # ------------------------------------------------------------------------------------------------
    for i, bus in enumerate(ibuses['number']):

        # select bus and put bus data in a row

        rd             = initdict(rowvars)
        rd['BUS']      = bus
        rd['BUSNAME']  = cbuses['exname'][i]
        rd['VOLTAGE']  = rbuses['pu'][i]
        rd['AREA']     = ibuses['area'][i]
        rd['ZONE']     = ibuses['zone'][i]

        row += 1
        rowvalues = [rd[each] for each in rowvars]
        xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value      = rowvalues
        xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Font.Bold  = True
        xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Font.ColorIndex = 5

        # check generation on selected bus
        plantbusidxes=busindexes(bus,iplants['number'])

        for idx in plantbusidxes:
            pcti = rplants['percent'][idx]
            if pcti == 0.0: pcti = ''
            rd             = initdict(rowvars)
            rd['DESC']     = 'FROM'
            rd['BUSNAME']  = 'GENERATION'
            rd['MW']       = rplants['pgen'][idx]
            rd['MVAR']     = rplants['qgen'][idx]
            rd['MVA']      = rplants['mva'][idx]
            rd['%I']       = pcti
            rd['VOLTAGE']  = rplants['kv'][idx]

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        # check total load on selected bus
        if bus in totalmva:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-PQ'
            rd['MW']       = totalmva[bus].real
            rd['MVAR']     = totalmva[bus].imag
            rd['MVA']      = abs(totalmva[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        if bus in totalil:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-I'
            rd['MW']       = totalil[bus].real
            rd['MVAR']     = totalil[bus].imag
            rd['MVA']      = abs(totalil[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        if bus in totalyl:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-Y'
            rd['MW']       = totalyl[bus].real
            rd['MVAR']     = totalyl[bus].imag
            rd['MVA']      = abs(totalyl[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        '''
        if bus in totalload:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-TOTAL'
            rd['MW']       = totalload[bus].real
            rd['MVAR']     = totalload[bus].imag
            rd['MVA']      = abs(totalload[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues
        '''

        if bus in totalys:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'SHUNT'
            rd['MW']       = totalys[bus].real
            rd['MVAR']     = totalys[bus].imag
            rd['MVA']      = abs(totalys[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        if bus in totalysw:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'SWITCHED SHUNT'
            rd['MW']       = totalysw[bus].real
            rd['MVAR']     = totalysw[bus].imag
            rd['MVA']      = abs(totalysw[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        """
        # Sometimes load/shunt/switch shunt area/owner/zone's could be different than the bus
        # to which it is connected. So when producing subsystem based reports, these equipment
        # might get excluded.

        # check loads on selected bus
        loadbusidxes=busindexes(bus,iloads['number'])
        pq_p = 0; pq_q = 0
        il_p = 0; il_q = 0
        yl_p = 0; yl_q = 0
        for idx in loadbusidxes:
            pq_p += xloads['mvaact'][idx].real
            pq_q += xloads['mvaact'][idx].imag
            il_p += xloads['ilact'][idx].real
            il_q += xloads['ilact'][idx].imag
            yl_p += xloads['ylact'][idx].real
            yl_q += xloads['ylact'][idx].imag

        pq_mva = abs(complex(pq_p,pq_q))
        il_mva = abs(complex(il_p,il_q))
        yl_mva = abs(complex(yl_p,yl_q))

        if pq_mva:  #PQ Loads
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-PQ'
            rd['MW']       = pq_p
            rd['MVAR']     = pq_q
            rd['MVA']      = pq_mva

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        if il_mva:   #I Loads
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-I'
            rd['MW']       = il_p
            rd['MVAR']     = il_q
            rd['MVA']      = il_mva

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        if yl_mva:   #Y Loads
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'LOAD-Y'
            rd['MW']       = yl_p
            rd['MVAR']     = yl_q
            rd['MVA']      = yl_mva

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        # check shunts on selected bus
        if abs(xbuses['shuntact'][i]):
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'SHUNT'
            rd['MW']       = xbuses['shuntact'][i].real
            rd['MVAR']     = xbuses['shuntact'][i].imag
            rd['MVA']      = abs(xbuses['shuntact'][i])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        # check switched shunts on selected bus
        swshbusidxes=busindexes(bus,iswsh['number'])
        pswsh = 0; qswsh = 0
        for idx in swshbusidxes:
            pswsh += xswsh['yswact'][idx].real
            qswsh += xswsh['yswact'][idx].imag
        mvaswsh = abs(complex(pswsh,qswsh))
        if mvaswsh:
            rd             = initdict(rowvars)
            rd['DESC']     = 'TO'
            rd['BUSNAME']  = 'SWITCHED SHUNT'
            rd['MW']       = pswsh
            rd['MVAR']     = qswsh
            rd['MVA']      = mvaswsh

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues
        """

        # check connected branches to selected bus
        flowfrombusidxes=busindexes(bus,iflow['fromnumber'])
        for idx in flowfrombusidxes:
            if iflow['tonumber'][idx]<10000000: #don't process 3-wdg xmer star-point buses
                tobusidx=busindexes(iflow['tonumber'][idx],ibusesall['number'])
                tobusVpu=rbusesall['pu'][tobusidx[0]]
                tobusarea=ibusesall['area'][tobusidx[0]]
                tobuszone=ibusesall['zone'][tobusidx[0]]
                pcti = rflow['pctrate'][idx]
                if pcti == 0.0: pcti = ''

                rd             = initdict(rowvars)
                rd['DESC']     = 'TO'
                rd['BUS']      = iflow['tonumber'][idx]
                rd['BUSNAME']  = cflow['toexname'][idx]
                rd['CKT']      = cflow['id'][idx]
                rd['MW']       = rflow['p'][idx]
                rd['MVAR']     = rflow['q'][idx]
                rd['MVA']      = rflow['mva'][idx]
                rd['%I']       = pcti
                rd['VOLTAGE']  = tobusVpu
                rd['MWLOSS']   = rflow['ploss'][idx]
                rd['MVARLOSS'] = rflow['qloss'][idx]
                rd['AREA']     = tobusarea
                rd['ZONE']     = tobuszone

                row += 1
                rowvalues = [rd[each] for each in rowvars]
                xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues

        # Bus Mismatch
        if bus in busmsm:
            rd             = initdict(rowvars)
            rd['BUSNAME']  = 'BUS MISMATCH'
            rd['MW']       = busmsm[bus].real
            rd['MVAR']     = busmsm[bus].imag
            rd['MVA']      = abs(busmsm[bus])

            row += 1
            rowvalues = [rd[each] for each in rowvars]
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Value = rowvalues
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Font.ColorIndex = 10
            xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Font.Bold = True

            # xlEdgeTop border set to xlThin
            xlSheet.Range(xlSheet.Cells(row,col+4),xlSheet.Cells(row,col+6)).Borders(8).Weight = 2

        # Insert EdgeBottom border(Borders(9)) with xlThin weight(2)
        xlSheet.Range(xlSheet.Cells(row,col),xlSheet.Cells(row,nclns)).Borders(9).Weight = 2

        # Insert blank row
        row += 1

    # ------------------------------------------------------------------------------------------------
    # Draw borders
    # Column Lable Row
    # xlEdgeTop border set to xlThin
    xlSheet.Range(xlSheet.Cells(clnlabelrow,1),xlSheet.Cells(clnlabelrow,nclns)).Borders(8).Weight = 2
    # xlEdgeBottom border set to xlThin
    xlSheet.Range(xlSheet.Cells(clnlabelrow,2),xlSheet.Cells(clnlabelrow,nclns)).Borders(9).Weight = 2

    # Remaining WorkSheet
    # xlEdgeLeft border set to xlThinline
    xlSheet.Range(xlSheet.Cells(clnlabelrow,1),xlSheet.Cells(row-1,nclns)).Borders(7).Weight = 2
    # xlEdgeRight border set to xlThinline
    xlSheet.Range(xlSheet.Cells(clnlabelrow,1),xlSheet.Cells(row-1,nclns)).Borders(10).Weight = 2
    # xlInsideVertical border set to xlHairline
    xlSheet.Range(xlSheet.Cells(clnlabelrow,1),xlSheet.Cells(row-1,nclns)).Borders(11).Weight = 1

    # ------------------------------------------------------------------------------------------------
    # Save the workbook and close the Excel application

    if xlsfile:  # xls file provided
        xlBook.SaveAs(Filename = xlsfile)
    else:
        xlsbookfilename = os.path.join(os.getcwd(),xlBook.Name)  # xlBook.Name returns without '.xls'
        xlBook.SaveAs(Filename = xlsbookfilename)
        xlsbookfilename = os.path.join(os.getcwd(),xlBook.Name)  # xlBook.Name returns '.xls' extn.

    if not show:
        xlBook.Close()
        xlApp.Quit()
        txt = '\n Power Flow Results saved to file %s' % xlsfile
        sys.stdout.write(txt)

# ====================================================================================================
# ====================================================================================================
if __name__ == '__main__':
    import os
    import sys
    sys.path.append(r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27');
    os.environ['PATH']+=';'+r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27';  #or where else you find the psspy.pyc
    sys.path.append(r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN');
    os.environ['PATH']+=';'+r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN';
    import psspy
    import redirect#psse34
    #pout_excel()
    # OR
    pout_excel(savfile='savnw.sav',show=True)

# ====================================================================================================
