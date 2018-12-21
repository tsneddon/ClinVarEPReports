from ftplib import FTP
import datetime
import time
import os
import sys
import csv
import gzip
import re
import pprint
import xlsxwriter

scvHash = {}
a2vHash = {}
HGVSHash = {}
EPHash = {}
EPList = []
geneHash = {}
geneList = []
today = datetime.datetime.today().strftime('%Y%m%d') #todays date YYYYMMDD


def get_file(file):
    '''This function gets ClinVar files from FTP'''

    domain = 'ftp.ncbi.nih.gov'
    path = '/pub/clinvar/tab_delimited/'
    user = 'anonymous'
    password = 'tsneddon@broadinstitute.org'

    ftp = FTP(domain)
    ftp.login(user, password)
    ftp.cwd(path)
    localfile = open(file, 'wb')
    ftp.retrbinary('RETR ' + file, localfile.write)
    raw_date = ftp.sendcmd('MDTM ' + file)
    date = datetime.datetime.strptime(raw_date[4:], "%Y%m%d%H%M%S").strftime("%m-%d-%Y")
    ftp.quit()
    localfile.close()

    return(date)


def make_directory(dir, date):
    '''This function makes a local directory for new files if directory does not already exist'''

    directory = dir + '/EP_Reports_' + date

    if not os.path.exists(directory):
        os.makedirs(directory)
    else:
        sys.exit('Program terminated, ' + directory + ' already exists.')

    return(directory)


def convert_date(date):
    '''This function converts a ClinVar date eg May 02, 2018 -> YYYYMMDD'''

    mon2num = dict(Jan='01', Feb='02', Mar='03', Apr='04', May='05', Jun='06',\
                   Jul='07', Aug='08', Sep='09', Oct='10', Nov='11', Dec='12')

    if '-' not in date:
        newDate = re.split(', | ',date)
        newMonth = mon2num[newDate[0]]
        convertDate = (newDate[2] + newMonth + newDate[1]) #YYYYMMDD, an integer for date comparisons
    else:
        convertDate = date

    return(convertDate)


def print_date(date):
    '''This function converts a date eg YYYYMMDD -> MM/DD/YYYY'''

    printDate = date[4:6] + "/" + date[6:8] + "/" + date[0:4] #MM/DD/YYYY, for printing to file
    return(printDate)


def create_scvHash(gzfile):
    '''This function makes a hash of each SCV in each VarID'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    varID = int(col[0])
                    clinSig = col[1]
                    rawDate = col[2]
                    dateLastEval = convert_date(rawDate) #convert date eg May 02, 2018 -> YYYYMMDD
                    revStat = col[6]

                    submitter = col[9]
                    submitter = re.sub(r'\s+', '_', submitter) #replace all spaces with an underscore
                    submitter = re.sub(r'/', '-', submitter) # replace all slashes with a hyphen
                    submitter = re.sub(r'\W+', '', submitter) #remove all non-alphanumerics

                    SCV = col[10]

                    if revStat == 'reviewed by expert panel' and 'PharmGKB' not in submitter: #-- to exclude PharmGKB records
                        EPHash[varID] = {'ClinSig':clinSig, 'Submitter':submitter, 'DateLastEval':dateLastEval}

                        if submitter not in EPList:
                            EPList.append(submitter)

                    else:
                        if varID not in scvHash.keys():
                            scvHash[varID] = {}

                        scvHash[varID][SCV] = {'ClinSig':clinSig, 'DateLastEval':dateLastEval, 'Submitter':submitter, 'ReviewStatus':revStat}

    #Add VCEPs that are not yet approved and have no variants in ClinVar
    EPList.append('Monogenic_Diabetes')

    input.close()
    os.remove(gzfile)
    return(scvHash, EPHash, EPList)


def create_a2vHash(gzfile):
    '''This function makes a dictionary of VarID to AlleleID'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()
            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    varID = int(col[0])
                    type = col[1]
                    alleleID = int(col[2])

                    #Ignore rows that are not Variant (simple type)
                    #This excludes Haplotype, CompoundHeterozygote, Complex, Phase unknown, Distinct chromosomes
                    if type == 'Variant':
                        a2vHash[alleleID] = varID

    input.close()
    os.remove(gzfile)
    return(a2vHash)


def create_HGVSHash(gzfile):
    '''This function makes a hash of metadata for each VarID'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    alleleID = int(col[0])
                    type = col[1]
                    HGVSname = col[2]
                    geneSym = col[4]
                    phenotype = col[13]
                    guidelines = col[26]

                    if alleleID in a2vHash:
                        HGVSHash[a2vHash[alleleID]] = {'VarType':type, 'HGVSname':HGVSname, 'GeneSym':geneSym,'Phenotype':phenotype,'Guidelines':guidelines}

    input.close()
    os.remove(gzfile)
    return(HGVSHash)


def create_geneList(file):
    '''This function creates a GeneList for the genes in scope for each EP'''

    with open(file, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()
            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    geneSym = col[0]
                    geneSym = re.sub(r'\s+', '', geneSym)
                    EPname = col[1]
                    EPname = re.sub(r'\W+', '', EPname) #remove all non-alphanumerics
                    geneHash[geneSym] = EPname
                    geneList.append(geneSym)

    input.close()
    return(geneList, geneHash)


def create_EPfiles(ExcelDir, excelFile, date):
    '''This function creates an Excel file for each EP in the EPList'''

    dir = ExcelDir

    for EP in EPList:
        EP_output_file = dir + '/' + EP + '_' + excelFile

        workbook = xlsxwriter.Workbook(EP_output_file)
        worksheet0 = workbook.add_worksheet('README')

        worksheet0.write(0, 0, "Date of ClinVar FTP file: " + date)
        worksheet0.write(2, 0, "Expert Panel: " + EP)
        worksheet0.write(4, 0, "This Excel file is the output of a script that takes the most recent submission_summary.txt file from the ClinVar FTP site and outputs all the variants that need updating/reviewing by the " + EP + " (EP).")
        worksheet0.write(5, 0, 'Each tab is the result of a different set of parameters as outlined below:')
        worksheet0.write(6, 0, '#Variants:')
        worksheet0.write(7, 1, '1. Alert: ClinVar variants with an LP/VUS Expert Panel SCV with a DateLastEvaluated > 2 years from the date of this file (may overlap with variants on Tabs 2, 3 and 4).')
        worksheet0.write(8, 1, '2. Alert: ClinVar variants with a P/LP Expert Panel SCV AND a newer VUS/LB/B non-EP SCV (with a DateLastEvaluated up to 1 year prior of EP DateLastEvaluated).')
        worksheet0.write(9, 1, '3. Alert: ClinVar variants with a VUS Expert Panel SCV AND a newer P/LP non-EP SCV (with a DateLastEvaluated up to 1 year prior of EP DateLastEvaluated).')
        worksheet0.write(10, 1, '4. Alert: ClinVar variants with a VUS Expert Panel SCV AND a newer LB/B non-EP SCV (with a DateLastEvaluated up to 1 year prior of EP DateLastEvaluated).')
        worksheet0.write(11, 1, '5. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with at least one P/LP SCV and at least one VUS/LB/B SCV (medically-significant conflict).')
        worksheet0.write(12, 1, '6. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with at least one VUS SCV and at least one LB/B SCV.')
        worksheet0.write(13, 1, '7. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with >=3 concordant VUS SCVs from different submitters.')
        worksheet0.write(14, 1, '8. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with at least one P/LP SCV from (at best) a no assertion criteria provided submitter.')

        tabList = [create_tab1, create_tab2, create_tab3, create_tab4, create_tab5, create_tab6, create_tab7, create_tab8]

        for tab in tabList:
            tab(EP, workbook, worksheet0)

        workbook.close()


def create_tab1(EP, workbook, worksheet0):
    '''This function creates the Tab#1 Alert (EP_OutOfDate) in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet1 = workbook.add_worksheet('1.Alert_EP_OutOfDate')

    for varID in EPHash:
        if EPHash[varID]['Submitter'] == EP and EPHash[varID]['DateLastEval'] != '-' and \
           int(EPHash[varID]['DateLastEval']) < (int(today) - 30000) and \
           (EPHash[varID]['ClinSig'] == 'Likely pathogenic' or EPHash[varID]['ClinSig'] == 'Uncertain significance'):
            p2fileVarIDs.append(varID)
            if varID in scvHash:
                for SCV in scvHash[varID]:
                    headerSubs.append(scvHash[varID][SCV]['Submitter'])

    headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet1, 5, 'Other')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', No DLE)')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet1, row, varID, 5, headerSubs, varSubs, i)

    print_stats(worksheet0, 7, 0, row)


def create_tab2(EP, workbook, worksheet0):
    '''This function creates the Tab#2 Alert (EP_PLPvsNewSub_VUSLBB) in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet2 = workbook.add_worksheet('2.Alert_EP_PLPvsNewSub_VUSLBB')

    for varID in EPHash:
        if varID in EPHash and EPHash[varID]['Submitter'] == EP and EPHash[varID]['DateLastEval'] != '-':
            if varID in scvHash:
                for SCV in scvHash[varID]:
                    if scvHash[varID][SCV]['DateLastEval'] != '-' and \
                    (int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) or \
                    int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) - 10000) and \
                    ((EPHash[varID]['ClinSig'] == 'Pathogenic' or EPHash[varID]['ClinSig'] == 'Likely pathogenic') and \
                    (scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance' or scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or \
                    scvHash[varID][SCV]['ClinSig'] == 'Benign')):
                        headerSubs.append(scvHash[varID][SCV]['Submitter'])
                        if varID not in p2fileVarIDs:
                            p2fileVarIDs.append(varID)

    headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet2, 5, 'Newer')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-' and \
                   (int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) or \
                   int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) - 10000) and \
                   (scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance' or scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or \
                   scvHash[varID][SCV]['ClinSig'] == 'Benign'):
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet2, row, varID, 5, headerSubs, varSubs, i)

    print_stats(worksheet0, 8, 0, row)

def create_tab3(EP, workbook, worksheet0):
    '''This function creates the Tab#3 Alert (EP_VUSvsNewSub_PLP) in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet3 = workbook.add_worksheet('3.Alert_EP_VUSvsNewSub_PLP')

    for varID in EPHash:
        if varID in EPHash and EPHash[varID]['Submitter'] == EP and EPHash[varID]['DateLastEval'] != '-':
            if varID in scvHash:
                for SCV in scvHash[varID]:
                    if scvHash[varID][SCV]['DateLastEval'] != '-' and \
                    (int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) or \
                    int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) - 10000) and \
                    ((EPHash[varID]['ClinSig'] == 'Uncertain significance') and \
                    (scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic' or scvHash[varID][SCV]['ClinSig'] == 'Pathogenic')):
                        headerSubs.append(scvHash[varID][SCV]['Submitter'])
                        if varID not in p2fileVarIDs:
                            p2fileVarIDs.append(varID)

    headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet3, 5, 'Newer')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-' and \
                   (int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) or \
                   int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) - 10000) and \
                   ((EPHash[varID]['ClinSig'] == 'Uncertain significance') and \
                   (scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic' or scvHash[varID][SCV]['ClinSig'] == 'Pathogenic')):
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet3, row, varID, 5, headerSubs, varSubs, i)

    print_stats(worksheet0, 9, 0, row)


def create_tab4(EP, workbook, worksheet0):
    '''This function creates the Tab#4 Alert (EP_VUSvsNewSub_LBB) in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet4 = workbook.add_worksheet('4.Alert_EP_VUSvsNewSub_LBB')

    for varID in EPHash:
        if varID in EPHash and EPHash[varID]['Submitter'] == EP and EPHash[varID]['DateLastEval'] != '-':
            if varID in scvHash:
                for SCV in scvHash[varID]:
                    if scvHash[varID][SCV]['DateLastEval'] != '-' and \
                    (int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) or \
                    int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) - 10000) and \
                    ((EPHash[varID]['ClinSig'] == 'Uncertain significance') and \
                    (scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or scvHash[varID][SCV]['ClinSig'] == 'Benign')):
                        headerSubs.append(scvHash[varID][SCV]['Submitter'])
                        if varID not in p2fileVarIDs:
                            p2fileVarIDs.append(varID)

    headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet4, 5, 'Newer')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-' and \
                   (int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) or \
                   int(scvHash[varID][SCV]['DateLastEval']) > int(EPHash[varID]['DateLastEval']) - 10000) and \
                   ((EPHash[varID]['ClinSig'] == 'Uncertain significance') and \
                   (scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or scvHash[varID][SCV]['ClinSig'] == 'Benign')):
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet4, row, varID, 5, headerSubs, varSubs, i)

    print_stats(worksheet0, 10, 0, row)


def create_tab5(EP, workbook, worksheet0):
    '''This function creates the Tab#5 Priority (PLPvsVUSLBB)in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet5 = workbook.add_worksheet('5.Priority_PLPvsVUSLBB')

    for varID in scvHash:
        submitters = []
        ClinSigList = []

        if varID not in EPHash:
            for SCV in scvHash[varID]:
                submitters.append(scvHash[varID][SCV]['Submitter'])
                ClinSigList.append(scvHash[varID][SCV]['ClinSig'])

            submitters = sorted(set(submitters))
            ClinSigList = sorted(set(ClinSigList))

            if ('Pathogenic' in ClinSigList or 'Likely pathogenic' in ClinSigList) \
                and ('Uncertain significance' in ClinSigList or 'Likely benign' in ClinSigList or 'Benign' in ClinSigList) \
                and varID in HGVSHash and HGVSHash[varID]['GeneSym'] in geneList \
                and geneHash[HGVSHash[varID]['GeneSym']] in EP:
                if varID not in p2fileVarIDs:
                    p2fileVarIDs.append(varID)
                if submitters:
                    headerSubs.extend(submitters)

            headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet5, 4, 'Conflicting')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', No DLE)')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet5, row, varID, 4, headerSubs, varSubs, i)

    print_stats(worksheet0, 11, 0, row)


def create_tab6(EP, workbook, worksheet0):
    '''This function creates the Tab#6 Priority (VUSvsLBB)in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet6 = workbook.add_worksheet('6.Priority_VUSvsLBB')

    for varID in scvHash:
        submitters = []
        ClinSigList = []

        if varID not in EPHash:
            for SCV in scvHash[varID]:
                submitters.append(scvHash[varID][SCV]['Submitter'])
                ClinSigList.append(scvHash[varID][SCV]['ClinSig'])

            submitters = sorted(set(submitters))
            ClinSigList = sorted(set(ClinSigList))

            if ('Uncertain significance' in ClinSigList) \
                and ('Likely benign' in ClinSigList or 'Benign' in ClinSigList) \
                and varID in HGVSHash and HGVSHash[varID]['GeneSym'] in geneList \
                and geneHash[HGVSHash[varID]['GeneSym']] in EP:
                if varID not in p2fileVarIDs:
                    p2fileVarIDs.append(varID)
                if submitters:
                    headerSubs.extend(submitters)

            headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet6, 4, 'Conflicting')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', No DLE)')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet6, row, varID, 4, headerSubs, varSubs, i)

    print_stats(worksheet0, 12, 0, row)


def create_tab7(EP, workbook, worksheet0):
    '''This function creates the Tab#7 Priority (multiVUS) in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet7 = workbook.add_worksheet('7.Priority_multiVUS')

    for varID in scvHash:
        submitters = []
        ClinSigList = []
        count = 0

        if varID not in EPHash:
            unique_subs = []

            for SCV in scvHash[varID]:
                submitters.append(scvHash[varID][SCV]['Submitter'])
                ClinSigList.append(scvHash[varID][SCV]['ClinSig'])
                #Don't double count (Illumina's) duplicate submissions!!!
                current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
                if current_sub not in unique_subs \
                   and scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance':
                    unique_subs.append(current_sub)
                    count += 1

            submitters = sorted(set(submitters))
            ClinSigList = sorted(set(ClinSigList))

            if count > 2 and varID in HGVSHash\
               and 'Pathogenic' not in ClinSigList and 'Likely pathogenic' not in ClinSigList \
               and 'Likely benign' not in ClinSigList and 'Benign' not in ClinSigList \
               and HGVSHash[varID]['GeneSym'] in geneList \
               and geneHash[HGVSHash[varID]['GeneSym']] in EP:
                if varID not in p2fileVarIDs:
                    p2fileVarIDs.append(varID)
                if submitters:
                    headerSubs.extend(submitters)

            headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet7, 4, 'Consensus')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', No DLE)')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet7, row, varID, 4, headerSubs, varSubs, i)

    print_stats(worksheet0, 13, 0, row)


def create_tab8(EP, workbook, worksheet0):
    '''This function creates the Tab#8 Priority (noCriteriaPLP) in the Excel file'''

    row = 0
    p2fileVarIDs = []
    headerSubs = []

    worksheet8 = workbook.add_worksheet('8.Priority_noCriteriaPLP')

    for varID in scvHash:
        submitters = []
        ReviewStatus = []

        if varID not in EPHash:
            unique_subs = []

            for SCV in scvHash[varID]:
                submitters.append(scvHash[varID][SCV]['Submitter'])
                ReviewStatus.append(scvHash[varID][SCV]['ReviewStatus'])
                #Don't double count (Illumina's) duplicate submissions!!!
                current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
                if current_sub not in unique_subs \
                   and (scvHash[varID][SCV]['ClinSig'] == 'Pathogenic' or scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic') and \
                   scvHash[varID][SCV]['ReviewStatus'] == 'no assertion criteria provided':
                    unique_subs.append(current_sub)

            submitters = sorted(set(submitters))
            ReviewStatus = sorted(set(ReviewStatus))

            if unique_subs != [] and varID in HGVSHash\
               and 'practice guideline' not in ReviewStatus and 'criteria provided, single submitter' not in ReviewStatus \
               and HGVSHash[varID]['GeneSym'] in geneList \
               and geneHash[HGVSHash[varID]['GeneSym']] in EP:
                if varID not in p2fileVarIDs:
                    p2fileVarIDs.append(varID)
                if submitters:
                    headerSubs.extend(submitters)

            headerSubs = sorted(set(headerSubs))

    i = print_header(p2fileVarIDs, headerSubs, worksheet8, 4, 'No_assertion')

    for varID in p2fileVarIDs:
        varSubs = []
        if varID in scvHash:
            for SCV in scvHash[varID]:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> MM/DD/YYYY
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', ' + subPrintDate + ')')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' (' + scvHash[varID][SCV]['ClinSig'] + ', No DLE)')

        varSubs = sorted(set(varSubs))

        row, i = print_variants(worksheet8, row, varID, 4, headerSubs, varSubs, i)

    print_stats(worksheet0, 14, 0, row)


def print_header(p2fileVarIDs, headerSubs, worksheet, i, type):
    '''This function prints all the header titles to the Excel tabs'''

    if p2fileVarIDs != []:
        worksheet.write(0, 0, 'VarID')
        worksheet.write(0, 1, 'Gene_symbol(s)')
        worksheet.write(0, 2, 'Phenotype(s)')
        worksheet.write(0, 3, 'HGVS_name')
        worksheet.write(0, 4, 'Expert_Panel(ClinSig, DateLastEval)')

        for sub in headerSubs:
            worksheet.write(0, i, sub)
            i += 1
        worksheet.write(0, i, type + '_submitters (ClinSig, DateLastEval)')
    else:
        worksheet.write(0, 0, 'No variants found')

    return(i)


def print_variants(worksheet, row, varID, j, headerSubs, varSubs, i):
    '''This function prints all the variants to the Excel tabs'''

    row += 1
    worksheet.write(row, 0, varID)

    if HGVSHash[varID]['GeneSym']:
        worksheet.write(row, 1, HGVSHash[varID]['GeneSym'])

    if HGVSHash[varID]['Phenotype']:
        worksheet.write(row, 2, HGVSHash[varID]['Phenotype'])

    if HGVSHash[varID]['HGVSname']:
        worksheet.write(row, 3, HGVSHash[varID]['HGVSname'])

    if j == 5:
        #Convert date from YYYYMMDD -> MM/DD/YYYY
        EPPrintDate = print_date(EPHash[varID]['DateLastEval'])
        worksheet.write(row, 4, EPHash[varID]['Submitter'] + ' (' + EPHash[varID]['ClinSig'] + ', ' + EPPrintDate + ')')

    for headerSub in headerSubs:
        p2file = 'no'
        for varSub in varSubs:
            if headerSub in varSub:
                p2file = varSub[varSub.find("(")+1:varSub.find(")")]
        if p2file != 'no':
            worksheet.write(row, j, p2file)
            j += 1
        else:
            j += 1

    if varSubs:
       sublist = ' | '. join(varSubs)
       worksheet.write(row, i, sublist)

    return(row, i)


def print_stats(worksheet0, line, column, row):
    '''This function prints the total variant count to the README Excel tab'''

    worksheet0.write(line, column, row)


def main():

    inputFile1 = 'submission_summary.txt.gz'
    inputFile2 = 'variation_allele.txt.gz'
    inputFile3 = 'variant_summary.txt.gz'
    geneFile = 'EP_GeneList.txt'

    dir = 'ClinVarExpertPanelReports'

    date = get_file(inputFile1)
    get_file(inputFile2)
    get_file(inputFile3)

    ExcelDir = make_directory(dir, date)

    excelFile = 'EP_Updates_' + date + '.xlsx'

    create_scvHash(inputFile1)
    create_a2vHash(inputFile2)
    create_HGVSHash(inputFile3)
    create_geneList(geneFile)

    create_EPfiles(ExcelDir, excelFile, date)

main()
